VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form FrmPeriodo 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Proceso de Pre-Facturación"
   ClientHeight    =   3300
   ClientLeft      =   5625
   ClientTop       =   3705
   ClientWidth     =   4740
   ClipControls    =   0   'False
   Icon            =   "FrmPeriodo.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   3300
   ScaleWidth      =   4740
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Caption         =   "Periodo ( Mes - Año )"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1110
      Left            =   630
      TabIndex        =   5
      Top             =   420
      Width           =   3405
      Begin VB.ComboBox CmbPeriodo 
         DataField       =   "TipoMovimientoCaja"
         Height          =   315
         Index           =   1
         ItemData        =   "FrmPeriodo.frx":000C
         Left            =   1875
         List            =   "FrmPeriodo.frx":000E
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   435
         Width           =   1020
      End
      Begin VB.ComboBox CmbPeriodo 
         DataField       =   "TipoMovimientoCaja"
         Height          =   315
         Index           =   0
         ItemData        =   "FrmPeriodo.frx":0010
         Left            =   465
         List            =   "FrmPeriodo.frx":003B
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   435
         Width           =   1320
      End
   End
   Begin VB.CommandButton CmdSalir 
      Cancel          =   -1  'True
      Caption         =   "Salir"
      Height          =   765
      Left            =   2460
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   2265
      Width           =   1005
   End
   Begin VB.CommandButton CmdOk 
      Caption         =   "Aceptar"
      Height          =   765
      Left            =   1125
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2265
      Width           =   1005
   End
   Begin ComctlLib.ProgressBar Barra 
      Height          =   390
      Left            =   660
      TabIndex        =   1
      Top             =   1635
      Visible         =   0   'False
      Width           =   3390
      _ExtentX        =   5980
      _ExtentY        =   688
      _Version        =   327682
      Appearance      =   1
   End
End
Attribute VB_Name = "FrmPeriodo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
     '---------------------------------------------------------------------------------------------
    '29/08/2002: Módulo de prefacturación. Proceso cerrado que genera los gastos los gastos fijos-
    'Comunes, los gastos fijos no comunes y los honorarios profesionales. Emite el Pre'Recibo de -
    'facturación. Necesarios parámetros de prefacturación y selección del usuario del periodo-----
    'Dim WrkPreFactura As Workspace      'Espacio de Trabajo para la facturación
    Dim cnnInmueble As ADODB.Connection 'Conexión publica a nivel de módulo al inmueble selec.
    Dim datPeriodo As Date                      'Periodo de pre - facturado
    Dim intAptos As Integer          '# de apartamentos en el inmueble
    
    Private Sub CmbPeriodo_KeyPress(Index As Integer, KeyAscii As Integer)
    '
    If KeyAscii = 13 And Index = 0 Then CmbPeriodo(1).SetFocus
    If KeyAscii = 13 And Index = 1 Then CmdOk.SetFocus
    '
    End Sub

    '---------------------------------------------------------------------------------------------
    Private Sub CmdOK_Click() '-Rutina Proceso Pre - Facturación
    '---------------------------------------------------------------------------------------------
    'Variables Locales
   
    Dim AdoFactura As ADODB.Recordset
    Dim CodRAP$, CodRLP$, strFE$, CodIDB$, tempFecha$, Pos%, codGA$, CodIVA$
    Dim Periodo As Date, booGA As Boolean
    Dim AdoInmueble As ADODB.Recordset          'Información general del Inmueble
    Dim curLimite@, curHono@                  'Limite Facturación (cálculo fondo Especial)
    Dim lngFalta As Long, E&
    '
    Set AdoFactura = New ADODB.Recordset
    datPeriodo = Format(CDate("01/" & CmbPeriodo(0) & "/" & CmbPeriodo(1)), "MM-DD-YY")
    
    AdoFactura.Open "SELECT * FROM factura WHERE periodo = #" & datPeriodo & "# AND Fact not LI" _
        & "KE 'CHD%'", cnnInmueble, adOpenKeyset, adLockOptimistic
        
    If Not AdoFactura.EOF And Not AdoFactura.BOF Then
        MsgBox "Sr. Usuario : " & vbCrLf & _
        " Este periodo ya fue facturado, por favor rectifique...", _
        vbInformation, "Facturacion"
        AdoFactura.Close
        Set AdoFactura = Nothing
        Exit Sub    'Sale del proceso.......
    End If
    AdoFactura.Close
    'confirma el procesamiento del mes
    If Not Respuesta("Seguro que desea realizar el proceso para el periodo:" _
        & CmbPeriodo(0) & "/" & CmbPeriodo(1)) Then    'Sale si niega la confirmación
        Set AdoFactura = Nothing
        Exit Sub
    End If
    '
    DoEvents
    AdoFactura.Open "SELECT CodGasto FROM Tgastos WHERE Titulo Like 'REINTEGRO AGUA PRESUPUESTA" _
    & "%';", cnnInmueble, adOpenKeyset, adLockOptimistic, adCmdText
    If Not AdoFactura.EOF Or AdoFactura.BOF Then CodRAP = AdoFactura!codGasto
    AdoFactura.Close
    'Configura la presentación del Formulario
    Barra.Visible = True
    Barra.Min = 1
    Barra.Max = 120
    Barra.Value = Barra.Min
    '
    '---------------------------------------------------------------------------------------------
    Set AdoInmueble = New ADODB.Recordset
    '---------------------------------------------------------------------------------------------
    Barra.Value = 20
    '
    On Error Resume Next
    'actualiza el catalogo de gastos los gastos de cuotas
    cnnInmueble.Execute "UPDATE Tgastos SET Fijo=False,Van=0,Cuotas=0 WHERE (Van=Cuotas) AND Cuotas>0;"
    cnnInmueble.BeginTrans      'Comienza al proceso por lotes
    '---------------------------------------------------------------------------------------------
    AdoInmueble.Open "SELECT * FROM inmueble WHERE codinm = '" & gcCodInm & "'", _
        cnnConexion, adOpenKeyset, adLockOptimistic
    '---------------------------------------------------------------------------------------------
    Barra.Value = 40
    '
    If AdoInmueble.EOF Then
        MsgBox ("No encontré Código del Inmueble....")
        AdoInmueble.Close
        Set AdoInmueble = Nothing
        cnnInmueble.RollbackTrans
        Exit Sub
    End If
    If IsNull(AdoInmueble("Unidad")) Or AdoInmueble("Unidad") = "" Then
        lngaltaF = MsgBox("Registre la cantidad de propietarios en la ficha del inmueble", vbExclamation, App.ProductName)
    ElseIf (IsNull(AdoInmueble("CodIDB")) Or AdoInmueble("CodIDB") = "") And gnIDB > 0 Then
        lngFalta = MsgBox("Verifique la ficha 'Administrativos' en la Ficha del inmueble" & vbCrLf _
        & "Falta Código del IDB", vbExclamation, App.ProductName)
    ElseIf (IsNull(AdoInmueble("CodIva")) Or AdoInmueble("CodIVa") = "") And gnIva > 0 Then
        lngFalta = MsgBox("Verifique la ficha 'Administrativos' en la ficha del inmueble" & vbCrLf & _
        "Falta Código del IVA", vbExclamation, App.ProductName)
    End If
    If lngFalta <> 0 Then
        AdoInmueble.Close
        Set AdoInmueble = Nothing
        cnnInmueble.RollbackTrans
        Exit Sub
    End If
    '---------------------------------------------------------------------------------------------
    'Asigna valores a las variables temporales
    With AdoInmueble
        '
        intAptos = !Unidad
        curLimite = !Facturacion
        strFE = IIf(IsNull(!CodFondoE), "", !CodFondoE) '
        CodIDB = !CodIDB
        CodIVA = !CodIVA
        If !FactGA Then
            booGA = True
            codGA = !CodGastoAdmin
            curHono = !Honorarios
        End If
        .Close
        Set AdoInmueble = Nothing
    End With
    MousePointer = vbHourglass
    Call Print_PreRecibo(datPeriodo)
    '---------------------------------------------------------------------------------------------
    Barra.Value = 60
    'actualiza el catálogo de gastos
    'gastos fijos con monto cera los marca como no fijos
    cnnInmueble.Execute "UPDATE TGastos SET Fijo=0 WHERE Fijo=True and MontoFijo=0"
    'Elimina Gastos Fijos comunes generados por este proceso-----------------------------------
    cnnInmueble.Execute "DELETE * FROM AsignaGasto WHERE (((AsignaGasto.Cargado)=#" _
    & datPeriodo & "#) AND ((AsignaGasto.Ndoc)='0' Or (AsignaGasto.Ndoc) Is Null)) OR (((AsignaG" _
    & "asto.Cargado)=#" & datPeriodo & "#) AND ((AsignaGasto.CodGasto)='" & gcCodFondo & "'));"
    '
    Barra.Value = 70
    'Elimina Gastos No Comunes generados por este proceso--------------------------------------
    cnnInmueble.Execute "DELETE * FROM GastoNocomun WHERE PF=True AND Periodo=#" & datPeriodo & "#"
    '
    Barra.Value = 80
    'Agrega GastosFijos No Comunes----------------------------------------------------------------
    Call asigna_GNC
    '
     Barra.Value = 100
    'Agrega Gastos Fijos Comunes------------------------------------------------------------------
    cnnInmueble.Execute "INSERT INTO AsignaGasto (CodGasto, Descripcion, Fijo,Comun, Alicuota, " _
        & "Cargado, Monto, Ndoc, Fecha, Hora, Usuario) SELECT CodGasto, Titulo & iif(Incrementa=" _
        & "True,' ' & Van + 1 & '/' & Cuotas,''), Fijo, Comun, Al" _
        & "icuota, #" & datPeriodo & "#,MontoFijo, '0' as doc,date() as ti, time() as da ,'" _
        & gcUsuario & "' FROM Tgastos WHERE CodGasto<>'" & gcCodFondo & "' AND Fijo=True AND Co" _
        & "mun=True AND Not MontoFijo Is Null AND CodGasto NOT IN (SELECT CodGasto FROm AsignaG" _
        & "asto WHERE Cargado=#" & datPeriodo & "# AND CodGasto in (SELECT CodGasto FROM Tgasto" _
        & "s WHERE Fijo=True) AND NDOC <>'0');"
        
    'agrega el registro de gastos administrativos si es un gasto comun
    If booGA Then
        cnnInmueble.Execute "INSERT INTO AsignaGasto (CodGasto, Descripcion, Fijo,Comun, Alicuo" _
        & "ta,Cargado, Monto,Ndoc,Fecha,Hora,Usuario) SELECT CodGasto, Titulo, Fijo, Comun," _
        & "Alicuota, #" & datPeriodo & "#,'" & curHono & "', '0',date() as ti, time() ,'" & _
        gcUsuario & "' FROM Tgastos WHERE CodGasto='" & codGA & "' AND Comun=True;", E
        If E > 0 Then
        '
            If gnIva > 0 Then
                'AGREGA EL IVA
                'cnnInmueble.Execute "INSERT INTO AsignaGasto (CodGasto, Descripcion, Fijo, Comun, Alicuo" _
                & "Ta,Cargado, Monto,Ndoc,Fecha,Hora,Usuario) SELECT CodGasto,Titulo,Fijo,-1," _
                & "Alicuota,#" & datPeriodo & "#,'" & curHono * gnIva / 100 & "',0,Date(),Time(),'" & _
                gcUsuario & "' FROM TGastos WHERE CodGasto='" & CodIVA & "'"
                cnnInmueble.Execute "UPDATE AsignaGasto SET Monto = Monto * (1 + '" & _
                gnIva / 100 & "') WHERE CodGasto='" & codGA & "' AND Cargado=#" & _
                datPeriodo & "#;"
            End If
            '
        End If
        
    End If
    'Reintegra los servicios LUZ,Agua,Telefono
    'cnnInmueble.Properties(51) = "5231"
    AdoFactura.Open "SELECT * FROM AsignaGasto WHERE CodGasto IN (SELECT CodGasto FROM Servicio" _
    & "s IN '" & gcPath & "\sac.mdb' WHERE Inmueble='" & gcCodInm & "' AND Tiposerv=0) AND Carg" _
    & "ado=#" & datPeriodo & "#;", cnnInmueble, adOpenStatic, adLockReadOnly, adCmdText
    '
    If Not AdoFactura.EOF Or Not AdoFactura.BOF Then    'si tiene registros
    
        AdoFactura.MoveFirst    'mueve al primero
        Do  'hace hasta finalizar el archivo
            CodRLP = Mid(AdoFactura!Descripcion, _
            IIf(InStr(AdoFactura!Descripcion, "10000") = 0, 1, _
            InStr(AdoFactura!Descripcion, "10000")), 12)
            '
            If IsDate("01/" & Right(AdoFactura!Descripcion, 7)) Then
                'Periodo = Date
            'Else
                Periodo = CDate("01/" & Right(AdoFactura!Descripcion, 7))
            'End If
            '
                cnnInmueble.Execute "INSERT INTO AsignaGasto(CodGasto,Descripcion,Fijo,Comun,Al" _
                & "icuota,Cargado,Monto,Ndoc,Fecha,Hora,Usuario) SELECT Left(CodGasto,2) & 2 & " _
                & "Right(CodGasto, 3),'REINTEGRO ' &  left(Descripcion,instr(Descripcion,' Mes')) & right(Descripcion,7) ,0,Comun,Alicuota,#" & _
                datPeriodo & "#,Monto * -1,'0',Date(),Time(),'" & gcUsuario & "' FROM AsignaGas" _
                & "to WHERE Descripcion Like 'LUZ%PRES%" & CodRLP & "%" _
                & Format(Periodo, "MM/YYYY") & "';", n
            '
            Else
                MsgBox "No se hará el reintegro a: " & AdoFactura!Descripcion, vbInformation, "PreFacturación"
                Call rtnBitacora("No se hará el reintegro a: " & AdoFactura!Descripcion)
            End If
            AdoFactura.MoveNext
        Loop Until AdoFactura.EOF
        '
    End If
    AdoFactura.Close
    
    'Reintegro agua presupuestada de la factura procesada en este período
    If CodRAP = "" Then CodRAP = "202005"
    'pendiente aqui tienes que cambiar el codigo 200003 por uno que este en la BD
    'con la finalidad de hacer el programa portable
    AdoFactura.Open "SELECT * FROM AsignaGasto WHERE CodGasto='200003' AND Cargado=#" & _
    datPeriodo & "#;", cnnInmueble, adOpenStatic, adLockReadOnly, adCmdText
    
    If Not AdoFactura.EOF Or Not AdoFactura.BOF Then    'Encuentra coincidencia con el criterio
        AdoFactura.MoveFirst
        Do
        If IsDate("01/" & Right(Trim(AdoFactura!Descripcion), 7)) Then
            Periodo = CDate("01/" & Right(Trim(AdoFactura!Descripcion), 7))
1000    cnnInmueble.Execute "INSERT INTO AsignaGasto(CodGasto,Descripcion,Fijo,Comun,Alicuota," _
            & "Cargado,Monto,Ndoc,Fecha,Hora,Usuario) SELECT '" & CodRAP & "','REINTEGRO AGUA P" _
            & "RESUPUESTADA " & Format(Periodo, "MM/YYYY") & "',0,Comun,Alicuota,#" & _
            datPeriodo & "#,Monto * -1,'0',Date(),Time(),'" & gcUsuario & "' FROM AsignaGasto W" _
            & "HERE Descripcion Like 'Agua%PRES%" & Format(Periodo, "MM/YYYY") & "';", n
            '
        Else
            tempFecha = "01/" & InputBox$("Introduzca el período correspondiente al agua presupuestada que " _
            & "desea reintegrar. Con el Formato siguiente MM/AAAA. Ejemplo: Para reintegrar " _
            & "el mes de mayo del 2003, escriba: 05/2003", App.EXEName)
            If IsDate(tempFecha) Then
                Periodo = CDate(tempFecha)
                GoTo 1000
            Else
                MsgBox "No se realizó correctamente el reintegro del agua." & vbCrLf & _
                "Consulte al administrador del sistema", vbInformation, App.ProductName
            End If
            '
        End If
        AdoFactura.MoveNext
        Loop Until AdoFactura.EOF
    End If
    '
    Barra.Value = 120
    '
    'Emite Pre-Recibo de Condominio
    If Err.Number = 0 Then
        cnnInmueble.CommitTrans
        'Agrega Gasto Bancario (IDB) si está activo
        If gnIDB > 0 Then
            If CodIDB = "" Then
                CodIDB = InputBox$("Verifique en la ficha del inmueble el Código correspondient" _
                & "e al Débito Bancario. Ingrese el Código del Gasto 'Débito Bancario'", "Código" _
                & " I.D.B.", "600001")
            End If
            If CodIDB <> "" Then
                cnnInmueble.Execute "INSERT INTO AsignaGasto(Ndoc,CodGasto,Descripcion,Comun,Al" _
                & "icuota,Cargado,Monto,Fecha,Hora,Usuario) SELECT '0' as D,'" & CodIDB & "' as" _
                & " CG,'DEBITO BANCARIO' as De,True,True,#" & datPeriodo & "#, Clng(SUM(Monto)* '" _
                & gnIDB & "' /100),Date(),Time(),'" & gcUsuario & "' FROM AsignaGasto WHERE Carga" _
                & "do=#" & datPeriodo & "# AND CodGasto Not In (SELECT CodGasto FROM Fact_IDB I" _
                & "N '" & gcPath & "\sac.mdb');"
            End If
        End If
        '
        'si condominio tiene fondo especial, lo calcula
        If curLimite > 0 Then
            Dim rstFonEsp As New ADODB.Recordset
            Dim lngFonEsp As Currency
            rstFonEsp.Open "SELECT SUM(Clng(Monto)) FROM AsignaGasto Where Comun = True AND Car" _
            & "gado =#" & datPeriodo & "# AND CodGasto<>'" & gcCodFondo & "';", cnnInmueble, _
            adOpenKeyset, adLockOptimistic
            lngFonEsp = ((curLimite - CLng(rstFonEsp.Fields(0))) * 100) / (100 + gnIDB)
            lngFonEsp = CLng(lngFonEsp)
            If lngFonEsp <> 0 Then  'si se hace algún ajuste
                If strFE = "" Then
                    strFE = InputBox("Introduzaca el Código del Fondo Especial de Emergencia")
                End If
                If strFE = "" Then strFE = "700001"
                'Actualiza la tabla asignagasto
                cnnInmueble.Execute "INSERT INTO AsignaGasto (Ndoc,CodGasto,Cargado,Descripcion" _
                & ",Comun,Alicuota,Monto,Usuario,Fecha,Hora) SELECT '0' as d,'" & strFE & "',#" _
                & datPeriodo & "#, Titulo ,Comun,Alicuota,'" & lngFonEsp & "','" & gcUsuario & _
                "', date(),Time() FROM Tgastos WHERE CodGasto='" & strFE & "';"
                'actualiza el debito bancario (si se cobra)
                If gnIDB > 0 Then
                    cnnInmueble.Execute "UPDATE AsignaGasto SET Monto=Clng(Monto + (" & _
                    CLng(lngFonEsp) & " * '" & gnIDB & "' /100)) WHERE Cargado=#" & datPeriodo & _
                    "# AND Codgasto='" & CodIDB & "';"
                End If
                '
            End If
            rstFonEsp.Close
            Set rstFonEsp = Nothing
        End If
        '
        'Actualiza la descripción de los gastos presupuestados
        strFE = " " & UCase(Format(DateAdd("m", 1, Format(datPeriodo, "mm/dd/yyyy")), "mm/yyyy"))
        cnnInmueble.Execute "UPDATE AsignaGasto SET Descripcion=Descripcion + '" & strFE & "' W" _
        & "HERE (Cargado=#" & datPeriodo & "#) AND (Ndoc='0') AND (CodGasto IN (SELECT CodGasto" _
        & " FROM Fact_Presupuestados IN '" & gcPath & "\sac.mdb'));"
        '
        'Actualiza la descripcion de los gastos reintegro S.S.O
        'Actualiza la descripcion de los gastos reintegro lph
        strFE = " " & Format(datPeriodo, "dd/yyyy")
        cnnInmueble.Execute "UPDATE AsignaGasto SET Descripcion=Descripcion + '" & strFE & "' W" _
        & "HERE (Cargado=#" & datPeriodo & "#) AND (Ndoc='0') AND (CodGasto IN ('190004','19202" _
        & "0'));"
    '
        strFE = UCase(Left(CmbPeriodo(0), 3)) & "-" & CmbPeriodo(1)
        'MsgBox LoadResString(530), vbInformation, LoadResString(531)
        'Call clear_Crystal(FrmAdmin.rptReporte)
        DoEvents
        Call Printer_PreRecibo(strFE, datPeriodo, , crPantalla)
    '
    '   Agrega al Fondo de Reserva a TDF AsignaGasto-------------------------------------------------
        cnnInmueble.Execute "INSERT INTO AsignaGasto(Ndoc,CodGasto,Cargado,Descripcion,Fijo,Com" _
        & "un,Alicuota,Monto,Usuario,Fecha,Hora) SELECT '0' as D,'" & gcCodFondo & "' as CG,#" & _
        datPeriodo & "#,'" & gnPorcFondo & "% FONDO DE RESERVA " & UCase(Left(CmbPeriodo(0), 3)) _
        & "/" & CmbPeriodo(1) & "',True,True,True, CCur((Sum(CCur(Monto))*" & gnPorcFondo _
        & "/100)) AS FR,'" & gcUsuario & "',DATE(),TIME() From AsignaGasto WHERE (((Comun" _
        & ")=True)" & " AND ((AsignaGasto.Cargado)=#" & datPeriodo & "#));"
    '
    Else 'el proceso generó errores durante la ejecución
        cnnInmueble.RollbackTrans
        MsgBox "Error Durante le ejecución, Pongase en contacto con el proveedor del sistema" & _
        vbCrLf & Err.Description
        Err.Clear
    End If
    MousePointer = vbNormal
    Barra.Visible = False
    '
    
    If Err.Number <> 0 Then
        MsgBox "Se produjeron errores durante el proceso" & vbCrLf & Err.Description, vbCritical, _
        "Error " & Err.Number
        Call rtnBitacora("Se produjo el siguiente error " & Err.Number & " durante la ejecució" _
        & "n del proceso. " & Err.Description)
    Else
        Unload Me
        Set FrmPeriodo = Nothing
    End If
    '
    End Sub

    Private Sub cmdSalir_Click()
    Unload Me
    Set FrmPeriodo = Nothing
    End Sub

    'Carga la presentación del formulario en pantalla---------------------------------------------
    Private Sub Form_Load()
    '---------------------------------------------------------------------------------------------
    On Error Resume Next
    Call CenterForm(FrmPeriodo)
    CmdOk.Picture = LoadResPicture("OK", vbResIcon)
    CmdSalir.Picture = LoadResPicture("Salir", vbResIcon)
    'Crea una instancia del espacio de trabajo para este proceso
    'Set WrkPreFactura = CreateWorkspace("", "Admin", "")
    '
    Set cnnInmueble = New ADODB.Connection
    cnnInmueble.Open cnnOLEDB & mcDatos
    
    Dim strSQL  As String
    Dim ultimoPeriodoFacturado As Date
    strSQL = "SELECT MAX(Periodo) FROM Factura WHERE Fact Not Like 'CH%' or IsNull(F" _
            & "act)"
    
    ultimoPeriodoFacturado = IIf(IsNull(cnnInmueble.Execute(strSQL)(0).Value), Date, cnnInmueble.Execute(strSQL)(0).Value)
    If (utlimoperiodofacturado = Date) Then
        For I = 0 To 1: CmbPeriodo(1).AddItem (Year(Date) + I)
        Next
        If Month(Date) = 1 Then CmbPeriodo(1).AddItem (Year(Date) - 1)
         'Presenta el periodo al mes actual
        CmbPeriodo(0).Text = CmbPeriodo(0).List(Month(Date) - 1)
        CmbPeriodo(1).Text = CmbPeriodo(1).List(IIf(Month(Date) = 1, 1, 0))
    Else
        ultimoPeriodoFacturado = DateAdd("m", 1, ultimoPeriodoFacturado)
        For I = 0 To 1: CmbPeriodo(1).AddItem (Year(ultimoPeriodoFacturado) + I)
        Next
        'If Month(ultimoPeriodoFacturado) = 1 Then CmbPeriodo(1).AddItem (Year(ultimoPeriodoFacturado) - 1)
         'Presenta el periodo al mes actual
        CmbPeriodo(0).Text = CmbPeriodo(0).List(Month(ultimoPeriodoFacturado) - 1)
        CmbPeriodo(1).Text = CmbPeriodo(1).List(IIf(Month(ultimoPeriodoFacturado) = 1, 1, 0))
    End If
   
    
    '
    End Sub


    Private Sub Form_Unload(Cancel As Integer)
    cnnInmueble.Close
    Set cnnInmueble = Nothing
    'WrkPreFactura.Close
    'Set WrkPreFactura = Nothing
    End Sub

    '---------------------------------------------------------------------------------------------
    '   Rutina asigna_GNC
    '
    '   Agrega los registros a TDF GastoNoComun, de los gastos fijos no comunes
    '   calculados por alicuota o partes iguales
    '---------------------------------------------------------------------------------------------
    Private Sub asigna_GNC()
    '
    On Error Resume Next
    Dim Calc$
    '
    For I = 0 To 1
        If I = 0 Then   'Por alícuota
            Calc = "MontoFijo * T2.Alicuota/100"
            Alicuota = "True"
        Else    'en partes iguales
            Calc = "MontoFijo /" & intAptos
            Alicuota = "False"
        End If
        cnnInmueble.Execute "INSERT INTO GastoNoComun(CodApto, CodGasto, Concepto, Monto, Perio" _
        & "do, Fecha, Hora, Usuario,PF) SELECT T2.Codigo, T1.CodGasto, T1.Titulo & iif(T1.Incre" _
        & "menta=True,' ' & T1.van + 1 & '/' & T1.Cuotas,''), CCur(" & Calc _
        & ") as Total,#" & datPeriodo & "# as Per,'" & Date & "' as Fec,'" & Format(Time, "hh:m" _
        & "m ampm") & "' as hor,'" & gcUsuario & "' as usu,True as p FROM Tgastos AS T1, Propie" _
        & "tarios AS T2 WHERE (((T2.Codigo)<>'U" & gcCodInm & "') AND ((T1.Fijo)=True) AND ((T1" _
        & ".Comun)=False) AND ((T1.Alicuota)=" & Alicuota & ")) ORDER BY T2.Codigo"
        '
    Next
    '
    Err.Clear
    '
    End Sub

