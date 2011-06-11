VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form FrmAvisoCobro 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   4710
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6630
   ClipControls    =   0   'False
   FillStyle       =   0  'Solid
   Icon            =   "FrmAvisoCobro.frx":0000
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4710
   ScaleWidth      =   6630
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FraAvisoCobro 
      Height          =   2070
      Index           =   2
      Left            =   4860
      TabIndex        =   13
      Top             =   2400
      Width           =   1530
      Begin VB.CommandButton cmdAviso 
         Caption         =   "Im&primir"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   800
         Index           =   0
         Left            =   225
         MouseIcon       =   "FrmAvisoCobro.frx":5C12
         MousePointer    =   99  'Custom
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   285
         Width           =   1095
      End
      Begin VB.CommandButton cmdAviso 
         Caption         =   "&Salir"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   800
         Index           =   1
         Left            =   210
         MouseIcon       =   "FrmAvisoCobro.frx":5F1C
         MousePointer    =   99  'Custom
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   1155
         Width           =   1095
      End
   End
   Begin VB.Frame FraAvisoCobro 
      Caption         =   "Selección Inmueble"
      Height          =   2070
      Index           =   0
      Left            =   200
      TabIndex        =   0
      Top             =   120
      Width           =   6180
      Begin VB.OptionButton optImp 
         Caption         =   "Correo electrónico"
         Height          =   270
         Index           =   2
         Left            =   3570
         TabIndex        =   29
         Top             =   1747
         Width           =   1890
      End
      Begin VB.OptionButton optImp 
         Caption         =   "Impresora"
         Height          =   270
         Index           =   1
         Left            =   945
         TabIndex        =   28
         Top             =   1747
         Width           =   1215
      End
      Begin VB.OptionButton optImp 
         Caption         =   "Ventana"
         Height          =   270
         Index           =   0
         Left            =   2295
         TabIndex        =   27
         Top             =   1747
         Value           =   -1  'True
         Width           =   1215
      End
      Begin MSDataListLib.DataCombo DtcAviso 
         Height          =   315
         Index           =   0
         Left            =   1170
         TabIndex        =   2
         Top             =   480
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   556
         _Version        =   393216
         ListField       =   "CodInm"
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo DtcAviso 
         Height          =   315
         Index           =   1
         Left            =   2265
         TabIndex        =   3
         Top             =   480
         Width           =   3765
         _ExtentX        =   6641
         _ExtentY        =   556
         _Version        =   393216
         ListField       =   "Nombre"
         Text            =   ""
      End
      Begin VB.Label Label1 
         Caption         =   "Salida:"
         Height          =   285
         Index           =   7
         Left            =   210
         TabIndex        =   26
         Top             =   1740
         Width           =   900
      End
      Begin VB.Label Label1 
         Height          =   285
         Index           =   5
         Left            =   1170
         TabIndex        =   15
         Top             =   1365
         Width           =   4155
      End
      Begin VB.Label Label1 
         Caption         =   "Ubicación:"
         Height          =   285
         Index           =   4
         Left            =   195
         TabIndex        =   14
         Top             =   1365
         Width           =   900
      End
      Begin VB.Label Label1 
         Height          =   285
         Index           =   3
         Left            =   1170
         TabIndex        =   5
         Top             =   945
         Width           =   4155
      End
      Begin VB.Label Label1 
         Caption         =   "Impresora:"
         Height          =   285
         Index           =   2
         Left            =   195
         TabIndex        =   4
         Top             =   945
         Width           =   900
      End
      Begin VB.Label Label1 
         Caption         =   "&Inmueble:"
         Height          =   285
         Index           =   0
         Left            =   200
         TabIndex        =   1
         Top             =   495
         Width           =   900
      End
   End
   Begin VB.Frame FraAvisoCobro 
      Caption         =   "Carta para propietarios:"
      Height          =   2070
      Index           =   1
      Left            =   180
      TabIndex        =   6
      Top             =   2415
      Width           =   4515
      Begin VB.TextBox TxtIntervalo 
         Enabled         =   0   'False
         Height          =   315
         Index           =   3
         Left            =   2415
         TabIndex        =   25
         Top             =   1050
         Width           =   1800
      End
      Begin VB.CheckBox ChkAviso 
         Caption         =   "&Emisión Telegramas"
         Height          =   285
         Index           =   4
         Left            =   135
         TabIndex        =   24
         Top             =   1050
         Width           =   2055
      End
      Begin VB.CheckBox ChkAviso 
         Caption         =   "Carta &más de tres meses"
         Height          =   285
         Index           =   1
         Left            =   135
         TabIndex        =   8
         Top             =   660
         Width           =   2055
      End
      Begin VB.CheckBox ChkAviso 
         Caption         =   "Carta &Tres meses vencidos"
         Height          =   285
         Index           =   0
         Left            =   135
         TabIndex        =   7
         Top             =   270
         Width           =   2745
      End
      Begin VB.TextBox TxtIntervalo 
         Enabled         =   0   'False
         Height          =   315
         Index           =   0
         Left            =   2415
         TabIndex        =   10
         Top             =   645
         Width           =   1800
      End
      Begin VB.Label Label1 
         Caption         =   "Escriba el número de recibos vencidos e intervalos separados por coma Ejm: 4,6,8-10,15"
         Height          =   510
         Index           =   1
         Left            =   195
         TabIndex        =   9
         Top             =   1515
         Width           =   4125
      End
   End
   Begin VB.Frame FraAvisoCobro 
      Caption         =   "Estado de Cuenta"
      Height          =   2070
      Index           =   3
      Left            =   180
      TabIndex        =   16
      Top             =   2415
      Width           =   4515
      Begin VB.OptionButton Opcion 
         Caption         =   "&D&escendente"
         Height          =   315
         Index           =   1
         Left            =   1980
         TabIndex        =   23
         Top             =   1230
         Value           =   -1  'True
         Width           =   1300
      End
      Begin VB.OptionButton Opcion 
         Caption         =   "&Ascendente"
         Height          =   315
         Index           =   0
         Left            =   585
         TabIndex        =   22
         Top             =   1230
         Width           =   1215
      End
      Begin VB.TextBox TxtIntervalo 
         Height          =   315
         Index           =   2
         Left            =   2760
         TabIndex        =   21
         Top             =   855
         Width           =   600
      End
      Begin VB.CheckBox ChkAviso 
         Caption         =   "Deuda &General Inmueble"
         Height          =   495
         Index           =   2
         Left            =   200
         MouseIcon       =   "FrmAvisoCobro.frx":6226
         Picture         =   "FrmAvisoCobro.frx":6668
         TabIndex        =   19
         Top             =   765
         Width           =   2520
      End
      Begin VB.CheckBox ChkAviso 
         Caption         =   "&Confidencial Junta"
         Height          =   495
         Index           =   3
         Left            =   200
         MouseIcon       =   "FrmAvisoCobro.frx":6972
         TabIndex        =   18
         Top             =   270
         Width           =   2100
      End
      Begin VB.TextBox TxtIntervalo 
         Height          =   315
         Index           =   1
         Left            =   2760
         TabIndex        =   17
         Top             =   360
         Width           =   600
      End
      Begin VB.Label Label1 
         Caption         =   "Escriba el número de recibos vencidos a partir de los cuales desea la información."
         Height          =   360
         Index           =   6
         Left            =   195
         TabIndex        =   20
         Top             =   1590
         Width           =   4125
      End
   End
End
Attribute VB_Name = "FrmAvisoCobro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
    'Option Explicit
    Dim objRst As ADODB.Recordset
    Dim StrRutaInmueble$
    Dim booAfected As Boolean
    Dim IntCarta As Integer, intMM As Integer, intMora As Integer
    Const sCARTA$ = "Carta"
    Const sTELEGRAMA$ = "Telegrama"
    Dim WithEvents poSendMail As clsSendMail
Attribute poSendMail.VB_VarHelpID = -1
    
    '
    Private Sub ChkAviso_Click(Index As Integer)
    'variables locales
    Select Case Index
        Case 1, 4
        I = IIf(Index = 1, 0, 3)
        With TxtIntervalo(I)
            .Enabled = Not .Enabled
            .Text = ""
            If .Enabled Then .SetFocus
        End With
    End Select
    '
    End Sub
'
    Private Sub ChkAviso_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then ChkAviso(Index).Value = IIf(ChkAviso(Index).Value = 0, 1, 0)
    End Sub

    Private Sub cmdAviso_Click(Index As Integer)
    'variables locales
    Dim errLocal&   'variables locales
    Dim IntVeces%
    Dim strFS$
    Dim rptLocal As ctlReport
    '
    Select Case Index   'SELECCIONA UN BOTON
        '
        Case 0  'PULSO IMPRIMIR
            '
            If StrRutaInmueble = "" Then MsgBox "Debe Seleccionar un Inmueble..", _
                vbCritical, App.ProductName: DtcAviso(0).SetFocus: Exit Sub
            booAfected = False
            
            If Me.Caption = "Impresión Avisos de Cobro" Then
            If optImp(2) Then
                'envia las notificaciones via correo electrónico
                Call Enviar_Correo
            Else
                '   SELECCIONA LOS REGISTROS DE ACUERDO A LA SELECCION DEL AVISO
                '
                    If ChkAviso(0).Value = 0 And ChkAviso(1) = 0 And ChkAviso(4) = 0 Then _
                        MsgBox "Debe Seleccionar el Tipo de Aviso de Cobro", _
                        vbInformation, App.ProductName: Exit Sub
                    
                    'Activado Recordatorio de Cobro
                    If ChkAviso(0).Value = 1 Then Call ftnMarcar_Aviso(sCARTA, "Recibos=3 AND Convenio=" _
                    & "False and Demanda=False", "Aviso11.rpt", "{Propietarios.Recibos}= 3" & _
                    " AND {Propietarios.Convenio}=False", "Carta 3 Meses")
                    '----------------
                    'Activado la Notificación
                    If ChkAviso(1).Value = 1 Then strFS = ftnFS(0, "Aviso2.rpt", "Carta más de 4 meses.")
                    '----------------
                    'Activados los telegramas
                    If ChkAviso(4) = 1 Then strFS = ftnFS(3, "telegrama1.rpt", "Telegramas")
                    '----------------
                    If booAfected Then
                        If Respuesta("Desea el reporte general?") Then
                            'With FrmAdmin.rptReporte
                            Set rptLocal = New ctlReport
                            With rptLocal
                                'Call clear_Crystal(FrmAdmin.rptReporte) 'limpia el control
                                '.ReportFileName = gcReport + "avisos_report.rpt"
                                .Reporte = gcReport + "avisos_report.rpt"
                                '.DataFiles(0) = gcPath & StrRutaInmueble & "inm.mdb"
                                .OrigenDatos(0) = gcPath & StrRutaInmueble & "inm.mdb"
                                .Formulas(0) = "Inmueble='" & DtcAviso(1) & "'"
                                '.Destination = Salida
                                .Salida = Salida
                                .TituloVentana = "Reporte General"
                                'errlocal = .PrintReport
                                .Imprimir
                                'If errlocal <> 0 Then MsgBox .LastErrorString, vbCritical, .LastErrorNumber
                            End With
                            Set rptLocal = Nothing
                        End If
                    End If
            End If
            '
            Else    'IMRIME LOS ESTADOS DE CUENTA
                
                IntVeces = 0
                
                For I = 2 To 3  '2=Deuda Genereal y 3=Deuda Confidencial
                
                If ChkAviso(I).Value = 1 Then
                    
                    Set rptLocal = New ctlReport
                    With rptLocal
                        .Reporte = gcReport & IIf(I = 3, "EdoCtaRes.rpt", "EdoCtaInm.rpt")
                        .OrigenDatos(0) = gcPath + StrRutaInmueble + "inm.mdb"
                        .OrigenDatos(1) = gcPath + StrRutaInmueble + "inm.mdb"
                        '
                        If I = 2 Then
                            If Opcion(0).Value = True Then
                                .FormuladeSeleccion = IIf(TxtIntervalo(2) = "", "{Factura.Saldo}<>0", _
                                    "{Factura.Saldo}<>0 and {Propietarios.Recibos}>=" & TxtIntervalo(2))
                            ElseIf Opcion(1).Value = True Then
                                .FormuladeSeleccion = IIf(TxtIntervalo(2) = "", "{Factura.Saldo}<>0", _
                                    "{Factura.Saldo}<>0 and {Propietarios.Recibos}<=" & TxtIntervalo(2))
                            End If
                        End If
                        If I = 3 Then
                            .FormuladeSeleccion = IIf(TxtIntervalo(1) = "", _
                            "", "{Propietarios.Recibos}>=" & TxtIntervalo(1))
                        End If
                        .Formulas(0) = "Inmueble ='" & DtcAviso(1) & "'"
                        .TituloVentana = "Edo Cta. " & DtcAviso(0)
                        .Salida = Salida
                        .Imprimir
                    End With
                    Set rptLocal = Nothing
                    IntVeces = IntVeces + 1
                End If
                Next
                '
                If IntVeces = 0 Then
                    MsgBox "Debe Marcar por lo menos un tipo de estado de cuenta,....", vbCritical, _
                    App.ProductName
                End If
                '
            End If
        
        Case 1  'PRESIONO SALIR
            Unload Me
            Set FrmAvisoCobro = Nothing
    End Select
    
    End Sub

    '20/08/2002-----------------------------------------------------------------------------------
    Private Sub DtcAviso_Click(Index As Integer, Area As Integer)
    '---------------------------------------------------------------------------------------------
    If Area = 2 Then    'SI SELECCIONA UN ELEMENTO DE LA LISTA
        Select Case Index
    '
            Case 0  'BUSCA POR CODIGO DE INMUEBLE
                Call RtnBusqueda("CodInm = '" & DtcAviso(0).Text & "'")
            
            Case 1  'BUSCA POR NOMBRE DEL INMUEBLE
                Call RtnBusqueda("Nombre like '%" & DtcAviso(1).Text & "%'")
    '
        End Select
    '
    End If
    '
    End Sub

    '20/08/2002-----------------------------------------------------------------------------------
    Private Sub DtcAviso_KeyPress(Index As Integer, KeyAscii As Integer)
    '---------------------------------------------------------------------------------------------
    'variables locales
    KeyAscii = Asc(UCase(Chr(KeyAscii)))   'CONVIERTE A MAYUSCULAR
    If KeyAscii = 13 Then   'SI PRESIONA ENTER
    '
        Select Case Index
            'BUSQUEDA POR CODIGO DE INMUEBLE
            Case 0: Call RtnBusqueda("CodInm = '" & DtcAviso(0).Text & "'")
            'BUSQUEDA POR NOMBRE DEL INMUEBLE
            Case 1: Call RtnBusqueda("Nombre like '%" & DtcAviso(1).Text & "%'")
    '
        End Select
    '
    End If
    '
    End Sub

    'rev.22/08/2002------------------------------carga el Formulario Avisos de Cobro--------------
    Private Sub Form_Load()
    '---------------------------------------------------------------------------------------------
    cmdAviso(0).Picture = LoadResPicture("Print", vbResIcon)
    cmdAviso(1).Picture = LoadResPicture("Salir", vbResIcon)
    'ESTABLECE LA PROPIEDAD CAPTION DE LAS ETIQUETAS
    Label1(3).Caption = Printer.DeviceName 'IMPRESORA
    Label1(5) = Printer.Port    'UBICACION IMPRESORA
    Set objRst = New ADODB.Recordset
    objRst.Open "SELECT * FROM Inmueble ORDER BY CodInm", cnnConexion, adOpenStatic, _
    adLockReadOnly
    Set DtcAviso(0).RowSource = objRst  'ESTABLECE EL ORIGEN DE LA LISTA
    Set objRst = New ADODB.Recordset
    objRst.Open "SELECT * FROM Inmueble ORDER BY Nombre", cnnConexion, adOpenStatic, _
    adLockReadOnly
    Set DtcAviso(1).RowSource = objRst  'ESTABLECE EL ORIGEN DE LA LISTA
    Set poSendMail = New clsSendMail
    
    '
    End Sub

    '20/08/2002-------Rutina que busca información sobre el inmueble seleccionado----------------
    Private Sub RtnBusqueda(StrExpresion As String)
    '--------------------------------------------------------------------------------------------
    'variables locales
    '
    With objRst
    '
        .MoveFirst
        .Find StrExpresion
        If .EOF Then
            MsgBox "Inmueble No Registrado", vbCritical, App.ProductName
            Exit Sub
        End If
        For I = 0 To 1
            DtcAviso(I) = .Fields(I)
        Next
        StrRutaInmueble = .Fields("Ubica")       'Carpeta del Inmueble
        intMM = .Fields("MesesMora")               'Constante Meses de Mora
        intMora = .Fields("HonoMorosidad")      'Procentaje de Honorarios
        '
    End With
    '
    End Sub

    '20/08/2002----Valida el campo a solo números, coma y guion-----------------------------------
    Private Sub TxtIntervalo_KeyPress(Index As Integer, KeyAscii As Integer)
    '---------------------------------------------------------------------------------------------
    'variables locales
    Call Validacion(KeyAscii, "0123456789-,")
    '
    End Sub

    '20/08/2002--------Imprime el reporte según parametros recibios-------------------------------
    Private Sub RtnPrintReport(StrReport$, StrSeleccion$, strTitulo$)
    '---------------------------------------------------------------------------------------------
    Dim rpReporte As ctlReport
    
    '
    Set rpReporte = New ctlReport
    With rpReporte
    '
        .Reporte = gcReport + StrReport
        .OrigenDatos(0) = gcPath + StrRutaInmueble + "inm.mdb"
        If (UCase(StrReport) <> "AVISO11.RPT" And UCase(StrReport) <> "TELEGRAMA1.RPT") Then .OrigenDatos(1) = gcPath + "\sac.mdb"
        .FormuladeSeleccion = StrSeleccion
        .Formulas(0) = "Fecha = '" & Format(Date, "Long date") & "'"
        .Formulas(1) = "Inmueble = '" & DtcAviso(1) & "'"
        .Formulas(2) = "Total =" & IIf(StrReport = "Aviso2.rpt", "({Propietarios.Deuda}*" _
        & intMora & "/100) + {Propietarios.Deuda}", "")
        .TituloVentana = strTitulo
        .Salida = Salida
        .Imprimir
        If Err <> 0 Then MsgBox Err.Description, vbCritical, Err
    '
    End With
    '
    End Sub

    '21/08/2002-----------------------------------------------------------------------------------
    '   funcion:    ftnFS
    '
    '
    '---------------------------------------------------------------------------------------------
    Private Function ftnFS(I%, r$, T$) As String
    '
    Dim strCadena$, strRango$, strDesde$, strHasta$, strCondicion$
    Dim bytComa As Byte, bytGuion As Byte
    '
    If TxtIntervalo(I) = "" Then    'Valida que el usuario especifique parámetros
        MsgBox "Debe especificar parámetros de impresión de las Notificaciones", _
            vbInformation, "Faltan Meses Vencidos"
            ftnFS = ""
        Exit Function
    End If
    '   ------------------------------------------------------------------------------------------
    '   Determina los parámetros de busqueda
    '   ---------------------------------------------------------------------------------------------
    TxtIntervalo(I) = _
    IIf(Right(TxtIntervalo(I), 1) = Chr(44) Or Right(TxtIntervalo(I), 1) = _
        Chr(45), Left(TxtIntervalo(I), Len(TxtIntervalo(I)) - 1), TxtIntervalo(I))
    strCadena = TxtIntervalo(I)
    While Len(strCadena) > 0
        bytComa = InStr(1, strCadena, Chr(44))
        If bytComa = 0 Then
            strRango = strCadena
            strCadena = ""
        Else
            strRango = Left(strCadena, bytComa - 1)
            strCadena = Right(strCadena, Len(strCadena) - bytComa)
        End If
        bytGuion = InStr(1, strRango, Chr(45))
        If bytGuion = 0 Then
            ftnFS = ftnFS & IIf(ftnFS = "", _
                "{Propietarios.Recibos}=" & strRango, " or {Propietarios.Recibos}=" _
                    & strRango)
            strCondicion = strCondicion & IIf(strCondicion = "", "Recibos=" & strRango, _
                " or Recibos=" & strRango)
        Else
            strDesde = "{Propietarios.Recibos}>=" & Left(strRango, bytGuion - 1)
            strHasta = "{Propietarios.Recibos}<=" _
                & Right(strRango, Len(strRango) - bytGuion)
            ftnFS = ftnFS & IIf(ftnFS = "", strDesde, _
                " or " & strDesde) & " AND " & strHasta
        
            strDesde = "Recibos>=" & Left(strRango, bytGuion - 1)
            strHasta = "Recibos<=" & Right(strRango, Len(strRango) - bytGuion)
            strCondicion = strCondicion & IIf(strCondicion = "", strDesde, _
                " or " & strDesde) & " AND " & strHasta
        End If
    Wend
    ftnFS = "(" & ftnFS & ") AND {Propietarios.Convenio}=False and {Propietarios.Demanda}=False"
    If strCondicion <> "" Then
        strCondicion = "(" & strCondicion & ") AND convenio=False and Demanda=False"
        If optImp(2) Then
            ftnFS = strCondicion
        Else
            If I = 0 Then   'Avisos de cobro
                Call ftnMarcar_Aviso(sCARTA, strCondicion, r, ftnFS, T)
            Else
                Call ftnMarcar_Aviso(sTELEGRAMA, strCondicion, r, ftnFS, T)
            End If
        End If
        '
    End If
    
    '
    End Function
    
    '---------------------------------------------------------------------------------------------
    '   Función: rtnMarcar
    '
    '   Entrada:    strCond,strCamp
    '
    '   Marca enviado el documento segun las condiciones recibidas en la
    '   variable strCond, devuelve la cantidad de registros actualizados
    '---------------------------------------------------------------------------------------------
    Private Sub ftnMarcar_Aviso(strCamp$, strCond$, Report$, Seleccion$, Titulo$)
    'variable locales
    Dim cnnInmueble As New ADODB.Connection
    Dim nReg&
    '
    cnnInmueble.Open cnnOLEDB & gcPath & StrRutaInmueble & "inm.mdb"
    cnnInmueble.Execute "UPDATE Propietarios SET " & strCamp & "=True WHERE " & strCond, nReg
    cnnInmueble.Close
    Set cnnInmueble = Nothing
    Call rtnBitacora("Printer  " & nReg & " " & Titulo & " Inm:" & DtcAviso(0))
    
    If nReg > 0 Then    'se actualizó por lo menos un registro
        Call RtnPrintReport(Report, Seleccion, Titulo)
        booAfected = True
    Else
        MsgBox "No hay registros que imprimir para " & Titulo
    End If
    '
    End Sub

    '---------------------------------------------------------------------------------------------
    '   Funcion:    Salida
    '
    '   Devuelve una constante de Destino
    '---------------------------------------------------------------------------------------------
    Private Function Salida() As crSalida
    '
    Select Case True
        Case optImp(0): Salida = crPantalla
        Case optImp(1): Salida = crImpresora
    End Select
    '
    End Function

    Private Sub Enviar_Correo()
    'variables locales
    Dim rstlocal As New ADODB.Recordset
    Dim strSQL As String, archivo As String, PyC As Integer
    Dim n(1) As Long
    MousePointer = vbHourglass
    If ChkAviso(0).Value = vbChecked Then
        strSQL = "SELECT * FROM PROPIETARIOs WHERE Recibos=3 AND Convenio=False and email<>''"
        archivo = gcPath & "/aviso1.txt"

    ElseIf ChkAviso(1).Value = vbChecked Then
        strSQL = "SELECT * FROM Propietarios WHERE " & ftnFS(0, "", "") & " AND email<>''"
        archivo = gcPath & "/aviso2.txt"
    Else
        MsgBox "Correos no enviados", vbCritical, App.ProductName
        MousePointer = vbDefault
        Exit Sub
    End If
    With rstlocal
        .Open strSQL, cnnOLEDB & gcPath & StrRutaInmueble & "inm.mdb", _
        adOpenKeyset, adLockOptimistic, adCmdText
        poSendMail.SMTPHostValidation = VALIDATE_HOST_DNS
        poSendMail.EmailAddressValidation = VALIDATE_SYNTAX
        poSendMail.Delimiter = ";"
        
        If Not .EOF And Not .BOF Then
            .MoveFirst
            poSendMail.SMTPHost = "mail.cantv.net"
            poSendMail.FromDisplayName = "Servicio de Administración de Condominio"
            poSendMail.from = "info@administradorasac.com"
            poSendMail.Subject = "Aviso de Cobro"
            
            'mail.HOST = "mail.cantv.net"
            'mail.FromName = "Servicio de Administración de Condominio"
            'mail.From = "facturacion@administradorasac.com.ve"
            'mail.from = "info@administradorasac.com"
            'mail.Subject = "Aviso de Cobro"
            'On Error Resume Next
            Do
                If InStr(!email, ";") = 0 Then
                    'mail.AddAddress !Email, !Nombre
                    poSendMail.Recipient = !email
                    poSendMail.RecipientDisplayName = !Nombre
                Else
                    PyC = InStr(!email, ";")
                    'mail.AddAddress Left(!Email, PyC - 1), !Nombre
                    'mail.AddAddress Trim(Mid(!Email, PyC + 1, 200)), Nombre
                    poSendMail.Recipient = Left(!email, PyC - 1) & ";" & Trim(Mid(!email, PyC + 1, 200))
                    poSendMail.RecipientDisplayName = !Nombre
                    'poSendMail.CcRecipient = Trim(Mid(!Email, PyC + 1, 200))
                    'poSendMail.CcDisplayName = Nombre
                End If
'                mail.AddAddress "ynfantes@cantv.net", !Nombre
                'mail.Body = Mensaje(Archivo, IIf(IsNull(!Nombre), "", !Nombre), !Recibos, Format(!Deuda, "#,##0.00"))
                'mail.Send
                poSendMail.Message = Mensaje(archivo, IIf(IsNull(!Nombre), "", !Nombre), !Recibos, Format(!Deuda, "#,##0.00"))
                poSendMail.Send
                If Err.Number = 0 Then 'si no ocurre ningun error registra el envio

                    cnnConexion.Execute "INSERT INTO Notificaciones_Email (CodInm,Propietario,Nombre,Aviso,Recibos," _
                    & "Deuda,Fecha,Hora,Usuario,PC,email,Ok) VALUES ('" & DtcAviso(0) & "','" & !Codigo & "','" & !Nombre & "'," & _
                    ChkAviso(0) & "," & !Recibos & ",'" & !Deuda & "',Date(),Time(),'" & gcUsuario & "','" & _
                    gcMAC & "','" & !email & "',-1)"
                    Call rtnBitacora("Aviso de Cobro vía email " & DtcAviso(0) & "/" & !Codigo & "....Ok.")
                    n(0) = n(0) + 1
                Else
                    cnnConexion.Execute "INSERT INTO Notificaciones_Email (CodInm,Propietario,Nombre,Aviso,Recibos," _
                    & "Deuda,Fecha,Hora,Usuario,PC,email,Ok) VALUES ('" & DtcAviso(0) & "','" & !Codigo & "','" & !Nombre & "'," & _
                    ChkAviso(0) & "," & !Recibos & ",'" & !Deuda & "',Date(),Time(),'" & gcUsuario & "','" & _
                    gcMAC & "','" & !email & "',0)"
                    Call rtnBitacora("Aviso de Cobro vía email " & DtcAviso(0) & "/" & !Codigo & "....Fallido.")
                    n(1) = n(1) + 1
                End If
                
                'mail.Reset
                'poSendMail.Shutdown
                
                .MoveNext
                
                
            Loop Until .EOF
            MsgBox "Proceso finalizado." & vbCrLf & vbCrLf & "mail enviado(s): " & n(0) + n(1) & vbCrLf & "Con éxito: " & n(0) & " - Fallido: " & n(1) & vbCrLf
        End If
    End With
    MousePointer = vbDefault
    End Sub


    Private Function Mensaje(archivo$, ParamArray datos()) As String

    
    'variables locales
    Dim n As Long
    Dim Dato As Byte
    Dim O As Integer
    n = FreeFile
    Open archivo For Binary As #n
        Do While Not EOF(n)
            Get n, , Dato
            If Dato = 63 Then
                Mensaje = Mensaje & Space(1) & datos(I)
                I = I + 1
            Else
                Mensaje = Mensaje & Chr(Dato)
            End If
        Loop
        
    Close #n

End Function
