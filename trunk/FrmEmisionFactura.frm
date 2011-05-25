VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form FrmEmisionFactura 
   AutoRedraw      =   -1  'True
   Caption         =   "Emision de Facturas"
   ClientHeight    =   3135
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   4680
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   3135
   ScaleWidth      =   4680
   Tag             =   "1"
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      Height          =   7770
      Left            =   180
      TabIndex        =   7
      Top             =   30
      Width           =   11415
      Begin TabDlg.SSTab sstFactura 
         Height          =   6015
         Left            =   390
         TabIndex        =   9
         Top             =   1500
         Width           =   10515
         _ExtentX        =   18547
         _ExtentY        =   10610
         _Version        =   393216
         Tabs            =   2
         TabHeight       =   520
         ShowFocusRect   =   0   'False
         TabCaption(0)   =   "Gastos &Comunes"
         TabPicture(0)   =   "FrmEmisionFactura.frx":0000
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "lblFact(0)"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "flexFactura(0)"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).Control(2)=   "txtFact"
         Tab(0).Control(2).Enabled=   0   'False
         Tab(0).ControlCount=   3
         TabCaption(1)   =   "Gastos &No Comunes"
         TabPicture(1)   =   "FrmEmisionFactura.frx":001C
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "lblFact(1)"
         Tab(1).Control(1)=   "lblFact(2)"
         Tab(1).Control(2)=   "lblFact(3)"
         Tab(1).Control(3)=   "lblFact(4)"
         Tab(1).Control(4)=   "lblFact(5)"
         Tab(1).Control(5)=   "lblFact(6)"
         Tab(1).Control(6)=   "lblFact(7)"
         Tab(1).Control(7)=   "lblFact(8)"
         Tab(1).Control(8)=   "lblFact(9)"
         Tab(1).Control(9)=   "lblFact(10)"
         Tab(1).Control(10)=   "flexFactura(1)"
         Tab(1).ControlCount=   11
         Begin VB.TextBox txtFact 
            Alignment       =   1  'Right Justify
            BackColor       =   &H80000002&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000009&
            Height          =   315
            Left            =   8670
            Locked          =   -1  'True
            TabIndex        =   15
            Top             =   5565
            Width           =   1590
         End
         Begin MSFlexGridLib.MSFlexGrid flexFactura 
            Height          =   4695
            Index           =   0
            Left            =   300
            TabIndex        =   10
            Tag             =   "1000|5000|1600|850|850"
            Top             =   675
            Width           =   9915
            _ExtentX        =   17489
            _ExtentY        =   8281
            _Version        =   393216
            Cols            =   5
            FixedCols       =   0
            BackColorFixed  =   -2147483646
            ForeColorFixed  =   -2147483639
            BackColorBkg    =   -2147483636
            FormatString    =   "Gasto|Descripción|Monto|Alícuota|Fondo"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin MSFlexGridLib.MSFlexGrid flexFactura 
            Height          =   4815
            Index           =   1
            Left            =   -74700
            TabIndex        =   11
            Tag             =   "1200|1200|5300|1600"
            Top             =   675
            Width           =   9915
            _ExtentX        =   17489
            _ExtentY        =   8493
            _Version        =   393216
            Cols            =   4
            FixedCols       =   0
            BackColorFixed  =   -2147483646
            ForeColorFixed  =   -2147483639
            BackColorBkg    =   -2147483636
            MergeCells      =   1
            FormatString    =   "Apartamento|Gasto|Concepto|Monto"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin VB.Label lblFact 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0,00 "
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   10
            Left            =   -74160
            TabIndex        =   25
            Tag             =   "900018"
            Top             =   5625
            Width           =   1200
         End
         Begin VB.Label lblFact 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0,00 "
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   9
            Left            =   -71760
            TabIndex        =   24
            Top             =   5625
            Width           =   1200
         End
         Begin VB.Label lblFact 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0,00 "
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   8
            Left            =   -69720
            TabIndex        =   23
            Top             =   5625
            Width           =   1200
         End
         Begin VB.Label lblFact 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0,00 "
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   7
            Left            =   -67680
            TabIndex        =   22
            Top             =   5625
            Width           =   1200
         End
         Begin VB.Label lblFact 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0,00 "
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   6
            Left            =   -65880
            TabIndex        =   21
            Top             =   5625
            Width           =   1200
         End
         Begin VB.Label lblFact 
            BackStyle       =   0  'Transparent
            Caption         =   "Otros:"
            Height          =   255
            Index           =   5
            Left            =   -66360
            TabIndex        =   20
            Top             =   5640
            Width           =   495
         End
         Begin VB.Label lblFact 
            BackStyle       =   0  'Transparent
            Caption         =   "Gestiones:"
            Height          =   255
            Index           =   4
            Left            =   -68520
            TabIndex        =   19
            Top             =   5640
            Width           =   855
         End
         Begin VB.Label lblFact 
            BackStyle       =   0  'Transparent
            Caption         =   "Intereses:"
            Height          =   255
            Index           =   3
            Left            =   -70440
            TabIndex        =   18
            Top             =   5640
            Width           =   855
         End
         Begin VB.Label lblFact 
            BackStyle       =   0  'Transparent
            Caption         =   "Notificaciones:"
            Height          =   255
            Index           =   2
            Left            =   -72840
            TabIndex        =   17
            Top             =   5640
            Width           =   1095
         End
         Begin VB.Label lblFact 
            BackStyle       =   0  'Transparent
            Caption         =   "G.Admin.:"
            Height          =   255
            Index           =   1
            Left            =   -74880
            TabIndex        =   16
            Top             =   5640
            Width           =   735
         End
         Begin VB.Label lblFact 
            Alignment       =   1  'Right Justify
            Caption         =   "Total Gastos Comunes:"
            Height          =   315
            Index           =   0
            Left            =   6390
            TabIndex        =   14
            Top             =   5565
            Width           =   2235
         End
      End
      Begin VB.PictureBox Picture1 
         AutoSize        =   -1  'True
         Height          =   690
         Left            =   6990
         ScaleHeight     =   630
         ScaleWidth      =   3840
         TabIndex        =   8
         Top             =   270
         Width           =   3900
         Begin VB.CommandButton cmdFactura 
            Caption         =   "&Salir"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   630
            Index           =   3
            Left            =   2895
            MouseIcon       =   "FrmEmisionFactura.frx":0038
            MousePointer    =   99  'Custom
            TabIndex        =   6
            Top             =   0
            Width           =   945
         End
         Begin VB.CommandButton cmdFactura 
            Caption         =   "A&visos de Cobro"
            Enabled         =   0   'False
            Height          =   630
            Index           =   2
            Left            =   1935
            MouseIcon       =   "FrmEmisionFactura.frx":018A
            MousePointer    =   99  'Custom
            TabIndex        =   5
            Top             =   0
            Width           =   945
         End
         Begin VB.CommandButton cmdFactura 
            Caption         =   "&Asignar Gastos"
            Height          =   630
            Index           =   0
            Left            =   15
            MouseIcon       =   "FrmEmisionFactura.frx":02DC
            MousePointer    =   99  'Custom
            TabIndex        =   3
            Top             =   0
            Width           =   945
         End
         Begin VB.CommandButton cmdFactura 
            Caption         =   "&Generar Facturas"
            Enabled         =   0   'False
            Height          =   630
            Index           =   1
            Left            =   975
            MouseIcon       =   "FrmEmisionFactura.frx":042E
            MousePointer    =   99  'Custom
            TabIndex        =   4
            Top             =   0
            Width           =   945
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "&Período (mes - año):"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   660
         Left            =   435
         TabIndex        =   0
         Top             =   315
         Width           =   2580
         Begin VB.ComboBox CmbPeriodo 
            Height          =   315
            Index           =   1
            ItemData        =   "FrmEmisionFactura.frx":0580
            Left            =   1470
            List            =   "FrmEmisionFactura.frx":0582
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   2
            Top             =   225
            Width           =   990
         End
         Begin VB.ComboBox CmbPeriodo 
            Height          =   315
            Index           =   0
            ItemData        =   "FrmEmisionFactura.frx":0584
            Left            =   120
            List            =   "FrmEmisionFactura.frx":05AC
            Style           =   2  'Dropdown List
            TabIndex        =   1
            Top             =   225
            Width           =   1245
         End
      End
      Begin VB.Image imgFactura 
         Enabled         =   0   'False
         Height          =   480
         Index           =   0
         Left            =   30
         Top             =   405
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Label lblFactura 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000009&
         Height          =   315
         Index           =   0
         Left            =   435
         TabIndex        =   12
         Top             =   1080
         Visible         =   0   'False
         Width           =   10470
      End
      Begin VB.Label lblFactura 
         BackColor       =   &H80000002&
         Height          =   315
         Index           =   1
         Left            =   420
         TabIndex        =   13
         Top             =   1080
         Visible         =   0   'False
         Width           =   15
      End
   End
   Begin VB.Image imgFactura 
      Enabled         =   0   'False
      Height          =   480
      Index           =   1
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   480
   End
End
Attribute VB_Name = "FrmEmisionFactura"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
     '29/08/2002---SAC Sistema de Administración de condominios------------------------------------
    'Módulo de Facturación. Genera:
    'Intereses de Mora, Gestión de Cobranzas, Avisos de cobro, Causula penal, Gastos administrati-
    'vos, cheques devueltos, gestion de cobro cheq. dev., Gastos Fijos, Gastos Comunes, Muestra el
    'Resultado Final del Proceso de Facturación según los parametros establecidos por el usuario--
    Dim cnn As ADODB.Connection             'Conexión Local al inmueble seleccionado
    Dim rstPropietario As New ADODB.Recordset
    Dim rstInmueble As ADODB.Recordset
    Dim intMora As Currency
    Dim intAptos As Integer
    Dim intGestion As Currency
    Dim intCheqDev As Currency
    Dim intMesMora As Integer
    Dim IntGA As Currency
    Dim strHora As String                       'hora del comienzo del proceso de fact.
    Dim strGastoMora(0 To 1) As String 'código y titulo de 'Intere7ses de Mora'
    Dim strGastoPenal(0 To 1) As String 'código y titulo de 'Clausula Penal'
    Dim strGestion(0 To 1) As String    'código y titulo de 'Gestion de Cobro'
    Dim strChdev(0 To 1) As String      'código y titulo de 'Cheques Devueltos'
    Dim strGesChDev(0 To 1) As String           'código y titulo de 'Gestión Cheq. Dev.'
    Dim strGastoAdmin(0 To 1) As String 'código y título de 'Gastos Administrativos'
    Dim strCarta(0 To 1) As String      'código y título de 'Gastos de Cartas de Cobro'
    Dim strTelegrama(0 To 1) As String          'Código y título de 'gastos de Telegramas'
    Dim strIVA(0 To 1) As String                'Código y título 'Impuesto Consumo a las Ventas'
    Dim curCarta As Currency                    'Valor de las cartas de deuda
    Dim curTelegrama As Currency                'valor de los telegramas de deuda
    Dim datPeriodo As Date              'contiene el período que se está facturando
    Dim vecFondo()
    '---------------------------------------------------------------------------------------------

    '---------------------------------------------------------------------------------------------
    Private Sub CmbPeriodo_KeyPress(index As Integer, KeyAscii As Integer)  '
    '---------------------------------------------------------------------------------------------
    '
    If KeyAscii = 13 Then
        Select Case index
            'Meses del año
            Case 0: CmbPeriodo(1).SetFocus
            '------------------------
            'Años
            Case 1: cmdFactura(0).SetFocus
            '------------------------
        End Select
    End If
    '
    End Sub

    '29/08/2002--Rutina que controla los sucecos que ocurren al presionar un elemento de la matriz
    Private Sub cmdFactura_Click(index As Integer)  'de botones de comando------------------------
    '---------------------------------------------------------------------------------------------
    '
    Select Case index
        Case 0  'Botón asignar gastos
    '   -------------------------------
            Call rtnAsignacion
            
        Case 1  'Botón generar facturas
    '   -------------------------------
            Call rtnFacturar
            Call rtnBitacora("Facturado Inm.:" & gcCodInm)
            
        Case 2  'Botón Avisos de cobro
    '   -------------------------------
            cmdFactura(2).Enabled = False
            cmdFactura(3).Enabled = False
            Call rtnAC(Left(CmbPeriodo(0), 3) & " - " & CmbPeriodo(1), _
            datPeriodo, crImpresora)
            cmdFactura(3).Enabled = True
            
            '
        Case 3  'Salir del formulario
    '   -------------------------------
            Unload Me
            Set FrmEmisionFactura = Nothing
            
    End Select
    '
    End Sub

    '29/08/2002-----------------------------------------------------------------------------------
    Private Sub Form_Load() 'Carga el Formulario y
    '---------------------------------------------------------------------------------------------
    'variables locales
    Dim Mensaje As String
    Dim I As Integer
    Dim strSQL As String
    '
    txtFact = "0,00"
    'Carga en la lista el año en curso y el siguiente
    For I = 0 To 1: CmbPeriodo(1).AddItem (Year(Date) + I)
    Next
    If Month(Date) = 1 Then CmbPeriodo(1).AddItem (Year(Date) - 1)
    imgFactura(0).Picture = LoadResPicture("CHECKED", vbResBitmap)
    imgFactura(1).Picture = LoadResPicture("UNCHECKED", vbResBitmap)
    
    'Presenta el periodo al mes actual
    CmbPeriodo(0).Text = CmbPeriodo(0).List(Month(Date) - 1)
    CmbPeriodo(1).Text = CmbPeriodo(1).List(IIf(Month(Date) = 1, 1, 0))
    '
    Set rstInmueble = New ADODB.Recordset
    Set cnn = New ADODB.Connection
    '
    'Conexión al inmueble seleccionado
    If (gcCodInm = "") Then
        MsgBox "Debe seleccionar un inmueble para poder llevar a cabo esta operación", vbCritical, App.ProductName
        Exit Sub
    End If
    If Dir(mcDatos, vbArchive) = "" Then
        MsgBox "No consigo la información del inmueble, " & _
        "si el problema persiste contacte al administrador del sistema.", vbCritical, App.ProductName
        Exit Sub
    End If
    cnn.CursorLocation = adUseClient
    cnn.Open cnnOLEDB & mcDatos
    'Carga el vector de fondos-------------------
    With rstInmueble
        .Open "SELECT * FROM Tgastos WHERE Fondo=true;", cnn, adOpenStatic, adLockReadOnly, _
        adCmdText
        ReDim vecFondo(.RecordCount)    'redimensiona la matriz
        If Not .EOF Or Not .BOF Then
            .MoveFirst
            Do
                vecFondo(.AbsolutePosition) = !codGasto
                .MoveNext
            Loop Until .EOF
        End If
        .Close
    End With
    '--------------------------------------------
    rstInmueble.Open "SELECT * FROM inmueble WHERE codinm = '" & gcCodInm & "'", _
        cnnConexion, adOpenKeyset, adLockOptimistic, adCmdText
    'Sale de la rutina si no encuentra información s/el inmueble
    If rstInmueble.EOF Then
        MsgBox ("No encontré Código del Inmueble...."), vbCritical
        Exit Sub
     End If
    '---------------------------
    'chequea campos para calculos de intereses de mora y gastos administrativos
    If IsNull(rstInmueble("codIntMora")) Then
        Mensaje = "Código de Cuenta INTERESES MORATORIOS"
    ElseIf IsNull(rstInmueble("codgastoadmin")) Then
        Mensaje = "Código de Cuenta GASTOS DE ADMINISTRACION"
    ElseIf IsNull(rstInmueble("interes")) Then
        Mensaje = "% de Intereses Moratorios"
    ElseIf IsNull(rstInmueble("CodGestion")) Then
        Mensaje = "Código de Cuenta GESTION DE COBRANZAS"
    ElseIf IsNull(rstInmueble("gestion")) Then
        Mensaje = "% de Gestion de Cobranzas"
    ElseIf IsNull(rstInmueble("Unidad")) Then
        Mensaje = "Nº de Apartamentos en el edificio"
    ElseIf IsNull(rstInmueble("Honorarios")) Then
        Mensaje = "% de HONORARIOS PROFESIONALES"
    ElseIf IsNull(rstInmueble("CodIva")) Then
        Mensaje = "Código del IVA"
    End If
    If Not Mensaje = "" Then
        MsgBox "Falta definir en la ficha del inmueble:" & vbCrLf & Mensaje, vbInformation, _
        App.ProductName
        Exit Sub
    End If
    
    With rstInmueble
    
        If gnMesesMora = 0 And gnPorIntMora > 0 Then
            MsgBox "El porcentaje de Intereses de Mora es: " & gnPorIntMora & vbCrLf & _
            "Los meses para efectuar el cálculo esta establecido en cero" & vbCrLf & _
            "Esta configuración puede ocasionar errores en el cálculo de los avisos de cobro." & vbCrLf & _
            "Haga los cambios pertinentes en los parámetos de facturación o " & vbCrLf & _
            "póngase en contacto con el administrador del sistema", vbCritical, App.ProductName
            Call rtnBitacora("Aviso: Meses Mora igual a cero")
        End If
        strGastoMora(0) = !CodIntMora       'Código Gastos de Mora
        strGastoPenal(0) = !CodGastoPenal   'Código Gastos penales
        strGastoAdmin(0) = !CodGastoAdmin   'Cod. Gastos Administrativos
        strGestion(0) = !CodGestion         'Cod. Gestión de Cobranzas
        strChdev(0) = !CodChdev             'Cod. Cheques Devueltos
        strGesChDev(0) = !CodGesChdev       'Cod.Gestión cheques Devueltos
        strCarta(0) = !CodCarta             'código de Cartas de cobros
        strTelegrama(0) = !CodTelegrama     'Código cobro de Telegramas
        strIVA(0) = !CodIVA     'Código de Impuesto a las Ventas
        intAptos = !Unidad             'Cantidad de Apartamentos del condominio
        IntGA = !Honorarios             'Monto por Gastos Administrativos
        intMora = !interes               'Porcentaje cobro por mora
        intMesMora = !MesesInt      'Parametro cálculo >
        intGestion = !Gestion           'Porcentaje Gestion de Cobranza
        intCheqDev = !CheqDev           'Porcentaje por gestion cobro cheq. dev.
        curCarta = !CostCarta           'Costo de Los avisos de cobro
        curTelegrama = !CostTelegrama   'Costo de los telegramas de cobro
        
    End With
    'ubica titulos de los Gastos
    
    strGastoMora(1) = ftnTitulo(strGastoMora(0))     'Intereses moratorios
    strGastoPenal(1) = ftnTitulo(strGastoPenal(0))  'Pago a cuenta demandada
    strGestion(1) = ftnTitulo(strGestion(0))        'Gestion de cobranzas
    strCarta(1) = ftnTitulo(strCarta(0))            'Emisión cartas de morosidad
    strTelegrama(1) = ftnTitulo(strTelegrama(0))    'Cancelación telegramas por deuda
    strChdev(1) = ftnTitulo(strChdev(0))            'Cheques devuelto
    strGesChDev(1) = ftnTitulo(strGesChDev(0))   'Gestion cheq.Dev
    strGastoAdmin(1) = ftnTitulo(strGastoAdmin(0))  'Gastos Administrativos
    strIVA(1) = ftnTitulo(strIVA(0))        'Código de Impuesto a las Ventas
    
    '
    'consultas necesarias para el proceso
    'strSql = "SELECT Sum(Monto) AS IntGes, Codigo as codprop FROM DetFact WHERE (CodGasto = '" & _
     strGastoMora(0) & "' or CodGasto = '" & strGestion(0) & "') and Fact in (SELECT FacT FROM Factura WHERE Saldo > 0) GROUP BY Codigo"
    cnn.Execute "DELETE * FROM InteGest"
    strSQL = "INSERT INTO InteGest(IntGes,CodProp) SELECT Sum(DetFact.Monto) AS IntGes, Factura.codprop FROM DetFact I" _
    & "NNER JOIN Factura ON DetFact.Fact = Factura.FACT WHERE (((Factura.Saldo)>0) AND ((DetFac" _
    & "t.CodGasto)='" & strGastoMora(0) & "' Or (DetFact.CodGasto)='" & strGestion(0) & "')) GR" _
    & "OUP BY Factura.codprop;"
    cnn.Execute strSQL
    
    'Call rtnGenerator(mcDatos, strSql, "InteGest")
    'configura los grid en la pantalla
    For I = 0 To 1
        Call centra_titulo(flexFactura(I), True)
        flexFactura(1).ColAlignment(I) = flexAlignCenterCenter
    Next
    flexFactura(0).ColAlignment(0) = flexAlignCenterCenter
    flexFactura(1).ColAlignment(2) = flexAlignLeftCenter
    '
    End Sub

    '---------------------------------------------------------------------------------------------
    Private Sub rtnFacturar()   '
    '---------------------------------------------------------------------------------------------
    'variables locales
    Dim I As Integer, Inicio, Fin
    Dim datNPeriodo As Date
    'Confirma la ejecución del proceso
    If Not Respuesta("Desea iniciar el proceso de facturación?") Then Exit Sub
    Inicio = Timer
    strHora = Format(Time(), "hh:mm:ss AMPM")
    cmdFactura(1).Enabled = False
    cmdFactura(0).Enabled = False
    'Call rtnver
    For I = 0 To 1: lblFactura(I).Visible = Not lblFactura(I).Visible
    Next
    lblFactura(0) = "Eliminando Información anterior..."
    lblFactura(1).Width = 1500
    For I = 0 To 1: lblFactura(I).Refresh
    Next
    Screen.MousePointer = vbHourglass
    If rstInmueble("factCH") Then
        cnn.Execute "UPDATE ChequeDevuelto SET Recuperado=True,Freg=Date(),Usuario='" & _
        gcUsuario & "' WHERE recuperado=False;"
        cnn.Execute "DELETE * FROM Factura WHERE Fact LIKE 'CHD%';"
    Else
        'aqui reasgina los cheques devueltos no cobrados a un mes
        'posterior
        datNPeriodo = Format(datPeriodo, "mm/dd/yy")
        datNPeriodo = DateAdd("M", 1, datNPeriodo)
        cnn.Execute "UPDATE Factura SET Periodo ='" & datNPeriodo & "' WHERE Fact Like 'CHD%'"
    End If
        Call rtnBitacora("Facturando Periodo " & Format(datPeriodo, "mm/dd/yy") & " Inmueble " & _
    gcCodInm)
    cnn.Execute "DELETE * FROM DetFact WHERE Periodo=#" & datPeriodo & "# OR Monto Is Null or Monto=0;"
    '---------------------------------------------------------------------------------------------
    'Declara variables temporales en esta instancia del módulo
    Dim strNF As String         ' Numero de Factura
    Dim strSQL As String
    Dim strD$, strDet$, datPer1$, codProv$, nomProv$
    Dim curGestCob@, curGesCheDev@, curCar@, curTel@, curTF@, curGA@, Saldo@, curIVA@
    Dim rstHonoGes As New ADODB.Recordset
    '---------------------------------------------------------------------------------------------
    'Agrega gastos No Comunes Detalle TDF detalle
    lblFactura(0) = "Generando Facturas..."
    lblFactura(1).Width = 2000
    For I = 0 To 1: lblFactura(I).Refresh
    Next
    strNF = Format(datPeriodo, "ddyy") & Right(gcCodInm, 3)
    cnn.Execute "INSERT INTO DetFact (FACT, Detalle,Codigo,CodGasto,Periodo,Monto,Fecha,Hora," _
    & "Usuario) SELECT '" & strNF & "' & FORMAT(P.ID,'000') as F,G.Concepto,G.Codapto,G.Cod" _
    & "Gasto,G.Periodo,G.Monto,Date(),'" & strHora & "',G.Usuario FROM GastoNoComun as G IN" _
    & "NER JOIN Propietarios as P ON G.CodApto=P.Codigo WHERE Periodo=#" & datPeriodo & "#"
    '
    lblFactura(0) = "Insertando Gastos Comunes..."
    lblFactura(1).Width = 3500
    For I = 0 To 1: lblFactura(I).Refresh
    Next
    'Agrega ahora los gastos comunes  Tdf Detalle
    Call asig_GC(strNF, strHora)
    'incluye encabezado de la factura
    lblFactura(0) = "Encabezado de la Factura..."
    lblFactura(1).Width = 4000
    For I = 0 To 1: lblFactura(I).Refresh
    Next
    cnn.Execute "INSERT INTO Factura (Fact,Periodo,codprop, Facturado, Pagado, Saldo, freg, usu" _
    & "ario, Fecha,FechaFactura ) SELECT '" & strNF & "' & Format(ID,'000'),Periodo,Codigo,Fact" _
    & "urado, 0, Facturado, DATE(),'" & gcUsuario & "','" & Left(strHora, 8) & "',Date() FROM Fact6;"
    '
    'Actualiza Movimientos de los Fondos del Condominio
    lblFactura(0) = "Actualización Cuentas de Fondos...."
    lblFactura(1).Width = 6000
    For I = 0 To 1: lblFactura(I).Refresh
    Next
    'Call rtnEtiqueta("Actualizando Fondos Especiales Inmueble '" & gcCodInm & "'", 6000)
    Dim rstFR As ADODB.Recordset
    Set rstFR = New ADODB.Recordset
    '
    rstFR.Open "SELECT CodGasto,cargado,Descripcion,Monto FROM AsignaGasto WHERE CodGasto IN (S" _
    & "ELECT CodGasto FROM Tgastos WHERE Fondo=True;) AND Cargado=#" & datPeriodo & "# UNION SE" _
    & "LECT CodGasto,Periodo,concepto, Sum(Monto)  FROM GastoNoComun GROUP BY CodGasto, Concept" _
    & "o, Periodo HAVING CodGasto In (Select CodGasto FROM Tgastos WHERE Fondo=True) AND Period" _
    & "o=#" & datPeriodo & "#;", cnn, adOpenKeyset, adLockOptimistic
    
    If Not rstFR.EOF Or Not rstFR.BOF Then
        '
        With rstFR
            .MoveFirst
            Do Until .EOF
                'Inserta el movimiento en TDF MovFondo
                cnn.Execute "INSERT INTO movFondo (CodGasto,Fecha,Tipo,Periodo,Concepto,Debe,Ha" _
                & "ber) VALUES ('" & !codGasto & "',Date(),'FA',#" & datPeriodo & "#,'" & _
                !Descripcion & "',0,'" & !Monto & "');"
                '
'                'Actualiza el saldo actual, Saldo del mes Facturado
'                Saldo = Saldo_Actual(!codGasto)
'                cnn.Execute "UPDATE Tgastos SET SaldoActual ='" & Saldo & "',Saldo" & _
'                Format(datPeriodo, "d") & "='" & Saldo & "' WHERE Codgasto='" & !codGasto & "';"
                .MoveNext
            Loop
            rstFR.Close
            'Actualiza la tabla Pago_Inf y el fondo de reserva del edificio
            Call Actuliza_Pago_Inf
        '
        End With
    '
    End If
    Set rstFR = Nothing
    'ACtualiza deuda de propietarios y del inmueble
    lblFactura(0) = "Actualización Deuda del inmueble..."
    lblFactura(1).Width = 6500
    For I = 0 To 1: lblFactura(I).Refresh
    Next
    'Call rtnEtiqueta("Actualizado Deuda Inmueble '" & gcCodInm & "'", 6500)
    Call Actulizar_DeudaInm
    lblFactura(0) = "Actualización Deuda Propietarios..."
    lblFactura(1).Width = 8000
    For I = 0 To 1: lblFactura(I).Refresh
    Next
    'Call rtnEtiqueta("Actualizando deuda propietarios...", 8000)
    Call Actualiza_Deuda_Propietario
    '
    'Genera la cuenta por pagar del inmueble Honorarios, Gestiones, cartas y telegramas
    ' impuesto a las ventas
    '------------------------
    curGA = CLng(CCur(Total_gasto(strGastoAdmin(0)) / (1 + (gnIva / 100))) * 100) / 100
    curGestCob = Total_gasto(strGestion(0))
    curGesCheDev = Total_gasto(strGesChDev(0))
    curCar = Total_gasto(strCarta(0))
    curTel = Total_gasto(strTelegrama(0))
    curIVA = CLng((curGA * gnIva / 100) * 100) / 100 'Total_gasto(strIVA(0))
    '
    'datPer1 = DateAdd("m", 1, Format(datPeriodo, "MM/DD/YY"))
    '
    If gnCta = CUENTA_INMUEBLE Then
    
        codProv = sysCodPro
        nomProv = sysEmpresa
        curTF = curGestCob + curGesCheDev + curCar + curTel + curGA + curIVA
        strDet = "HONORARIOS ADMINISTRATIVOS MES: " & Format(datPeriodo, "dd/yyyy") & ", GESTIONES" _
        & " PERIODO: " & Format(datPeriodo, "DD/YYYY") & IIf(curCar > 0, ", CARTAS", "") & _
        IIf(curTel > 0, ", TELEGRAMAS", "") & IIf(curGesCheDev > 0, ", GEST. CHEQ. DEV", "") & " - " _
        & gcNomInm
        '
        'Instruccion SQL para cargar los gastos de la factura
        strSQL = "SELECT SUM(Monto),Detalle as Concepto,CodGasto FROM DetFact WHERE CodGasto IN ('" _
        & strGestion(0) & "','" & strGesChDev(0) & "','" & strCarta(0) & "','" & strTelegrama(0) _
        & "','" & strGastoAdmin(0) & "') AND Periodo=#" & datPeriodo & "# GROUP BY CodGasto,Detalle UNION SELECT '" & curIVA & "', Tgastos.Titulo,T" _
        & "Gastos.CodGasto From Tgastos WHERE Tgastos.CodGasto='" & strIVA(0) & "';"
        '
    Else
        '
        codProv = "0557"
        nomProv = "JORGE MARVAL"
        curTF = curGA + curGestCob + curIVA
        strDet = "HONORARIOS ADMINISTRATIVOS MES: " & Format(datPeriodo, "dd/yyyy") & ", G" _
        & "ESTIONES PERIODO: " & Format(datPeriodo, "DD/YYYY") & " - " & gcNomInm
        '
        'instruccion sql para el cargado de la cxp
        strSQL = "SELECT SUM(Monto),Detalle as Concepto,CodGasto FROM DetFact WHERE CodGasto IN ('" & _
        strGestion(0) & "','" & strGastoAdmin(0) & "') AND Periodo=#" & datPeriodo & "# GROUP BY CodGasto,Detalle UNION SELECT '" & curIVA & "', Tgastos.Titulo,T" _
        & "Gastos.CodGasto From Tgastos WHERE Tgastos.CodGasto='" & strIVA(0) & "';"
        '
    End If
    '
    'Ingresar factura a Cpp
    strD = FrmFactura.FntStrDoc 'obtiene el correlativo del documento
    '
    cnnConexion.Execute "INSERT INTO Cpp(Tipo,Ndoc,Fact,CodProv,Benef,Detalle,Monto,Ivm,Tot" _
    & "al,FRecep,Fecr,Fven,CodInm,Moneda,Estatus,Usuario,Freg) VALUES('FC','" & strD & "','" _
    & "F" & gcCodInm & "','" & codProv & "','" & nomProv & "','" & strDet & "','" & curTF & _
    "',0,'" & curTF & "',Date(),Date(),'" & DateAdd("D", 30, Date) & "','" & _
    gcCodInm & "','BS'" & ",'ASIGNADO','" & gcUsuario & "',DATE())"
    '
    'Agrega la informacion a la tabla cargado
    rstHonoGes.Open strSQL, cnn, adOpenStatic, adLockReadOnly, adCmdText
    '
    With rstHonoGes
    '
        If Not .EOF Or Not .BOF Then
            '
            .MoveFirst
            '
            Do
                
                cnn.Execute "INSERT INTO Cargado(Ndoc,CodGasto,Detalle,Periodo,Monto,Fecha," _
                & "Hora,Usuario) VALUES ('" & strD & "','" & !codGasto & "','" & !Concepto _
                & "',#" & datPeriodo & "#,'" & IIf(!codGasto = strGastoAdmin(0), _
                CLng((.Fields(0) / (1 + (gnIva / 100))) * 100) / 100, CLng(.Fields(0) * 100) / 100) _
                & "',Date(),Time(),'" & gcUsuario & "')"
                
                .MoveNext
                
            Loop Until .EOF
            '
        End If
    '
    End With
    '
    rstHonoGes.Close
    '
'    'agrega al cargado el registro correspondiente a los gastos de administración
'    cnn.Execute "INSERT INTO Cargado(Ndoc,CodGasto,Detalle,Periodo,Monto,Fecha,Hora,Usuario) VA" _
'    & "LUES ('" & strD & "','" & strGastoAdmin(0) & "','" & strGastoAdmin(1) & "','" & datPer1 _
'    & "','" & curGA & "',Date(),Time(),'" & gcUsuario & "')"
'    '
'    'veirificamos la correspondencia de los gastos administrativos pagados por anticipado
'    strSql = "SELECT * FROM Cargado WHERE CodGasto ='" & strGastoAdmin(0) & "' AND Periodo =#" _
'    & datPeriodo & "#"
'
'    rstHonoGes.Open strSql, cnn, adOpenKeyset, adLockOptimistic, adCmdText
'
'    If Not IsNull(rstHonoGes("Monto")) And (Not rstHonoGes.EOF And Not rstHonoGes.BOF) Then
'
'        If rstHonoGes("Monto") < curGA Then 'el monto cargado es mayor al monto pagado
'                                                                    'se genera la factura de ajuste
'            curTF = curGA - rstHonoGes("Monto")
'            strD = FrmFactura.FntStrDoc 'obtiene el correlativo del documento
'            strDet = "AJUSTE HONORARIOS ADMINISTRATIVOS MES: " & Format(datPeriodo, "DD/YYYY")
'            '
'            'ingresa la cuenta por pagar
'            cnnConexion.Execute "INSERT INTO Cpp(Tipo,Ndoc,Fact,CodProv,Benef,Detalle,Monto,Ivm" _
'            & ",Total,FRecep,Fecr,Fven,CodInm,Moneda,Estatus,Usuario,Freg) VALUES('FC','" & strD _
'            & "','" & "F" & gcCodInm & "','" & codProv & "','" & nomProv & "','" & strDet & _
'            "','" & curTF & "',0,'" & curTF & "',Date(),Date(),'" & DateAdd("D", 30, Date) & "','" & _
'            gcCodInm & "','BS'" & ",'ASIGNADO','" & gcUsuario & "',DATE())"
'            '
'            'ingresa el cargado respectivo
'            cnn.Execute "INSERT INTO Cargado(Ndoc,CodGasto,Detalle,Periodo,Monto,Fecha,Hora,Usu" _
'            & "ario) VALUES ('" & strD & "','" & strGastoAdmin(0) & "','" & strGastoAdmin(1) & _
'            "','" & Format(datPeriodo, "mm/dd/yyyy") & "','" & curTF & "',Date(),Time(),'" & _
'            gcUsuario & "')"
'
'        End If
'
'    End If
'    rstHonoGes.Close
'    Set rstHonoGes = Nothing
    
    cnn.Execute "UPDATE Propietarios SET Carta=False, Telegrama=False;"
    '
    'actualiza el catalogo de gastos
    cnn.Execute "UPDATE Tgastos SET Van=Van + 1 WHERE Cuotas > 0;"
    'Imprime el paquete completo
    lblFactura(0) = "Imprimiendo Reportes....."
    lblFactura(1).Width = 9000
    For I = 0 To 1: lblFactura(I).Refresh
    Next
    
    curDeuda = Deuda_Total(datPeriodo)
    Call Printer_PaqueteCompleto(datPeriodo, curDeuda, crImpresora, True)
'    +++++++++++++++++++++++++++++++++++++++++++++++
'    VAMOS ELIMINAR LA GENERACION DE LOS ARCHIVOS HTML EN ESTE
'    MONETO PARA AGILIZAR EL PROCESO DE FACTURACION
'    +++++++++++++++++++++++++++++++++++++++++++++++
'    LblFactura(0) = "Generando archivos *.html...."
'    LblFactura(0) = "Finalizando el proceso....."
'    LblFactura(1).Width = 10470
'    For I = 0 To 1: LblFactura(I).Refresh
'    Next
'    Call rtnEtiqueta("Generando archivos .html......", 10470)
'    Call Enviar_ACemail(CDate(datPeriodo), , True)
'    +++++++++++++++++++++++++++++++++++++++++++++++
    For I = 0 To 1: lblFactura(I).Visible = Not lblFactura(I).Visible
    Next
    Screen.MousePointer = vbNormal
    Fin = Timer
    cmdFactura(2).Enabled = True
    '
    End Sub

    '---------------------------------------------------------------------------------------------
    Private Sub rtnAsignacion() '
    '---------------------------------------------------------------------------------------------
    'variables locales
    Dim AdoFactura As ADODB.Recordset
    Dim rstGAF As ADODB.Recordset
    Dim rstGC As ADODB.Recordset
    Dim rstGNC As ADODB.Recordset
    Dim strGC As String
    Dim strGNC As String
    Dim strCriterio As String
    Dim strFact As String
    Dim strSQL As String
    Dim vecGasto(2 To 5) As String
    Dim datHora As Date
    Dim curMontoTotal As Currency
    Dim I As Integer, K As Integer
    
    '
    Set AdoFactura = New ADODB.Recordset
    '
    Call rtnBitacora("Asignando Gastos Periodo " _
    & "01/" & CmbPeriodo(0) & "/" & CmbPeriodo(1) & " Inmueble: " & gcCodInm)
    datPeriodo = Format(CDate("01/" & CmbPeriodo(0) & "/" & CmbPeriodo(1)), "mm/dd/YY")
    'Confirma que el periodo seleccionado no este ya procesado
    AdoFactura.Open "SELECT * FROM factura WHERE periodo = #" & datPeriodo _
        & "# AND Fact not LIKE 'CHD%'", cnn, adOpenKeyset, adLockOptimistic
    If Not AdoFactura.EOF And Not AdoFactura.BOF Then
        MsgBox "Sr. Usuario : " & vbCrLf & _
        " Este periodo ya fue facturado, por favor rectifique...", _
        vbInformation, "Facturacion"
        AdoFactura.Close
        Set AdoFactura = Nothing
        Call rtnBitacora("Período ya facturado...")
        Exit Sub    'Sale del proceso.......
    End If
    AdoFactura.Close
    '------
    strSQL = "SELECT * FROM AsignaGasto WHERE CodGasto='" & gcCodFondo & "' AND Cargado=" _
    & "#" & datPeriodo & "#;"
    AdoFactura.Open strSQL, cnn, adOpenKeyset, adLockOptimistic, adCmdText
    
    If AdoFactura.EOF And AdoFactura.BOF Then
    
        MsgBox "Sr. Usuario debe Pre-Facturar antes de Facturar este Período..", vbInformation, _
        "Facturación"
        AdoFactura.Close
        Set AdoFactura = Nothing
        Call rtnBitacora("Período sin Pre-Facturar...")
        Exit Sub
        
    End If
    MousePointer = vbHourglass
    'Call rtnver
    For I = 0 To 1: lblFactura(I).Visible = Not lblFactura(I).Visible
    Next
    lblFactura(0) = "Generando Espacio de Trabajo..."
    lblFactura(1).Width = 872
    For I = 0 To 1: lblFactura(I).Refresh
    Next
    'Call rtnEtiqueta("Generando Espacio de Trabajo.....", 872)
    strFact = "'" & Right(gcCodInm, 2) & Format(datPeriodo, "ddyy") & "'"
    '
    'Genera las consultas GastoAdmin,GastoMora y gestion
    vecGasto(2) = "='" & strGastoAdmin(0) & "'"
    vecGasto(3) = "='" & strGastoMora(0) & "'"
    vecGasto(4) = "='" & strGestion(0) & "'"
    vecGasto(5) = "NOT IN ('" & strGastoAdmin(0) & "','" & strGastoMora(0) & "','" _
    & strGestion(0) & "')"
    
    
    Call rtnMakeQuery(datPeriodo, vecGasto(2), vecGasto(3), vecGasto(4), vecGasto(5))
    '
    '---------------------------------------------------------------------------------------------
    datHora = Format(Time, "hh:mm ampm")
    cnn.BeginTrans
    'Elimina los intereses de mora, Gestion de cobranzas, Cartas y Telegramas de Cobranza
    'de haberlos creado anteriormente
    lblFactura(0) = "Eliminando Información Anterior..."
    lblFactura(1).Width = 1744
    For I = 0 To 1: lblFactura(I).Refresh
    Next
    'Call rtnEtiqueta("Eliminando Información anterior.....", 1744)
    cnn.Execute "DELETE * FROM GastoNoComun WHERE codgasto IN ('" & strGastoMora(0) & "','" _
    & strGastoPenal(0) & "','" & strGestion(0) & "','" & strCarta(0) & "','" & strTelegrama(0) _
    & "','" & strChdev(0) & "','" & strGesChDev(0) & "','" & strGastoAdmin(0) & "','" & strIVA(0) _
    & "') AND periodo = #" & datPeriodo & "# AND PF=TRUE"
    '
    'Genera Honorarios por Administración del Inmueble
    '
    If rstInmueble("FactGA") Then   'verifica si se cobran los gastos de administración
        lblFactura(0) = "Calculando Honorarios por Administración..."
        lblFactura(1).Width = 2616
        For I = 0 To 1: lblFactura(I).Refresh
        Next
        'Call rtnEtiqueta("Calculando Honorarios por Administración...", 2616)
        '--------------
        Set rstGAF = New ADODB.Recordset
        strSQL = "SELECT * FROM Tgastos WHERE CodGasto='" & strGastoAdmin(0) & "'"
        rstGAF.Open strSQL, cnn, adOpenStatic, adLockOptimistic, adCmdText
        '
        If rstGAF!Alicuota Then
            strCriterio = "* Alicuota/100"
        Else
            strCriterio = "/" & intAptos
        End If
        '
        If Not rstGAF!Comun Then 'si el gasto no es común
            '
            cnn.Execute "INSERT INTO GastoNoComun (PF,CodApto, CodGasto, Concepto,Monto,periodo,Fec" _
            & "ha,Hora,Usuario) SELECT True as p, Codigo,'" & strGastoAdmin(0) & "','" _
            & strGastoAdmin(1) & "',Clng(CCur('" & IntGA & "'" & strCriterio & ")*100)/100 as M,#" & datPeriodo & "#" _
            & ",DATE() as F,'" & datHora & "','" & gcUsuario & "' FROM Propietarios WHERE Codig" _
            & "o<>'U" & gcCodInm & "'"
            'Calcula el IVA si esta áctivo
            If gnIva > 0 Then
                '
                'cnn.Execute "INSERT INTO GastoNoComun (PF,CodApto, CodGasto, Concepto,Monto,periodo" _
                & ",Fecha,Hora,Usuario) SELECT True as p, Codigo,'" & strIVA(0) & "','" & strIVA(1) _
                & "', Clng((" & IntGA & strCriterio & ") *" & gnIva & "/100) as M,#" & datPeriodo & "#" _
                & ",DATE() as F,'" & datHora & "','" & gcUsuario & "' FROM Propietarios WHERE Codig" _
                & "o<>'U" & gcCodInm & "'"
                cnn.Execute "UPDATE GastoNoComun SET Monto = Clng(Monto * (1 + '" & _
                gnIva / 100 & "')*100)/100 WHERE CodGasto='" & strGastoAdmin(0) & _
                "' AND Periodo=#" & datPeriodo & "#;"
            '
            End If
            '
        End If
        rstGAF.Close
        Set rstGAF = Nothing
    End If
    '
        'Genera intereses de mora a/c propietario factIM {Si/No}---------------------------
    If rstInmueble("factIM") Then
        lblFactura(0) = "Calculando Intereses de Mora..."
        lblFactura(1).Width = 3488
        For I = 0 To 1: lblFactura(I).Refresh
        Next
        'Call rtnEtiqueta("Calculando Intereses de Mora...", 3488)
        '
        cnn.Execute "INSERT INTO GastoNoComun(PF,  CodApto, CodGasto, Concepto, Monto,Periodo, " _
            & "Fecha, Hora, Usuario) SELECT True AS P, Factura.codprop, '" & strGastoMora(0) & _
            "', '" & strGastoMora(1) & "', CCur((Sum(Factura.Saldo) -iif(isnull(InteGest.IntGes" _
            & "),0,InteGest.IntGes))* '" & intMora & "' /100), #" & datPeriodo & "#, '" & Date & _
            "', '" & datHora & "', '" & gcUsuario & "' FROM Factura LEFT JOIN InteGest ON Factu" _
            & "ra.codprop = InteGest.codprop WHERE (((Factura.Saldo)<>0) AND ((Factura.Periodo)" _
            & "<#" & datPeriodo & "#) AND ((Factura.codprop) In (SELECT Codigo FROM Propietario" _
            & "s WHERE Recibos >= " & intMesMora & "))) GROUP BY Factura.codprop, InteGest.IntG" _
            & "es, InteGest.IntGes;"

'            cnn.Execute "INSERT INTO GastoNoComun(PF,  CodApto, CodGasto, Concepto, Monto,Periodo, " _
            & "Fecha, Hora, Usuario) SELECT True as P, CodProp,'" & strGastoMora(0) & "','" _
            & strGastoMora(1) & "', CLng(Sum(clng(Saldo))*" & intMora & "/100) as Hono,#" & datPeriodo _
            & "#,'" & Date & "' as Fec,'" & datHora & "' as Hor,'" & gcUsuario & "' FROM Factur" _
            & "a WHERE Factura.Saldo<>0 AND Factura.Periodo<#" & datPeriodo & "# AND CodProp IN" _
            & " (SELECT Codigo FROM Propietarios WHERE Recibos >= " & intMesMora & ") GROUP BY F" _
            & "actura.codprop;"
        
    '
    End If
    '
    'Se Genera la Gestión de Cobranzas factGC {Si/No}----------------------------------
    If rstInmueble("factGC") Then
        lblFactura(0) = "Calculando Gestión de Cobranzas..."
        lblFactura(1).Width = 4360
        For I = 0 To 1: lblFactura(I).Refresh
        Next
'        If Not rstInmueble("factiM") Then
'            strsql = "SELECT Sum(Monto) AS IntGes, codprop FROM DetFact WHERE (CodGasto = '" & _
'            strGastoMora(0) & "' or CodGasto ='" & strGestion(0) & "') and Fact" _
'            & " in (SELECT FacT FROM Factura WHERE Saldo >0)"
'
'            Call rtnGenerator(mcDatos, strsql, "InteGest")
'        End If
        'Call rtnEtiqueta("Calculando Gestión de Cobranzas...", 4360)
        cnn.Execute "INSERT INTO GastoNoComun(PF,  CodApto, CodGasto, Concepto, Monto,Periodo, " _
            & "Fecha, Hora, Usuario) SELECT True AS P, Factura.codprop, '" & strGestion(0) & _
            "', '" & strGestion(1) & "', CCur((Sum(Factura.Saldo) -iif(isnull(InteGest.IntGes" _
            & "),0,InteGest.IntGes))* '" & intGestion & "' /100), #" & datPeriodo & "#, '" & Date & _
            "', '" & datHora & "', '" & gcUsuario & "' FROM Factura LEFT JOIN InteGest ON Factu" _
            & "ra.codprop = InteGest.codprop WHERE (((Factura.Saldo)<>0) AND ((Factura.Periodo)" _
            & "<#" & datPeriodo & "#) AND ((Factura.codprop) In (SELECT Codigo FROM Propietario" _
            & "s WHERE Recibos >= " & intMesMora & "))) GROUP BY Factura.codprop, InteGest.IntG" _
            & "es, InteGest.IntGes;"

    '
    End If
    '
    'Genera el cobro de Cartas y Telegramas--------------------------------------------
     If rstInmueble("EmiteC1") Then
        lblFactura(0) = "Cobranza Cartas y Telegramas...."
        lblFactura(1).Width = 5232
        For I = 0 To 1: lblFactura(I).Refresh
        Next
        'Call rtnEtiqueta("Cobrando Cartas y Telegramas...", 5232)
        cnn.Execute "INSERT INTO GastoNoComun (PF, CodApto, CodGasto, Concepto, Monto, Periodo," _
        & "Fecha,Hora,Usuario) SELECT True, Codigo,'" & strCarta(0) & "','" & strCarta(1) & _
        "','" & curCarta & "',#" & datPeriodo & "#,Date(),'" & datHora & "','" & gcUsuario & _
        "' FROM Propietarios WHERE Carta=True"
    End If
    If rstInmueble("EmiteT1") Then
        cnn.Execute "INSERT INTO GastoNoComun (PF,CodApto,CodGasto,Concepto,Monto, Periodo,Fech" _
        & "a,Hora,Usuario) SELECT True, Codigo,'" & strTelegrama(0) & "','" & strTelegrama(1) & _
        "','" & curTelegrama & "',#" & datPeriodo & "#,Date(),'" & datHora & "','" & gcUsuario _
        & "' FROM Propietarios WHERE Telegrama=True"
    End If
    '
    'Genera cobro de cheques devueltos y gestion cobro cheq.dev
    If rstInmueble("factCH") Then
        lblFactura(0) = "Verificando Cheques Devueltos..."
        lblFactura(1).Width = 6104
        For I = 0 To 1: lblFactura(I).Refresh
        Next
        'Call rtnEtiqueta("Verificando Cheques Devueltos...", 6104)
        cnn.Execute "INSERT INTO GastoNoComun (PF,CodApto,CodGasto,Concepto,Monto,Periodo,Fecha" _
        & ",Hora,Usuario) SELECT True,Codigo,'" & strChdev(0) & "', 'CHEQUE DEVUELTO BCO. '" & _
        " & Banco & '  # ' & NumCheque , Monto, #" & datPeriodo & "#,Date(),'" & datHora & "','" _
        & gcUsuario & "' FROM ChequeDevuelto WHERE Recuperado=False"
    '
        
        'Call rtnEtiqueta("Calculando Gestión Cobro Cheques Devueltos", 6976)
        cnn.Execute "INSERT INTO GastoNoComun (PF,Codapto, codGasto, Concepto, Monto, " _
        & "Periodo, Fecha, Hora, Usuario) SELECT True, Codigo,'" & strGesChDev(0) & "','" & _
        strGesChDev(1) & "', CCur(Sum(Monto) *" & intCheqDev & "/100),#" & datPeriodo & "#,Date" _
        & "(),'" & datHora & "','" & gcUsuario & "' FROM ChequeDevuelto WHERE Recuperado=False " _
        & "AND Comision = True GROUP BY codigo"
        'Elimina la información de cheques devueltos
        cnn.Execute "DELETE * FROM ChequeDevuelto IN '" & gcPath & "\sac.mdb' WHERE CodInm='" & _
        gcCodInm & "' AND Numero IN (SELECT NumCheque FROM ChequeDevuelto WHERE Recuperado=False)"
        '
    End If
    '
    cnn.CommitTrans
    '       Proceso de Asignar Gastos Comunes
    '       ubica los gastos {COMUNES / NO COMUNES} del periodo seleccionado
    lblFactura(0) = "Ubicando Gastos Comunes"
    lblFactura(1).Width = 7000
    For I = 0 To 1: lblFactura(I).Refresh
    Next
    'Call rtnEtiqueta("Ubicando Gastos Comunes / No comunes", 7000)
    'crea una instrancia de estos ADODB.Recordsets
    Set rstGC = New ADODB.Recordset
    Set rstGNC = New ADODB.Recordset
    '
    strGC = "SELECT AG.CodGasto, AG.Descripcion, CCur(Sum(AG.Monto)) AS Total, AG.Alicuota" _
    & ", AG.Comun FROM AsignaGasto AS AG Where (((AG.Comun) = True) And ((AG.Cargado) = #" _
    & datPeriodo & "#)) GROUP BY AG.CodGasto, AG.Descripcion, AG.Alicuota, AG.Comun, AG.Cargado" _
    & " ORDER BY AG.CodGasto;"
    '
    strGNC = "SELECT * FROM GastoNoComun WHERE Periodo=#" & datPeriodo & "# ORDER BY CodApto, CodGasto;"
    'Crea una consulta Gastos No Comunes
    Call rtnGenerator(mcDatos, strGNC, "qdfGNC")
    'abre el par de recodsets
    rstGC.Open strGC, cnn, adOpenKeyset, adLockOptimistic, adCmdText
    rstGNC.Open strGNC, cnn, adOpenKeyset, adLockOptimistic, adCmdText
    'ubica los conceptos de gastos que sea el fondo de reserva
    lblFactura(0) = "Configurando Vista..."
    lblFactura(1).Width = 8000
    For I = 0 To 1: lblFactura(I).Refresh
    Next
    'Call rtnEtiqueta("Configurando Vista...", 8000)
    
    With flexFactura(0)
        .Visible = False
        .Rows = rstGC.RecordCount + 1
        lblFact(0) = "Total Gastos Comunes (" & .Rows - 1& & "):"
        I = 1
        curMontoTotal = 0
        rstGC.MoveFirst
        .ColAlignment(1) = flexAlignLeftCenter
        Do While Not rstGC.EOF
            If rstGC("comun") Then
                .TextMatrix(I, 0) = rstGC("codgasto")
                .TextMatrix(I, 1) = IIf(IsNull(rstGC("Descripcion")), "", rstGC("Descripcion"))
                .TextMatrix(I, 2) = Format(rstGC("Total"), "##,##0.00")
                .Col = 3: .Row = I
                If rstGC!Alicuota Then
                    Set .CellPicture = imgFactura(0)
                Else
                    Set .CellPicture = imgFactura(1)
                End If
                .CellPictureAlignment = flexAlignCenterCenter
                .Col = 4: .Row = I
                For E = LBound(vecFondo) To UBound(vecFondo)
                    If rstGC!codGasto = vecFondo(E) Then
                        Set .CellPicture = imgFactura(0)
                        Exit For
                    Else
                        Set .CellPicture = imgFactura(1)
                    End If
                Next
                
                .CellPictureAlignment = flexAlignCenterCenter
                curMontoTotal = curMontoTotal + .TextMatrix(I, 2)
            End If
            rstGC.MoveNext
            I = I + 1
        Loop
        .Visible = True
        rstGC.Close
    End With
    Set rstGC = Nothing
    txtFact = Format(CCur(curMontoTotal), "#,##0.00")
    'Proceso de Asignar Gastos No Comunes
    'vecGNC(4) As Currency
    lblFactura(0) = "Fin del Proceso."
    lblFactura(1).Width = 10470
    For I = 0 To 1: lblFactura(I).Refresh
    Next
    'Call rtnEtiqueta("Finalizando Procéso...", 10470)
    With flexFactura(1)
        
        .Rows = rstGNC.RecordCount + 1
        I = 1
        .MergeCol(0) = True
        For K = 6 To 10: lblFact(K) = "0,00"
        Next
        K = 1
        Do While Not rstGNC.EOF
            
            If lblFact(0).Tag = Trim(rstGNC("CODAPTO")) Then
                '
                K = K + 1
                If flexFactura(0).Rows - 1 + K = 50 Then
                    MsgBox "Los gastos no comunes del apto. " & rstGNC("CODAPTO") & " exceden e" _
                    & "l límite de 50 líneas del prerecibo", vbInformation, App.ProductName
                End If
                '
            Else
                K = 1
            End If
            .TextMatrix(I, 0) = Trim(rstGNC("codapto"))
            .TextMatrix(I, 1) = rstGNC("codgasto")
            .TextMatrix(I, 2) = rstGNC("concepto")
            .TextMatrix(I, 3) = Format(CCur(rstGNC("monto")), "##,##0.00")
            If rstGNC("CodGasto") = strGastoAdmin(0) Then
                lblFact(10) = Format(CCur(lblFact(10) + CCur(.TextMatrix(I, 3))), "#,##0.00")
            ElseIf rstGNC("CodGAsto") = strCarta(0) Or rstGNC("CodGasto") = strTelegrama(0) Then
                lblFact(9) = Format(CCur(lblFact(9) + CCur(.TextMatrix(I, 3))), "#,##0.00")
            ElseIf rstGNC("CodGasto") = strGastoMora(0) Then
                lblFact(8) = Format(CCur(lblFact(8) + CCur(.TextMatrix(I, 3))), "#,##0.00")
            ElseIf rstGNC("CodGasto") = strGestion(0) Then
                lblFact(7) = Format(CCur(lblFact(7) + CCur(.TextMatrix(I, 3))), "#,##0.00")
            Else
                lblFact(6) = Format(CCur(lblFact(6) + CCur(.TextMatrix(I, 3))), "#,##0.00")
            End If
            lblFact(0).Tag = .TextMatrix(I, 0)
            rstGNC.MoveNext
            I = I + 1
        Loop
        rstGNC.Close
        '
    End With
    Set rstGNC = Nothing
    MousePointer = vbDefault
    'Call rtnver
    For I = 0 To 1: lblFactura(I).Visible = Not lblFactura(I).Visible
    Next
    If gnCta = CUENTA_POTE Then
        If sysCodPro = 0 Or sysCodPro = "" Then
            MsgBox "Este inmueble es administrado bajo la modalidad 'Cuenta Pote', " & vbCrLf & _
            "debe registrar a " & sysEmpresa & " como proveedor. " & vbCrLf & "Luego entre al menú " & _
            "Utilidades -> Datos de la empresa, y registre el código de proveedor de " & sysEmpresa & _
            "." & vbCrLf & "Si el problema persiste, contacte al administrador del sistema.", vbInformation, App.ProductName
            Exit Sub
        End If
    End If
    cmdFactura(1).Enabled = True
    
    
    '
    End Sub
    

'    '---------------------------------------------------------------------------------------------
'    Public Sub rtnEtiqueta(Texto As String, Ancho As Integer)  '
'    '---------------------------------------------------------------------------------------------
'    LblFactura(0) = Texto
'    LblFactura(1).Width = Ancho
'    For i = 0 To 1: LblFactura(i).Refresh
'    Next
'    '
'    End Sub
    
'    '---------------------------------------------------------------------------------------------
'    Private Sub rtnver()    '
'    '---------------------------------------------------------------------------------------------
'    '
'    For i = 0 To 1: lblFactura(i).Visible = Not lblFactura(i).Visible
'    Next
'    '
'    End Sub
    
    
    Private Sub Form_Resize()
    With Frame1
        .Top = (ScaleHeight / 2) - .Height / 2
        .Left = ScaleWidth / 2 - .Width / 2
    End With
    End Sub


     '---------------------------------------------------------------------------------------------
    '
    '   Rutina: Actualiza_Deuda_propietario
    '
    '   Actualiza la deuda de c/ propietario y la cantidad de meses pendientes
    '---------------------------------------------------------------------------------------------
    Public Sub Actualiza_Deuda_Propietario()
    '
    Dim strSQL As String    'Variables locales
    Dim lngAn As Long
    Dim I As Integer
    '
    strSQL = "SELECT CodProp as Propietario,Count(Saldo) as MP, Sum(Saldo) as Deuda FROM Factura " _
    & "WHERE Saldo > 0 GROUP BY CodProp;"
    
    rstPropietario.Open strSQL, cnn, adOpenKeyset, adLockOptimistic, adCmdText
    
    With rstPropietario
        If .RecordCount > 0 Then .MoveFirst
        lngAn = 8000
        Do
            cnn.Execute "UPDATE Propietarios SET Recibos=" & !MP & ",Deuda='" & !Deuda & "' WHE" _
            & "RE Codigo='" & !Propietario & "';"
            '
            lblFactura(0) = "Propietario '" & !Propietario & "'"
            lblFactura(1).Width = lngAn + .AbsolutePosition
            For I = 0 To 1: lblFactura(I).Refresh
            Next
            .MoveNext
        Loop Until .EOF
        .Close
    End With
    'Aplica los abonos a futuro de tenerlos
    strSQL = "SELECT M.InmuebleMovimientoCaja, M.AptoMovimientoCaja as Apto, Sum(A.Mont" _
    & "o) AS Abo FROM MovimientoCaja as M INNER JOIN TDFAbonos as A ON M.IDRecibo = A.I" _
    & "DRecibo GROUP BY M.InmuebleMovimientoCaja, M.AptoMovimientoCaja HAVING M.Inmuebl" _
    & "eMovimientoCaja='" & gcCodInm & "';"
    
    rstPropietario.Open strSQL, cnnConexion, adOpenKeyset, adLockOptimistic, adCmdText
    
    With rstPropietario
        If Not .EOF Then
            .MoveFirst
            Do Until .EOF   'resta a la deuda de c/propietario el abono a futuro
                cnn.Execute "UPDATE Propietarios SET Deuda = Deuda -'" & !Abo & "' WHERE Codigo" _
                & "='" & !apto & "'"
                .MoveNext
            Loop
        End If
        rstPropietario.Close
    End With
    'Actualiza al campo ultima facturacion de la tabla propietarios y la fecha de facturacion
    cnn.Execute "UPDATE Propietarios INNER JOIN Factura ON Propietarios.Codigo = Factura.codpro" _
    & "p SET Propietarios.UltFact = factura.Facturado,Propietarios.FecREg=Date(),Propietarios.U" _
    & "suario='" & gcUsuario & "' WHERE Factura.Periodo=#" & datPeriodo & "#"
    '
    End Sub

    '---------------------------------------------------------------------------------------------
    '
    '
    '---------------------------------------------------------------------------------------------
    Sub Actulizar_DeudaInm()
    '
    Dim curDeuda As Currency    'variables locales
    '
    
    rstPropietario.Open "SELECT SUM(Saldo) FROM Factura;", cnn, adOpenKeyset, adLockOptimistic
    curDeuda = IIf(IsNull(rstPropietario.Fields(0)), 0, rstPropietario.Fields(0))
    rstPropietario.Close
    
    cnnConexion.Execute "UPDATE Inmueble SET Deuda='" & curDeuda & "' WHERE CodInm='" _
    & gcCodInm & "'"
    
    '
    rstPropietario.Open "SELECT Sum(T2.Monto) FROM MovimientoCaja as T1 INNER JOIN TDFAbonos AS" _
    & " T2 ON T1.IDRecibo=T2.IDRecibo WHERE T1.InmuebleMovimientoCaja='" & gcCodInm & "';", _
    cnnConexion, adOpenKeyset, adLockOptimistic
    
    If Not IsNull(rstPropietario.Fields(0)) Then
        curDeuda = rstPropietario.Fields(0)
        cnnConexion.Execute "UPDATE Inmueble SET Deuda=Deuda - '" & curDeuda & "' WHERE CodInm='" _
        & gcCodInm & "'"
    End If
    rstPropietario.Close
    '
    rstPropietario.Open "SELECT Sum(Saldo) FROM Factura WHERE Periodo=#" & datPeriodo & "#", cnn, _
        adOpenKeyset, adLockOptimistic
    curDeuda = IIf(IsNull(rstPropietario.Fields(0)), 0, rstPropietario.Fields(0))
    rstPropietario.Close
    '
    End Sub

    '---------------------------------------------------------------------------------------------
    '
    '   Funtion:    Total_Gasto
    '
    '   Entrada:    Gasto, codigo del gasto
    '
    '   Retorna un valor moneda que representa el total de un determinado gasto
    '   en un periodo
    '---------------------------------------------------------------------------------------------
    Public Function Total_gasto(Gasto As String) As Currency
    '
    Dim strSQL As String    'variables locales
    '
    strSQL = "SELECT SUM(Monto) FROM DetFact WHERE CodGasto='" & Gasto & "' AND Periodo=#" & _
    datPeriodo & "#;"
    rstPropietario.Open strSQL, cnn, adOpenStatic, adLockReadOnly
    If Not IsNull(rstPropietario.Fields(0)) Then
        Total_gasto = CLng(rstPropietario.Fields(0) * 100) / 100
    Else
        Total_gasto = 0
    End If
    'Cierra el ADODB.Recordset y la conexion a los datos
    rstPropietario.Close
    '
    End Function
    

    '---------------------------------------------------------------------------------------------
    '   Funcion:    ftnTitulo
    '
    '   Salida:     Devuelve un valor cadena que representa la descripción de
    '   un determinado codigo de gasto
    '---------------------------------------------------------------------------------------------
    Public Function ftnTitulo(Gasto As String) As String
    '
    Dim strSQL As String    'Variable local
    '
    strSQL = "SELECT Titulo FROM Tgastos WHERE CodGasto ='" & Gasto & "'"
    rstPropietario.Open strSQL, cnn, adOpenKeyset, adLockOptimistic
    ftnTitulo = rstPropietario.Fields("Titulo")
    rstPropietario.Close
    '
    End Function

    '---------------------------------------------------------------------------------------------
    '   Rutina:     asig_GC
    '
    '   Rutina que asigna los gastos Comunes y no comunes: {alícuota - partes iguales}
    '   de TDF AsignaGasto
    '---------------------------------------------------------------------------------------------
    Sub asig_GC(strFact As String, Hora As String)
    '
    Dim strCalc As String   'variables locales
    Dim strAlic As String
    Dim I As Integer
    '
    For I = 0 To 1
        If I = 0 Then   'Por alícuota
            strCalc = "CSng(G.Monto) * P.Alicuota/100"
            strAlic = "True"
        Else    'Por partes iguales
            strCalc = "CSng(G.Monto) /" & intAptos
            strAlic = "False"
        End If
    
        cnn.Execute "INSERT INTO DetFact (FACT, Detalle,Codigo,CodGasto,Periodo,Monto,Fecha,Hora,Us" _
        & "uario) SELECT '" & strFact & "' & FORMAT(P.ID,'000'),G.Descripcion,P.Codigo,G.CodGasto,G" _
        & ".Cargado," & strCalc & ",Date(),'" & Hora & "','" & gcUsuario & "' FROM AsignaGasto a" _
        & "s G, Propietarios as P WHERE P.Codigo<> 'U" & gcCodInm & "' AND G.Comun=True AND " _
        & "G.Alicuota=" & strAlic & " AND Cargado=#" & datPeriodo & "#;"
    '
    Next
    '
    End Sub
    
    '---------------------------------------------------------------------------------------------
    '   Funcion:    Movimientos
    '
    '   Devuelve el total de créditos y débitos
    '---------------------------------------------------------------------------------------------
    Public Function Movimientos(strGasto As String, strCol As String) As Currency
    'Variables locales
    Dim strSQL As String
    Dim rstlocal As New ADODB.Recordset
    Dim Desde As Date
    '
    Desde = Format(datPeriodo, "m/d/yy")
    Desde = Format(DateAdd("M", -1, Desde), "m/d/yy")
    '
    strSQL = "SELECT Fecha FROM MovFondo WHERE Periodo=#" & Desde & "# AND " _
    & "Tipo='FA' AND CodGasto='" & strGasto & "'"
    '
    rstlocal.Open strSQL, cnn, adOpenKeyset, adLockOptimistic, adCmdText
    If Not rstlocal.EOF And Not rstlocal.BOF Then
        Desde = Format(rstlocal("Fecha"), "m/d/yy")
        rstlocal.Close
        '
        strSQL = "SELECT SUM(" & strCol & ") FROM MovFondo WHERE FEcha>#" & Desde & "# AND " _
        & "CodGasto='" & strGasto & "' AND Del=False;"
        '
        rstlocal.Open strSQL, cnn, adOpenStatic, adLockReadOnly
        Movimientos = IIf(IsNull(rstlocal.Fields(0)), 0, rstlocal.Fields(0))
    End If
        '
    rstlocal.Close
    '
    End Function
    
    '---------------------------------------------------------------------------------------------
    '   Rutina: Actualiza_pago_inf
    '
    '   Actualiza la tabla pago_inf y el fondo de reserva del edificio.
    '
    '---------------------------------------------------------------------------------------------
    Sub Actuliza_Pago_Inf()
    ''Variables locales
    Dim curSA@, curCRE@, curDB@, strSQL$, curFondo@, campo$, Saldo@
    '
    strSQL = "SELECT * FROM Tgastos WHERE Fondo=True;"
    '
    rstPropietario.Open strSQL, cnn, adOpenStatic, adLockReadOnly, adCmdText
    '
    With rstPropietario
        
        If Not .EOF Or Not .BOF Then
            .MoveFirst
            
            campo = Format(DateAdd("d", -1, datPeriodo), "d")
            
            If campo = "31" Then campo = 12
            campo = "Saldo" & campo
            Do
                'actualiza el saldo actual y el campo de cierre
                Saldo = Saldo_Actual(!codGasto)
                '
                cnn.Execute "UPDATE Tgastos SET SaldoActual ='" & Saldo & "',Saldo" & _
                Format(datPeriodo, "d") & "='" & Saldo & "' WHERE Codgasto='" & !codGasto & "';"
                '
                curCRE = Movimientos(!codGasto, "HABER")
                curDB = Movimientos(!codGasto, "DEBE")
                curSA = Saldo_Anterior(!codGasto)
                cnn.Execute "INSERT INTO Pago_INF(Periodo,CtaFondo,SA,CR,DB) VALUES(#" & _
                datPeriodo & "#,'" & !codGasto & "','" & curSA & "','" & curCRE & "','" _
                & curDB & "')"
                .MoveNext
            Loop Until .EOF
        End If
    .Close
    End With
    '
    'Asigna el valor de SaldoACtual del Fondo de Reserva a lngFondo
    strSQL = "SELECT SaldoActual FROM Tgastos WHERE CodGasto='" & gcCodFondo & "';"
    rstPropietario.Open strSQL, cnn, adOpenKeyset, adLockOptimistic, adCmdText
    curFondo = IIf(rstPropietario.EOF Or rstPropietario.BOF, 0, rstPropietario.Fields(0))
    rstPropietario.Close
    '
    'Actualiza Fondo en la tabla Inmueble
    cnnConexion.Execute "UPDATE Inmueble SET FondoAct='" & curFondo & "', Usuario='" & _
    gcUsuario & "',Freg=Date() WHERE Codinm='" & gcCodInm & "';"
    '
    End Sub

    Private Sub Form_Unload(Cancel As Integer)
    'variables locales
    On Error Resume Next
    Set rstPropietario = Nothing
    Set cnn = Nothing
    Set rstInmueble = Nothing
    Screen.MousePointer = vbDefault
    '
    End Sub

    '---------------------------------------------------------------------------------------------
    '   Funcion:    Saldo_Actual
    '   Entrada:    Codigo del Gasto(Codigo_Gasto) tipo cadena
    '   Devuelve la diferencia entre la sumatoria del haber y la sumatoria del debe
    '---------------------------------------------------------------------------------------------
    Private Function Saldo_Actual(Codigo_Gasto$) As Currency
    '
    Dim rstSaldoActual As New ADODB.Recordset   'variables locales
    Dim strSQL$
    '
    strSQL = "SELECT Sum(Haber)-Sum(Debe) FROM MovFondo Where CodGasto='" & Codigo_Gasto & _
    "' AND Del=False;"
    '
    With rstSaldoActual
        .Open strSQL, cnn, adOpenStatic, adLockReadOnly
        Saldo_Actual = IIf(IsNull(.Fields(0)), 0, .Fields(0))
        .Close
    End With
    Set rstSaldoActual = Nothing
    '
    End Function

    Private Function Saldo_Anterior(Codigo_Gasto$) As Currency
    Dim rstSaldoActual As New ADODB.Recordset   'variables locales
    Dim strSQL$, datPer As Date
    datPer = Format(datPeriodo, "mm/dd/yy")
    datPer = Format(DateAdd("m", -1, datPer), "mm/dd/yy")
    '
    strSQL = "SELECT Sum(Haber)-Sum(Debe) FROM MovFondo Where CodGasto='" & Codigo_Gasto & _
    "' AND Del=False AND Fecha<=(SELECT TOP 1 Fecha FROM MovFondo WHERE Tipo='FA' AND " _
    & "Periodo=#" & datPer & "#);"
    '
    With rstSaldoActual
        .Open strSQL, cnn, adOpenStatic, adLockReadOnly
        Saldo_Anterior = IIf(IsNull(.Fields(0)), 0, .Fields(0))
        .Close
    End With
    Set rstSaldoActual = Nothing
    End Function
