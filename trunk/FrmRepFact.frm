VERSION 5.00
Begin VB.Form FrmRepFact 
   Caption         =   "Reportes de Facturación"
   ClientHeight    =   2715
   ClientLeft      =   60
   ClientTop       =   420
   ClientWidth     =   4095
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   2715
   ScaleWidth      =   4095
   Begin VB.CommandButton cmdRfact 
      Caption         =   "Aceptar"
      Height          =   765
      Index           =   0
      Left            =   855
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1725
      Width           =   1005
   End
   Begin VB.CommandButton cmdRfact 
      Cancel          =   -1  'True
      Caption         =   "Salir"
      Height          =   765
      Index           =   1
      Left            =   2190
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1725
      Width           =   1005
   End
   Begin VB.Frame fraRfact 
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
      Left            =   330
      TabIndex        =   0
      Top             =   255
      Width           =   3405
      Begin VB.ComboBox cmbRfact 
         DataField       =   "TipoMovimientoCaja"
         Height          =   315
         Index           =   0
         ItemData        =   "FrmRepFact.frx":0000
         Left            =   465
         List            =   "FrmRepFact.frx":002B
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   435
         Width           =   1320
      End
      Begin VB.ComboBox cmbRfact 
         DataField       =   "TipoMovimientoCaja"
         Height          =   315
         Index           =   1
         ItemData        =   "FrmRepFact.frx":0094
         Left            =   1875
         List            =   "FrmRepFact.frx":0096
         Sorted          =   -1  'True
         TabIndex        =   2
         Text            =   "cmbRfact"
         Top             =   435
         Width           =   1020
      End
   End
End
Attribute VB_Name = "FrmRepFact"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
    '----------------------------------------------------------
    'Modulo Facturación Impresion de Reportes: Aviso de Cobro'-
    'Pre-Recibo/Analis de Facturación/Reporte Facturacion    '-
    '17/12/2002------------------------------------------------
    Dim vecGasto(2 To 5) As String
    Dim VecInm() As String
    'Public eMail As Boolean
    
    Private Sub cmbRfact_KeyPress(Index As Integer, KeyAscii As Integer)
    On Error Resume Next
    If Index = 1 Then Call Validacion(KeyAscii, "0123456789")
    If KeyAscii = 13 And Index = 0 Then cmbRfact(1).SetFocus
    If KeyAscii = 13 And Index = 1 Then cmdRfact(0).SetFocus
    End Sub

    Private Sub cmdRfact_Click(Index As Integer)
    'Variables locales
    Dim ErrorLocal&, Periodo$, Reporte$, strPregunta$
    Dim datPer1 As Date
    Dim FechaI As Date, FechaF As Date
    '
    Select Case Index
    '
        Case 0  'Botón imprimir
        '--------------------
        If mcReport = "nom_cartaagui.rpt" Then
            Periodo = cmbRfact(0) & "/" & cmbRfact(1)
            imprimir_Carta_Aguinaldo (Periodo)
            Exit Sub
        End If
        datPer1 = "01/" & cmbRfact(0) & "/" & cmbRfact(1)
        datPer1 = Format(datPer1, "mm/dd/yy")
        Periodo = Left(cmbRfact(0), 3) & "-" & cmbRfact(1)
        
        If strLlamada = "F" Then
            If cmbRfact(0) = "" Then
                MsgBox "No ha seleccionado ningún mes de la lista..", vbCritical, App.ProductName
                Exit Sub
            ElseIf cmbRfact(1) = "" Then
                MsgBox "No ha seleccionado ningún año de la lista..", vbCritical, App.ProductName
                Exit Sub
            End If
            
            If mcReport <> "GM" And mcReport <> "IVA" Then If Validar_Periodo(datPer1) Then Exit Sub
            MousePointer = vbHourglass
            
            Select Case mcReport
            '
                Case "fact_Aviso.rpt"
                '------------
                    Unload Me
                    Call rtnAC(Periodo, datPer1, crPantalla)
                    'MousePointer = vbDefault
                    
                Case "PreRecibo.Rpt"
                '------------
                    Call rtnBitacora("Print Pre-Recibo Periodo: " & Periodo & " Inm.:" & gcCodInm)
                    Call Printer_PreRecibo(Periodo, datPer1, , crPantalla)
                    
                Case "fact_mes.Rpt"
                '------------
                    Call rtnBitacora("Print Facturación Mensual Periodo: " & Periodo & " Inm.:" _
                    & gcCodInm)
                    Call Print_RepFact(datPer1)
                    
                Case "fact_analisis.rpt"
                '------------
                    Call rtnBitacora("Print Análisis de Facturación Periodo: " & Periodo & _
                    " Inm.:" & gcCodInm)
                    Call Print_AnaFact(datPer1)
                    
                Case "fact_GNC.rpt"
                '------------
                    Call rtnBitacora("Print Gastos No Comunes Periodo: " & Periodo & " Inm.:" _
                    & gcCodInm)
                    Call Printer_GNC(Periodo, datPer1, crPantalla)
                    '
                Case "Fact_control.rpt"
                '------------
                    Call rtnBitacora("Print Control de Facturación Periodo: " & Periodo & _
                    " Inm.:" & gcCodInm)
                    Call Printer_Control_Facturacion(Periodo, datPer1, crPantalla)
                    
                Case "PAQUETECOMPLETO"
                '------------
                    Call rtnMakeQuery(datPer1, vecGasto(2), vecGasto(3), vecGasto(4), vecGasto(5))
                    Call Printer_PaqueteCompleto(datPer1, Deuda_Total(datPer1), crPantalla)
                    
                Case "aviso_email"
                '------------
                    Call Enviar_ACemail(datPer1)
                
                Case "GM"
                '------------
                    Call FrmAdmin.Reporte_GM(CStr(datPer1), crPantalla, Guarda_Copia:=False)
                
                Case "IVA"
                    datPer1 = Format(datPer1, "mm/dd/yy")
                    FechaI = datPer1
                    FechaF = DateAdd("d", -1, DateAdd("m", 1, FechaI))
                    FechaI = Format(FechaI, "mm/dd/yy")
                    FechaF = Format(FechaF, "mm/dd/yy")
                    
                    Reporte = "SELECT Cpp.*,Inmueble.* FROM Cpp INNER JOIN Inmueble ON Cpp.CodInm = Inmueble.CodInm WHERE Cpp.FRecep Between #" & FechaI & "#" _
                    & " AND #" & FechaF & "# and Cpp.Fact='F' & Cpp.CodInm ORDER BY  Cpp.CodInm"
                    
                    Call Printer_Report(Reporte, "Libro de Ventas (I.V.A.)", "Mes: " & Format(datPer1, "mm/yyyy"), True)
                    
                    '
            End Select

        ElseIf strLlamada = "R" Then
            
            strPregunta = "¿Desea reversar la facturación del período seleccionado...?"
            If Respuesta(strPregunta) Then
                Dim revFacturacion As New stRFactura
                Dim rpReporte As ctlReport
                MousePointer = vbHourglass
                revFacturacion.Codigo_Inmueble = gcCodInm
                revFacturacion.Periodo = datPer1
                revFacturacion.sConexionP = gcPath & "\sac.mdb"
                revFacturacion.sConexionS = mcDatos
                revFacturacion.Ruta = gcPath & gcUbica & "Reportes\"
                revFacturacion.Reporte = Left(cmbRfact(0), 3) & Right(cmbRfact(1), 2) & ".rpt"
                Call Print_RepFact(datPer1)
                
                If Not revFacturacion Then
                    '
                    Set rpReporte = New ctlReport
                    With rpReporte
                        .Reporte = gcPath & gcUbica & "Reportes\R" & Left(cmbRfact(0), 3) & _
                        Right(cmbRfact(1), 2) & ".rpt"
                        .Salida = crPantalla
                        .Imprimir
                        If Err <> 0 Then
                            Kill gcPath & gcUbica & "Reportes\F" & Left(cmbRfact(0), 3) & _
                            Right(cmbRfact(1), 2) & ".rpt"
                        End If
                    End With
                    Call rtnBitacora("Facturación Reversada")
                    MsgBox "Facturación Reversada....", vbInformation, App.ProductName
                Else
                    Call rtnBitacora("Mensajes durante el reverso de la facturación")
                End If
                Set revFacturacion = Nothing
            End If
            MousePointer = vbDefault
            
        End If
        MousePointer = vbDefault
        
        Case 1  'botón salir
        '--------------------
            Unload Me
            Set FrmRepFact = Nothing
            '
    End Select
    If ErrorLocal <> 0 Then MsgBox Err.Description, vbCritical, "Error " & ErrorLocal
    '
    End Sub


    Private Sub Form_Load()
    '
    CenterForm Me
    Me.Caption = mcTitulo
    cmdRfact(0).Picture = LoadResPicture("OK", vbResIcon)
    cmdRfact(1).Picture = LoadResPicture("SALIR", vbResIcon)
'    For i = 0 To 1: cmbRfact(1).AddItem (Year(Date) + i)
'    Next
    For I = 0 To Year(Date) - 2002: cmbRfact(1).AddItem (2003 + I)
    Next
    'Presenta el periodo al mes actual
    cmbRfact(0).Text = cmbRfact(0).List(Month(Date) - 1)
    cmbRfact(1).Text = Year(Date)
    '
    With FrmAdmin.objRst
    
        If .State = 0 Then .Open
        .MoveFirst
        .Find "CodInm='" & gcCodInm & "'"
        If .EOF Then .MoveLast
        vecGasto(2) = "='" & !CodGastoAdmin & "'"
        vecGasto(3) = "='" & !CodIntMora & "'"
        vecGasto(4) = "='" & !CodGestion & "'"
        vecGasto(5) = "NOT IN ('" & !CodGastoAdmin & "','" & !CodIntMora & "','" _
        & !CodGestion & "')"
        
    End With
    '
    End Sub

    '---------------------------------------------------------------------------------------------
    '   Rutina:     Print_RepFact
    '
    '   Entradas:   datPeriodo, Mes del que se desea la información
    '
    '   Genera la consulta necesaria para la re-impresión del reporte
    '   Reporte de Facturación
    '---------------------------------------------------------------------------------------------
    Private Sub Print_RepFact(datPeriodo As Date)
    '
    Dim strSQL As String    'Variables locales
    '
    If strLlamada = "F" Then
        Call rtnMakeQuery(datPeriodo, vecGasto(2), vecGasto(3), vecGasto(4), vecGasto(5))
    End If
    strSQL = UCase(Left(cmbRfact(0), 3)) & "-" & cmbRfact(1)
    Call Printer_Facturacion_Mensual(strSQL, crPantalla)
    '
    End Sub
    
    '---------------------------------------------------------------------------------------------
    Private Function qdf_RFact(strCodGasto As String, DatFecha As Date) As String  '
    '---------------------------------------------------------------------------------------------
    qdf_RFact = "SELECT CodApto, Periodo, Sum(Monto) AS GNC FROM GastoNoComun WHERE CodGasto " _
    & strCodGasto & " AND Periodo=#" & DatFecha & "# AND CodApto <> 'U" & gcCodInm & "' GROUP B" _
    & "Y CodApto, Periodo UNION SELECT Codigo,'" & Format(DatFecha, "mm/dd/yyyy") & "' as P, 0 " _
    & "as G FROM Propietarios WHERE Codigo Not In (SELECT CodApto FROM GastoNoComun WHERE CodG" _
    & "asto " & strCodGasto & " AND Periodo=#" & DatFecha & "#) AND Codigo <> 'U" & gcCodInm & "'"
    '
    End Function
    
    '---------------------------------------------------------------------------------------------
    '   Rutina:     Print_AnaFact
    '
    '   Entradas:   datPeriodo, Mes del que se desea la información
    '
    '   Genera la consulta necesaria para la re-impresión del reporte
    '   análisis de facturación
    '---------------------------------------------------------------------------------------------
    Private Sub Print_AnaFact(datPeriodo As Date)
    '
    Dim strGNC As String
    Dim curT As Currency
    
    strGNC = "SELECT * FROM GastoNoComun WHERE Periodo=#" & datPeriodo & "# UNION  SELECT 0,''," _
    & "'','',0,#" & datPeriodo & "#,DATE(),TIME(),'" & gcUsuario & "' FROM GastoNoComun ORDER B" _
    & "Y CodApto, CodGasto;"

    '
    'Crea una consulta Gastos No Comunes
    Call rtnGenerator(mcDatos, strGNC, "qdfGNC")
    strGNC = UCase(Format(Format(datPeriodo, "mm/dd/yyyy"), "mmm - yyyy"))
    curT = Deuda_Total(datPeriodo)
    Call Printer_Analisis_Facturacion(strGNC, datPeriodo, curT, crPantalla)
    '
    End Sub

    
    '---------------------------------------------------------------------------------------------
    '   Funcion:    Validar_Periodo
    '
    '   Devuelve True si el periodo seleccionado no ha sido facturado aun
    '---------------------------------------------------------------------------------------------
    Private Function Validar_Periodo(datPeriodo As Date) As Boolean
    'variables locales
    Dim rstValida As New ADODB.Recordset
    Dim cnnValida As New ADODB.Connection
    '
    cnnValida.Open cnnOLEDB & mcDatos
    rstValida.Open "SELECT * FROM Factura WHERE Fact Not LIKE 'CH*' AND Periodo=#" & datPeriodo _
    & "#;", cnnValida, adOpenStatic, adLockReadOnly
    If rstValida.RecordCount <= 1 Then
        Validar_Periodo = MsgBox("No se tiene información del periodo " _
        & UCase(Left(cmbRfact(0), 3)) & "-" & cmbRfact(1), vbInformation + vbOKOnly)
    End If
    '
    Set rstValida = Nothing
    Set cnnValida = Nothing
    '
    End Function
    
