VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmAquinal 
   Caption         =   "Form1"
   ClientHeight    =   7380
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11880
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7380
   ScaleWidth      =   11880
   Tag             =   "1000|1000|1000|1000|1000|1000|1000|1000|1000|100011000"
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmd 
      Cancel          =   -1  'True
      Caption         =   "Cartas a Residencia"
      Height          =   900
      Index           =   2
      Left            =   11835
      TabIndex        =   5
      Top             =   9705
      Width           =   975
   End
   Begin VB.CommandButton cmd 
      Caption         =   "Imprimir"
      Height          =   900
      Index           =   1
      Left            =   12810
      TabIndex        =   4
      Top             =   9690
      Width           =   975
   End
   Begin VB.CommandButton cmd 
      Caption         =   "Salir"
      Height          =   900
      Index           =   0
      Left            =   13800
      TabIndex        =   2
      Top             =   9675
      Width           =   975
   End
   Begin VB.Frame fra 
      Height          =   6015
      Left            =   270
      TabIndex        =   0
      Top             =   210
      Width           =   11265
      Begin VB.PictureBox pic 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   765
         Left            =   345
         ScaleHeight     =   765
         ScaleWidth      =   4815
         TabIndex        =   3
         Top             =   285
         Width           =   4815
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid Grid 
         Height          =   4440
         Left            =   315
         TabIndex        =   1
         Top             =   1260
         Width           =   10320
         _ExtentX        =   18203
         _ExtentY        =   7832
         _Version        =   393216
         Cols            =   9
         FixedCols       =   0
         BackColorFixed  =   -2147483646
         ForeColorFixed  =   -2147483639
         BackColorBkg    =   -2147483636
         GridColor       =   -2147483633
         WordWrap        =   -1  'True
         AllowBigSelection=   0   'False
         HighLight       =   2
         GridLinesFixed  =   0
         SelectionMode   =   1
         BorderStyle     =   0
         FormatString    =   "Cod.Emp|Nombre y Apellido|Cargo|Sueldo Mensual|Fecha Ingreso|Dias Aguinaldo|Aquinalso s/Ley|d/apro. 2004|d/apro. 2005"
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
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _NumberOfBands  =   1
         _Band(0).Cols   =   9
         _Band(0).GridLinesBand=   0
         _Band(0).TextStyleBand=   0
         _Band(0).TextStyleHeader=   0
      End
   End
End
Attribute VB_Name = "frmAquinal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ultAgui As Long

Private Sub cmd_Click(Index As Integer)
Dim cReport As ctlReport

Select Case Index
    Case 0
        Unload Me
        Set frmAquinal = Nothing
    Case 1  'reporte general
        MousePointer = vbHourglass
        Set cReport = New ctlReport
        cReport.Reporte = gcReport & "nom_agui.rpt"
        cReport.OrigenDatos(0) = gcPath & "\sac.mdb"
        cReport.Formulas(0) = "periodo='" & Year(Date) & "'"
        cReport.Salida = crImpresora
        cReport.Imprimir
        MousePointer = vbDefault
        Set cReport = Nothing
    Case 2  'cartas al banco
        mcTitulo = "Cartas Aguinaldos"
        mcReport = "nom_cartaagui.rpt"
        Call FrmAdmin.Muestra_Formulario(FrmRepFact, "Impresión Cartas Aguinaldos")
'        Set cReport = New ctlReport
'        cReport.Reporte = gcReport & "nom_cartaagui.rpt"
'        cReport.OrigenDatos(0) = gcPath & "\sac.mdb"
'        cReport.OrigenDatos(0) = gcPath & "\sac.mdb"
'        cReport.Salida = crPantalla
'        cReport.ArbolGrupo = True
'        cReport.TituloVentana = "Cartas Aguinaldos"
'        cReport.Imprimir
        
        Set cReport = Nothing
End Select
End Sub

Private Sub Form_Load()
'variables locales
Dim ancho()
Dim rstlocal As ADODB.Recordset
Dim strSQL As String, Inm As String, sNombre As String
Dim SueldoNeto As Double, subTotal As Double
'establece el ancho de las columnas
'tomando en cuenta la resolución del
If Screen.Width / Screen.TwipsPerPixelX >= 1024 Then 'Mayor o igual a 1024 x 780
    ancho = Array(1000, 4000, 2500, 1500, 1200, 1200, 1000, 1000, 1000, 1000)
Else
    ancho = Array(600, 2200, 1500, 1100, 1000, 900, 900, 900, 900)
End If
Me.Caption = "Aguinaldos " & Year(Date)
'consultamos los últimos aguinaldos cancelados
strSQL = "select TOP 1 IDNomina FROM Nom_Inf WHERE Left(IDNomina,1)=3 ORDER BY Fecha DESC"
Set rstlocal = New ADODB.Recordset
rstlocal.CursorLocation = adUseClient

rstlocal.Open strSQL, cnnConexion, adOpenStatic, adLockReadOnly, adCmdText
If Not (rstlocal.EOF And Not rstlocal.BOF) Then
    ultAgui = Right(rstlocal("IDNomina"), 4)
End If
rstlocal.Close
'genera las consultas necesarias para mostrar la información
'strSql = "SELECT Nom_Detalle.CodEmp, Nom_Detalle.Dias_libres FROM Nom_Detalle " & _
'"WHERE Nom_Detalle.IDNom=312" & ultAgui
'Call rtnGenerator(gcPath & "\sac.mdb", strSql, "qdfAguiOld")
'
'For i = 0 To 1000000
'    DoEvents
'Next

With Grid
    .RowHeight(0) = 450
    .ColAlignment(5) = flexAlignCenterCenter
    .ColAlignment(7) = flexAlignCenterCenter
    .ColAlignment(8) = flexAlignCenterCenter
    .ColAlignment(0) = flexAlignCenterCenter
    .TextArray(0) = "Cod." & vbCrLf & "Emp"
    .TextArray(1) = "Nombre y" & vbCrLf & "Apellido"
    .TextArray(2) = "Cargo"
    .TextArray(3) = "Sueldo" & vbCrLf & "Mensual"
    .TextArray(4) = "Fecha" & vbCrLf & "Ingreso"
    .TextArray(5) = "Días" & vbCrLf & "Aguinaldo"
    .TextArray(6) = "Aguinaldo" & vbCrLf & "s/Ley"
    .TextArray(7) = "Días Apro." & vbCrLf & IIf(ultAgui = 0, "--", ultAgui)
    .TextArray(8) = "Días Apro." & vbCrLf & Year(Date)
    
    .FillStyle = flexFillRepeat
    .Row = 0
    .RowSel = 1
    .Col = 0
    .ColSel = .Cols - 1
    
    .CellAlignment = flexAlignCenterCenter
    .FillStyle = flexFillSingle
    .Row = 1
    For I = 0 To .Cols - 1
        .ColWidth(I) = ancho(I)
    Next
End With
rstlocal.Open "qdfAquinaldos", cnnConexion, adOpenStatic, adLockReadOnly, adCmdTable
rstlocal.Sort = "CodInm,CodEmp"

If Not (rstlocal.EOF And rstlocal.BOF) Then
    rstlocal.MoveFirst
    Call rtnLimpiar_Grid(Grid)
    Grid.Rows = rstlocal.RecordCount + 1
    Grid.MergeCells = flexMergeRestrictRows
    I = 0
    
    Do
        DoEvents
        I = I + 1
        
        If Inm <> rstlocal("CodInm") Then
            
            If I > 1 Then
                Grid.AddItem ""
                Grid.TextMatrix(I, 6) = Format(subTotal, "#,##0.00")
                Grid.Row = I
                Grid.Col = 6
                Grid.CellFontBold = True
                'Grid.CellFontBold = False
                Grid.RowHeight(I) = 200
                I = I + 1
                
            End If
            Grid.AddItem ""
            
            Grid.MergeRow(I) = True
            Grid.TextMatrix(I, 0) = rstlocal("CodInm") & " " & rstlocal("Nombre")
            Grid.TextMatrix(I, 1) = rstlocal("CodInm") & " " & rstlocal("Nombre")
            Grid.TextMatrix(I, 2) = rstlocal("CodInm") & " " & rstlocal("Nombre")
            Grid.Row = I
            
            Grid.RowHeight(I) = 250
            Grid.Col = 0
            Grid.ColSel = 3
            Grid.CellAlignment = flexAlignLeftCenter
            Grid.CellFontBold = True
            I = I + 1
            subTotal = 0
        End If
            sNombre = rstlocal("Nombres")
            sNombre = Left(rstlocal("Nombres"), InStr(rstlocal("Nombres"), " "))
            If sNombre = "" Then sNombre = rstlocal("Nombres")
            Grid.RowHeight(I) = IIf(Screen.Width / Screen.TwipsPerPixelX >= 1024, 250, 215)
            Grid.TextMatrix(I, 0) = rstlocal("CodEmp")
            Grid.TextMatrix(I, 1) = rstlocal("Apellidos") & Space(1) & sNombre
            Grid.TextMatrix(I, 2) = rstlocal("NombreCargo")
            SueldoNeto = (rstlocal("Sueldo") * rstlocal("BonoNoc") / 100) + rstlocal("Sueldo")
            Grid.TextMatrix(I, 3) = Format(SueldoNeto, "#,##0.00 ")
            Grid.TextMatrix(I, 4) = rstlocal("Fingreso")
            Grid.TextMatrix(I, 5) = IIf(DateDiff("m", rstlocal("Fingreso"), "30/11/" & Year(Date)) >= 12, 15, 15 / 12 * DateDiff("m", rstlocal("FIngreso"), "30/11/" & Year(Date)))
            Grid.TextMatrix(I, 6) = Format(SueldoNeto / 30 * Grid.TextMatrix(I, 5), "#,##0.00")
            subTotal = subTotal + Grid.TextMatrix(I, 6)
            Grid.TextMatrix(I, 7) = IIf(IsNull(rstlocal("Dias_Libres")), 0, rstlocal("Dias_Libres"))
            Grid.TextMatrix(I, 8) = IIf(IsNull(rstlocal("Dias")), "--", rstlocal("Dias"))
            
            Grid.Col = 0
            Grid.Row = I
            Grid.ColSel = Grid.Cols - 1
            Grid.RowSel = I
            Grid.FillStyle = flexFillRepeat
            Grid.CellFontName = "Arial Narrow"
            Grid.CellFontSize = IIf(Screen.Width / Screen.TwipsPerPixelX >= 1024, 10, 8)
            Grid.FillStyle = flexFillSingle
            Inm = rstlocal("CodInm")
            rstlocal.MoveNext
    Loop Until rstlocal.EOF
    '
End If
Set rstlocal = Nothing
pic.FontSize = 16
pic.FontBold = True
pic.Print "Cálculo Aguinaldos"
pic.FontSize = 12
pic.Print "Período: " & Year(Date)

End Sub

Private Sub Form_Resize()
'variables locales
If FrmAdmin.WindowState <> vbMinimized Then
    Cmd(0).Top = Me.ScaleHeight - Cmd(0).Height - 300
    Cmd(0).Left = Me.ScaleWidth - Cmd(0).Width - 300
    Cmd(1).Top = Cmd(0).Top
    Cmd(2).Top = Cmd(0).Top
    Cmd(1).Left = Cmd(0).Left - Cmd(1).Width
    Cmd(2).Left = Cmd(1).Left - Cmd(2).Width
    Fra.Width = Me.ScaleWidth - 300
    Fra.Left = 150
    Fra.Height = Me.ScaleHeight - Fra.Top - Cmd(0).Width - 300
    Grid.Height = Fra.Height - 200 - Grid.Top
    Grid.Width = Fra.Width - 300
    Grid.Left = 150
    pic.Left = 150
    Grid.Top = pic.Top + pic.Height + 200
    '
End If
'
End Sub

Private Sub facturar()
'variables locales
Dim cnnLocal As ADODB.Connection
Dim rstlocal As ADODB.Recordset
Dim SueldoNeto As Double
Dim strInm As String, strSQL As String, NDoc As String

Set rstlocal = New ADODB.Recordset
rstlocal.Open "qdfAquinaldos", cnnConexion, adOpenKeyset, adLockOptimistic, adCmdTable
If Not (rstlocal.EOF And rstlocal.BOF) Then
    rstlocal.MoveFirst
    Set cnnLocal = New ADODB.Connection
    Do
        
        
        
        'carga los datos a la facturacion
        'NDoc = FrmFactura.FntStrDoc
        SueldoNeto = (rstlocal("Sueldo") * rstlocal("BonoNoc") / 100) + rstlocal("Sueldo") / 30
        strSQL = "INSERT INTO AsignaGasto(Ndoc,CodGasto,Cargado,Descripcion,Fijo,Comun," _
        & "Alicuota,Monto,Usuario,Fecha,Hora) VALUES('AGU2005','190002','01/11/2005'," _
        & "'AGUINALDOS PERSONAL APROBADO POR JUNTA',0,-1,-1,'" & SueldoNeto * rstlocal("Dias") & "','SUPERVISOR',Date(),Time())"
        'strSql = "DELETE * FROM Asignagasto WHERE CodGasto='190002' And Cargado=#11/01/2005#"
        If strInm <> rstlocal("CodInm") Then cnnLocal.Open cnnOLEDB + gcPath + "\" + rstlocal("CodInm") + "\inm.mdb"
        If Not IsNull(rstlocal("Dias")) Then
            If rstlocal("Dias") > 0 Then cnnLocal.Execute strSQL, N
        End If
        strInm = rstlocal("CodInm")
        rstlocal.MoveNext
        If strInm <> rstlocal("CodInm") Then cnnLocal.Close
    Loop Until rstlocal.EOF
End If
rstlocal.Close
Set rstlocal = Nothing
End Sub

Private Sub emitir_carta_banco(IDNomina As Long)
'Dim rstNom As New ADODB.Recordset
'Dim strSql As String, Calc As String
'Dim k As Date, q As Date, J As Date
'Dim Sem As Long
''
'If IDNomina = 0 Then   'asgina valor al id. de la nomina
'    IDNomina = IIf(Day(Date) <= 15, 1, 2) & Format(Month(DateAdd("m", J, Date)), "00") & _
'    Year(Date)
'End If
''
''varifica que la nónima no este ya cerrada
'Set rstNom = cnnConexion.Execute("SELECT * FROM Nom_Inf WHERE IDNomina=" & IDNomina)
''
'If Not rstNom.EOF And Not rstNom.BOF Then
'    rstNom.Close
'    MsgBox "Nómina ya cerrada", vbInformation, App.ProductName
'    Exit Sub
'End If
'rstNom.Close
'Set rstNom = Nothing
'
'k = "01/" & Mid(IDNomina, 2, 2) & "/" & Right(IDNomina, 4)
'q = DateAdd("d", -1, (DateAdd("m", 1, k)))
'
'Calc = 15  'DIAS TRABAJADOS
'
''elimina el calculo anterior de esta nómina
'Call rtnBitacora("Eliminando la información de la Nómina " & IDNomina)
'cnnConexion.Execute "DELETE * FROM Nom_Detalle WHERE IDNom=" & IDNomina
''
'Call rtnBitacora("Ingresando información base Nómina " & IDNomina)
''if DateDiff ("m",{qdfAquinaldos.FIngreso} ,date(year(Today),11,30) ) >= 12 then
''15
''Else
''  15/12 * (DateDiff ("m",{qdfAquinaldos.FIngreso} ,date(year(Today),11,30)))
'
'
'strSql = "INSERT INTO Nom_Detalle (IDNom,CodEmp,Sueldo,Dias_Trab,Dias_NoTrab,Bono_Noc,Bono_" _
'& "Otros,Otras_Asignaciones,SSO,SPF,LPH,Otras_Deducciones,Porc_BonoNoc,Dias_Libres) SELECT " & IDNomina _
'& ",CodEmp,Sueldo,IIf(DateDiff('m', Fingreso, '30/11/" & Right(IDNomina, 4) & "') >= 12, 15, 15 / 12 * DateDiff('m', Fingreso, '30/11/" & Right(IDNomina, 4) & "')),0, 0,0,0,0,0,0," _
'& "0,BonoNoc,Dias FROM qdfAquinaldos"
''
'cnnConexion.Execute strSql, Calc
'
Dim rstNom As ADODB.Recordset

Dim Rep(6) As String
Set rstNom = New ADODB.Recordset

rstNom.Open "qdfNomina_Bco", cnnConexion, adOpenKeyset, adLockOptimistic, adCmdTable
    rstNom.Filter = "(Cuenta='' and Caja<>'88') or (Cuenta = Null and Caja<>'88')"
    '
    'impresion de reportes (cartas al banco y listado de cheques)
    Call RtnProUtility("Imprimiendo pagos en cheques...", 0)
    '
    If Not rstNom.EOF And Not rstNom.BOF Then
        '
        rstNom.MoveFirst
        '
        Do
            
            Call RtnProUtility("Cargando Cpp inm. " & rstNom("CodInm"), rstNom.AbsolutePosition * 6025 / rstNom.RecordCount)
            'GENERA LA FACTURA DE CPP Y ASIGNA EL GASTO PARA SACAR EL CHEQUE
            '
            strD = FrmFactura.FntStrDoc
            '
            subT = "AÑO 2005"
            cnnConexion.Execute "INSERT INTO Cpp(Tipo,Ndoc,Fact,CodProv,Benef,Detalle,Monto,Ivm" _
            & ",Total,FRecep,Fecr,Fven,CodInm,Moneda,Estatus,Usuario,Freg) VALUES('NO','" & _
            strD & "','" & Left(ID, 3) & Right(ID, 2) & _
            Format(rstNom.AbsolutePosition, "00") & "','" & sysCodPro & "','" & rstNom("Name") & "' ,'AGUINALDOS " _
            & " " & rstNom("NombreCargo") & " " & subT & "','" & rstNom("Neto") & "',0,'" & _
            rstNom("Neto") & "',Date(),Date(),'" & DateAdd("d", 30, Date) & "','" & _
            rstNom("CodInm") & "','BS','ASIGNADO','" & gcUsuario & "',Date())"
            '
            'ingresa el cargado
            '
            cnnConexion.Execute "INSERT INTO Cargado(Ndoc,CodGasto,Detalle,Periodo,Monto,Fecha," _
            & "Hora,Usuario) IN '" & gcPath & "\" & rstNom("CodInm") & "\inm.mdb' SELECT '" & _
            strD & "','190002',Titulo,'01/11/2005','" & rstNom("Neto") _
            & "',Date(),Time(),'" & gcUsuario & "' FROM Tgastos IN '" & gcPath & "\" & _
            rstNom("CodInm") & "\inm.mdb' WHERE CodGasto='190002'"
            '
            rstNom.MoveNext
            '
        Loop Until rstNom.EOF
        '
        Call rtnBitacora("Emitiendo listado de cheques")
        'Call clear_Crystal(FrmAdmin.rptReporte)
        'listado de cheque
        Dim crReporte As ctlReport
        Set crReporte = New ctlReport
        crReporte.Reporte = gcReport + "nom_chq.rpt"
        crReporte.OrigenDatos(0) = gcPath & "\sac.mdb"
        crReporte.Formulas(0) = "Subtitulo='" & subT & "'"
        'guarda una copia del reporte
        crReporte.Salida = crArchivoDisco
        crReporte.ArchivoSalida = RepNom & "CH" & ID & ".rpt"
        crReporte.Imprimir
        'impresión en papel
        crReporte.Salida = crImpresora
        crReporte.Imprimir
        Set crReporte = Nothing
'        With FrmAdmin.rptReporte
'            '
'            .ReportFileName = gcReport + "nom_chq.rpt"
'            .DataFiles(0) = gcPath & "\sac.mdb"
'            .Formulas(0) = "Subtitulo='AGUINALDOS " & subT & "'"
'            'guarda una copia del reporte
'            .Destination = crptToFile
'            .PrintFileType = crptCrystal
'            .PrintFileName = RepNom & "CH" & IDNomina & ".rpt"
'            errLocal = .PrintReport
'            '
'            If errLocal <> 0 Then
'                MsgBox "Ocurrio el siguiente error al guardar el reporte de cheques: " & _
'                .LastErrorString, vbCritical, "Error " & .LastErrorNumber
'            End If
'            '
'            .Destination = crptToPrinter
'            errLocal = .PrintReport
'            If errLocal <> 0 Then
'            '
'                MsgBox "Ocurrio el siguiente error al imprimir el reporte de cheques: " & _
'                .LastErrorString, vbCritical, "Error " & .LastErrorNumber
'                '
'            End If
'            '
'        End With
        '
    End If
    '
    
    'Call clear_Crystal(FrmAdmin.rptReporte)
    'cartas al banco
    Call rtnBitacora("Emitiendo cartas al banco")
    rstNom.Filter = "Cuenta<>'' and CodInm<>'8888'"
    rstNom.Sort = "Caja DESC,CodInm DESC"
    Acredita = "29/11/2005"
    '
    If Not rstNom.EOF And Not rstNom.BOF Then
        '
        rstNom.MoveFirst: Inm = rstNom("Caja")

        '
        rstCG.Open "SELECT Cuentas.NumCuenta, Bancos.NombreBanco, Bancos.Agencia, Banco" _
        & "s.Contacto, Cuentas.IDCuenta,Cuentas.Titular FROM Bancos INNER JOIN Cuentas ON " _
        & "Bancos.IDBanco = Cuentas.IDBanco WHERE Cuentas.IDCuenta=" & rstNom("CtaInm"), _
        cnnOLEDB + gcPath + IIf(rstNom("Caja") = sysCodCaja, "\" + sysCodInm + "\", "\" & _
        rstNom("CodInm") & "\") + "inm.mdb", adOpenKeyset, adLockOptimistic, adCmdText
        '
        Rep(0) = IIf(IsNull(rstCG("Agencia")), "", rstCG("Agencia"))
        Rep(1) = IIf(IsNull(rstCG("Contacto")), "", rstCG("Contacto"))
        Rep(2) = IIf(IsNull(rstCG("NombreBanco")), "", rstCG("NombreBanco"))
        Rep(3) = IIf(IsNull(rstCG("NumCuenta")), "", rstCG("NumCuenta"))
        Rep(4) = IIf(IsNull(rstCG("Titular")), "", rstCG("Titular"))
        Rep(5) = IIf(IsNull(rstNom("Caja")), "", rstNom("Caja"))
        Rep(6) = 0

        numToLetras.Moneda = "Bs."
        rstCG.Close
        '
        Do
        
            Call RtnProUtility("Imprimiendo carta al banco inm. '" & rstNom("CodInm") & "'", _
            rstNom.AbsolutePosition * 6025 / rstNom.RecordCount)
                
            If Inm = rstNom("Caja") Then Rep(6) = CCur(Rep(6)) + rstNom("Neto")
            '
            If Inm <> rstNom("Caja") Then
                
                Call rtnBitacora("Imprimiendo carta " & Inm)
                '
100             numToLetras.Numero = CCur(Rep(6))
                Set crReporte = New ctlReport
                crReporte.Reporte = gcReport + "nom_bco.rpt"
                crReporte.OrigenDatos(0) = gcPath & "\sac.mdb"
                crReporte.Formulas(0) = "agencia='Agencia " & Rep(0) & "'"
                crReporte.Formulas(1) = "aletras='" & UCase(numToLetras.ALetra) & "'"
                crReporte.Formulas(2) = "atencion='Atencion: " & Rep(1) & "'"
                crReporte.Formulas(3) = "banco='Banco " & Rep(2) & "'"
                crReporte.Formulas(4) = "cuenta='" & Rep(3) & "'"
                crReporte.Formulas(5) = "efectivo='" & Acredita & "'"
                crReporte.Formulas(6) = "inmueble='" & Rep(4) & "'"
                crReporte.Formulas(7) = "quincena='" & subT & "'"
                crReporte.Formulas(8) = "param='" & Rep(5) & "'"
                'guarda copia local
                crReporte.Salida = crArchivoDisco
                crReporte.ArchivoSalida = RepNom & "CB" & Inm & ID & ".rpt"
                crReporte.Imprimir
                'enviar a impresora
                crReporte.Salida = crImpresora
                crReporte.Imprimir 2
                Set crReporte = Nothing
                
'                With FrmAdmin.rptReporte
'                    '
'                    .ReportFileName = gcReport + "nom_bco.rpt"
'                    .DataFiles(0) = gcPath & "\sac.mdb"
'                    .Formulas(0) = "agencia='Agencia " & Rep(0) & "'"
'                    .Formulas(1) = "aletras='" & UCase(numToLetras.ALetra) & "'"
'                    .Formulas(2) = "atencion='Atencion: " & Rep(1) & "'"
'                    .Formulas(3) = "banco='Banco " & Rep(2) & "'"
'                    .Formulas(4) = "cuenta='" & Rep(3) & "'"
'                    .Formulas(5) = "efectivo='" & Acredita & "'"
'                    .Formulas(6) = "inmueble='" & Rep(4) & "'"
'                    .Formulas(7) = "quincena='AGUINALDOS " & subT & "'"
'                    .Formulas(8) = "param='" & Rep(5) & "'"
'                    'guarda una copia del reporte
'                    .Destination = crptToFile
'                    .PrintFileName = RepNom & "CB" & Inm & ID & ".rpt"
'                    .PrintFileType = crptCrystal
'                    errLocal = .PrintReport
'                    If errLocal <> 0 Then
'
'                        MsgBox "Ocurrio el siguiente error al guardar la carta al banco de la cuenta" _
'                        & " [ " & Inm & " ] " & .LastErrorString, vbCritical, "Error " & _
'                        .LastErrorNumber
'
'                    End If
'                    .CopiesToPrinter = 2
'                    .Destination = crptToPrinter
'                    errLocal = .PrintReport
'                    '
'                    If errLocal <> 0 Then MsgBox .LastErrorString, vbCritical, "Error " & .LastErrorNumber
'                    '
'                End With
                '
                If rstNom.EOF Then Exit Do
                '
                Inm = rstNom("Caja")
                rstCG.Open "SELECT Cuentas.NumCuenta, Bancos.NombreBanco, Bancos.Agencia, Banco" _
                & "s.Contacto, Cuentas.IDCuenta FROM Bancos INNER JOIN Cuentas ON Bancos.IDBanc" _
                & "o = Cuentas.IDBanco WHERE Cuentas.IDCuenta=" & rstNom("CtaInm"), cnnOLEDB + _
                gcPath + IIf(rstNom("Caja") = sysCodCaja, "\" + sysCodInm + "\", "\" & _
                rstNom("CodInm") & "\") + "inm.mdb", adOpenKeyset, adLockOptimistic, adCmdText
            '
                Rep(0) = IIf(IsNull(rstCG("Agencia")), "", rstCG("Agencia"))
                Rep(1) = IIf(IsNull(rstCG("Contacto")), "", rstCG("Contacto"))
                Rep(2) = IIf(IsNull(rstCG("NombreBanco")), "", rstCG("NombreBanco"))
                Rep(3) = IIf(IsNull(rstCG("NumCuenta")), "", rstCG("NumCuenta"))
                Rep(4) = IIf(IsNull(rstNom("Nombre")), "", rstNom("Nombre"))
                Rep(5) = IIf(IsNull(rstNom("Caja")), "", rstNom("Caja"))
                Rep(6) = rstNom("Neto")
                numToLetras.Moneda = "Bs."
                rstCG.Close
                '
            End If
            '
            rstNom.MoveNext
            If rstNom.EOF Then GoTo 100
        Loop Until rstNom.EOF
        'imprime el último reporte
        '
    End If
    rstNom.Close
    Set rstNom = Nothing
End Sub
