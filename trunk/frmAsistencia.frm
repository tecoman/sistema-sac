VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmAsistencia 
   AutoRedraw      =   -1  'True
   Caption         =   "Asistencia Personal"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   WindowState     =   2  'Maximized
   Begin VB.Frame fra 
      Caption         =   "Otras Opciones: "
      Height          =   3435
      Index           =   1
      Left            =   165
      TabIndex        =   12
      Top             =   2670
      Width           =   4440
      Begin VB.CheckBox chkAsistencia 
         Caption         =   "RETARDOS"
         Height          =   465
         Index           =   1
         Left            =   180
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   2100
         Width           =   1965
      End
      Begin VB.CheckBox chkAsistencia 
         Caption         =   "INASISTENCIAS"
         Height          =   465
         Index           =   2
         Left            =   2280
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   2100
         Width           =   1965
      End
      Begin VB.OptionButton optAsistencia 
         Caption         =   "Impresora"
         Height          =   375
         Index           =   1
         Left            =   1320
         TabIndex        =   15
         Top             =   840
         Width           =   1215
      End
      Begin VB.OptionButton optAsistencia 
         Caption         =   "Ventana"
         Height          =   375
         Index           =   0
         Left            =   1320
         TabIndex        =   14
         Top             =   480
         Value           =   -1  'True
         Width           =   1215
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         Index           =   1
         X1              =   195
         X2              =   4245
         Y1              =   1560
         Y2              =   1560
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00808080&
         BorderWidth     =   2
         Index           =   0
         X1              =   195
         X2              =   4245
         Y1              =   1560
         Y2              =   1560
      End
      Begin VB.Label lbl 
         Caption         =   "Imprimir en:"
         Height          =   330
         Index           =   2
         Left            =   240
         TabIndex        =   13
         Top             =   480
         Width           =   1215
      End
   End
   Begin VB.Frame fra 
      Height          =   1590
      Index           =   0
      Left            =   165
      TabIndex        =   4
      Top             =   6300
      Width           =   4470
      Begin VB.CommandButton cmd 
         Caption         =   "&Actualizar"
         Height          =   1005
         Index           =   2
         Left            =   270
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   375
         Width           =   1215
      End
      Begin VB.CommandButton cmd 
         Caption         =   "&Imprimir"
         Height          =   1005
         Index           =   1
         Left            =   1680
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   375
         Width           =   1215
      End
      Begin VB.CommandButton cmd 
         Caption         =   "&Salir"
         Height          =   1005
         Index           =   0
         Left            =   3030
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   375
         Width           =   1215
      End
   End
   Begin VB.Frame fraAsistencia 
      Caption         =   "Filtro: "
      Height          =   2055
      Left            =   135
      TabIndex        =   1
      Top             =   360
      Width           =   4470
      Begin VB.CheckBox chkAsistencia 
         Alignment       =   1  'Right Justify
         Caption         =   "Usuario:"
         Height          =   330
         Index           =   0
         Left            =   210
         TabIndex        =   3
         Top             =   1492
         Width           =   1800
      End
      Begin MSComCtl2.DTPicker dtpAsistencia 
         Height          =   330
         Index           =   0
         Left            =   2100
         TabIndex        =   9
         Top             =   480
         Width           =   1560
         _ExtentX        =   2752
         _ExtentY        =   582
         _Version        =   393216
         CalendarTitleBackColor=   -2147483646
         CalendarTitleForeColor=   -2147483639
         CheckBox        =   -1  'True
         Format          =   54067201
         CurrentDate     =   37943
      End
      Begin MSDataListLib.DataCombo dtcAsistencia 
         Bindings        =   "frmAsistencia.frx":0000
         Height          =   315
         Left            =   2100
         TabIndex        =   10
         Top             =   1500
         Width           =   2070
         _ExtentX        =   3651
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         ListField       =   "NombreUsuario"
         Text            =   ""
      End
      Begin MSComCtl2.DTPicker dtpAsistencia 
         Height          =   330
         Index           =   1
         Left            =   2085
         TabIndex        =   11
         Top             =   990
         Width           =   1560
         _ExtentX        =   2752
         _ExtentY        =   582
         _Version        =   393216
         CalendarTitleBackColor=   -2147483646
         CalendarTitleForeColor=   -2147483639
         CheckBox        =   -1  'True
         Format          =   54067201
         CurrentDate     =   37943
      End
      Begin VB.Label lbl 
         Caption         =   "Fecha Final:"
         Height          =   330
         Index           =   1
         Left            =   255
         TabIndex        =   8
         Top             =   990
         Width           =   1215
      End
      Begin VB.Label lbl 
         Caption         =   "Fecha Inicial:"
         Height          =   330
         Index           =   0
         Left            =   270
         TabIndex        =   2
         Top             =   480
         Width           =   1215
      End
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grid 
      CausesValidation=   0   'False
      Height          =   7530
      Left            =   4860
      TabIndex        =   0
      Tag             =   "1000|1800|1200|1200|1200|0"
      Top             =   390
      Width           =   6795
      _ExtentX        =   11986
      _ExtentY        =   13282
      _Version        =   393216
      Cols            =   6
      FixedCols       =   0
      BackColorFixed  =   -2147483646
      ForeColorFixed  =   -2147483639
      BackColorBkg    =   -2147483636
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483633
      WordWrap        =   -1  'True
      FocusRect       =   2
      HighLight       =   0
      MergeCells      =   3
      AllowUserResizing=   1
      FormatString    =   "^Fecha|Usuario|Entrada|Salida|Duración|"
      BandDisplay     =   1
      _NumberOfBands  =   1
      _Band(0).Cols   =   6
   End
End
Attribute VB_Name = "frmAsistencia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
    Dim rstAsistencia(1) As New ADODB.Recordset
    
    Private Sub chkAsistencia_Click(Index As Integer)
    'variables locales
    Call Aplicar_filtro
    End Sub
    

    Private Sub Cmd_Click(Index As Integer)
    '
    Select Case Index
        Case 0: Unload Me
        Case 1: Call Print_report
        Case 2
            Set rstAsistencia(0) = ModGeneral.ejecutar_procedure("procasistencia", Me.dtpAsistencia(0), Me.dtpAsistencia(1))
            Call Aplicar_filtro
    End Select
    '
    End Sub

    Private Sub dtcAsistencia_Click(Area As Integer)
    If Area = 2 Then Call Aplicar_filtro
    End Sub

    Private Sub dtpAsistencia_Click(Index As Integer): Call Aplicar_filtro
    End Sub

    Private Sub Form_Load()
    '
    Call generar_procedure
    dtpAsistencia(0).Value = Date
    dtpAsistencia(1).Value = Date
    cmd(0).Picture = LoadResPicture("SALIR", vbResIcon)
    cmd(1).Picture = LoadResPicture("PRINT", vbResIcon)
    '
    'configura la presentación del grid
    Grid.MergeCol(0) = True
    Set rstAsistencia(0) = ModGeneral.ejecutar_procedure("procasistencia", Me.dtpAsistencia(0), Me.dtpAsistencia(1))
'    rstAsistencia(0).Open "SELECT * FROM qdfasistencia WHERE Entrada<>#12/30/1899# and Salid" _
'    & "a<>#12/30/1899# ORDER BY Fecha,entrada", cnnConexion, adOpenKeyset, adLockOptimistic, _
'    adCmdText
    Call Aplicar_filtro
    rstAsistencia(1).Open "SELECT * FROM USUARIOS WHERE Nivel > 1;", cnnOLEDB + gcPath + _
    "\tablas.mdb", adOpenKeyset, adLockOptimistic, adCmdText
    Set dtcAsistencia.RowSource = rstAsistencia(1)
    '
    End Sub

    Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    For I = 0 To 1
        rstAsistencia(I).Close
        Set rstAsistencia(I) = Nothing
    Next
    End Sub

    Sub Aplicar_filtro()
    'variables locales
    Dim strFiltro As String, strUser As String
    Dim HEntrada As Date
    Dim HSalida As Date
    Dim F As Date
    Dim rstHora As New ADODB.Recordset
    Const FF = "m/d/yyyy"
    Const FC = "m/d/yyyy h:m:s"
    '
    'obtiene el horario de trabajo de la empresa
    rstHora.Open "Ambiente", cnnConexion, adOpenKeyset, adLockOptimistic, adCmdTable
    HEntrada = rstHora("HEntrada")
    HSalida = rstHora("HSalida")
    rstHora.Close
    Set rstHora = Nothing
    '------------------------
    Set Grid.DataSource = Nothing
    
    strFiltro = ""
    strUser = ""
    
    If dtcAsistencia <> "" And chkAsistencia(0) Then
        strUser = " AND Usuario ='" & dtcAsistencia & "'"
    End If
    'muestra los retardos
    If chkAsistencia(1) Then    'filtrar retardos
        
        strFiltro = "(Entrada>#" & HEntrada & "# and Entrada<#" & _
        HSalida & "#" & strUser & ") OR (Salida<#" & _
        HSalida & "# and Salida>#" & HEntrada & "#" & strUser & ")"
        
    End If
    'muestra las inasistencias
    If chkAsistencia(2) Then
        strFiltro = strFiltro & IIf(strFiltro = "", "", " OR ") & _
        "((Entrada=#12:00:00 a.m.#" & strUser & _
        ") OR (Salida=#12:00:00 a.m.#" & strUser & "))"
    End If
    If strUser <> "" And strFiltro = "" Then strFiltro = Replace(strUser, " AND ", "")
    
    rstAsistencia(0).Filter = strFiltro
    
    If rstAsistencia(0).RecordCount <= 0 Then
        Grid.Rows = 2
        Call rtnLimpiar_Grid(Grid)
    Else
        '
        Grid.Rows = rstAsistencia(0).RecordCount + 1
        With rstAsistencia(0)
            If Not (.EOF And .BOF) Then
                .MoveFirst: I = 1
                MousePointer = vbHourglass
                Do
                    For j = 0 To 4
                        Grid.TextMatrix(I, j) = .Fields(j)
                    Next
                    Grid.Row = I
                    Grid.Col = 2
                    Grid.CellBackColor = vbWhite
                    If CDate(!Entrada) > HEntrada Then Grid.CellBackColor = &HFFFF80
                    Grid.Col = 3
                    Grid.CellBackColor = vbWhite
                    If CDate(!Salida) < HSalida Then Grid.CellBackColor = &HFFFF80
                    Grid.TopRow = IIf((I - 28) + 1 < 1, 1, (I - 28) + 1)
                    'DoEvents
                    I = I + 1
                    .MoveNext
                Loop Until .EOF
                MousePointer = vbDefault
            End If
        End With
'        Set grid.Recordset = rstAsistencia(0)
'        grid.Rows = grid.Recordset.RecordCount + 1
    End If
    Set Grid.FontFixed = LetraTitulo(LoadResString(527), 9, , True)
    Set Grid.Font = LetraTitulo(LoadResString(528), 8)
    Call centra_titulo(Grid, True)
    '
    'On Error Resume Next
'    With grid
'        .Visible = False
'        For i = 1 To .Rows - 1
'            .Row = i
'            .Col = 2
'            .CellBackColor = vbWhite
'            If CDate(.Text) > HEntrada Then .CellBackColor = &HFFFF80
'            .Col = 3
'            .CellBackColor = vbWhite
'            If CDate(.Text) < HSalida And CDate(.TextMatrix(i, 0)) < Date Then .CellBackColor = _
'            &HFFFF80
'        Next i
'        .Visible = True
'    End With
    
    '
    End Sub

    '---------------------------------------------------------------------------------------------
    '   Rutina: Print_Report
    '
    '   Imprime la vista del usuario.
    '---------------------------------------------------------------------------------------------
    Private Sub Print_report()
    'variables locales
    Dim cnnLocal As New ADODB.Connection
    Dim BD As Database
    Dim TDF As TableDef
    Dim blnExiste As Boolean
    Dim errLocal As Long
    Dim rpReporte As ctlReport
    '
'    On Error Resume Next
'    'si no existe la bd local la crea
'    If Dir(App.Path & "\Temp.mdb") = "" Then
'        DBEngine.CreateDatabase "Temp.mdb", dbLangSpanish
'    End If
'
'    Set BD = DBEngine.OpenDatabase(App.Path & "\Temp.mdb")
'    For Each TDF In BD.TableDefs
'        If UCase(TDF.Name) = UCase("Asistencia") Then blnExiste = True: Exit For
'    Next
'    cnnLocal.Open cnnOLEDB + App.Path + "\temp.mdb"
'    'si no existe la tabla la crea
'    If Not blnExiste Then
'        cnnLocal.Execute "CREATE TABLE Asistencia (Fecha DATETIME,Usuario TEXT(20),HEntrada TE" _
'        & "XT(11),HSalida TEXT(11),Duracion TEXT(10))"
'    Else
'        cnnLocal.Execute "DELETE * FROM Asistencia"
'    End If
'    'introduce los datos del grid
'    With grid
'        For I = 1 To .Rows - 1
'            cnnLocal.Execute "INSERT INTO Asistencia (Fecha,Usuario,Hentrada,Hsalida,Duracion) " _
'            & "VALUES ('" & .TextMatrix(I, 0) & "','" & .TextMatrix(I, 1) & "','" & _
'            .TextMatrix(I, 2) & "','" & .TextMatrix(I, 3) & "','" & .TextMatrix(I, 4) & "')"
'        Next
'        cnnLocal.Close
'    End With
'    Set cnnLocal = Nothing
'    BD.Close
'    Set BD = Nothing
    Dim strFiltro As String, strUser As String
    
    strFiltro = ""
    strUser = ""
    
    If dtcAsistencia <> "" And chkAsistencia(0) Then
        strUser = " {procAsistencia.Usuario}='" & dtcAsistencia & "'"
    End If
    'muestra los retardos
    If chkAsistencia(1) Then    'filtrar retardos
        
        strFiltro = "({procAsistencia.Entrada}>{Ambiente.HEntrada} and " & _
        "{procAsistencia.Entrada}>{Ambiente.HSalida}" & strUser & ") OR " & _
        "({procAsistencia.Salida}<{Ambiente.HSalida} and " & _
        "{procAsistencia.Salida}>{Ambiente.HEntrada}" & strUser & ")"
        
    End If
    'muestra las inasistencias
    If chkAsistencia(2) Then
        strFiltro = strFiltro & IIf(strFiltro = "", "", " OR ") & _
        "(({procAsistencia.Entrada} = CDateTime (1899, 12, 30, 00, 00, 00)" _
        & strUser & _
        ") OR ({procAsistencia.Entrada} = CDateTime (1899, 12, 30, 00, 00, 00)" _
        & strUser & "))"
    End If
    If strUser <> "" And strFiltro = "" Then strFiltro = Replace(strUser, " AND ", "")
    Set rpReporte = New ctlReport
    With rpReporte
        .Reporte = gcReport & "\Emp_Asistencia_new.rpt"
        .OrigenDatos(0) = gcPath & "\sac.mdb"
        .OrigenDatos(1) = gcPath & "\sac.mdb"
        .Formulas(0) = "Desde='" & dtpAsistencia(0) & "'"
        .Formulas(1) = "Hasta='" & dtpAsistencia(1) & "'"
        .Formulas(2) = "Subtitulo='Fase de Prueba'"
        .Parametros(0) = CDate(Me.dtpAsistencia(0))
        .Parametros(1) = CDate(Me.dtpAsistencia(1))
        .FormuladeSeleccion = strFiltro
        'destino de imrpesión
        If optAsistencia(0).Value = True Then
            .Salida = crPantalla
        Else
            .Salida = crImpresora
        End If
        .TituloVentana = "Control de Asistencia"
        .Imprimir
        Call rtnBitacora("Impresión Control de Asistencia")
        
    End With
    '
    If Err.Number <> 0 Then
        MsgBox "Ha ocurrido errores durante el proceso..." & vbCrLf & "Consulte al administrado" _
        & "r del sistema.", vbInformation, App.ProductName
        Call rtnBitacora(Err.Number & " - " & Err.Description)
    End If
    '
    End Sub

    Private Sub generar_procedure()
    Dim BD As Database
    Dim Qdf As QueryDef
    Dim blnExiste As Boolean
    Dim strSQL As String
    
    Set BD = DBEngine.OpenDatabase(gcPath & "\sac.mdb")
    For Each Qdf In BD.QueryDefs
        If UCase(Qdf.Name) = UCase("procAsistencia") Then blnExiste = True: Exit For
    Next
    If Not blnExiste Then
    
        strSQL = "PARAMETERS [Desde] DateTime, [Hasta] DateTime;" & _
        "SELECT Emp_Asistencia.Fecha, Emp_Asistencia.Usuario, " & _
        "Emp_Asistencia.Entrada, Emp_Asistencia.Salida, " & _
        "Format(CDate([Emp_Asistencia]![Salida]-[Emp_Asistencia]!" & _
        "[Entrada]),'hh"" h. ""nn"" m""') AS Duracion " & _
        "from Emp_Asistencia " & _
        "WHERE (((Emp_Asistencia.Fecha)>=[Desde] And (Emp_Asistencia.Fecha)<=[Hasta]));"
        BD.CreateQueryDef "procAsistencia", strSQL
'        strsql = "Create view procAsistencia as " & strsql
        'cnnConexion.Execute strsql
    End If
    End Sub
