VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmPlanPago 
   AutoRedraw      =   -1  'True
   Caption         =   "Plan de Pagos"
   ClientHeight    =   30
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   2115
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   30
   ScaleWidth      =   2115
   WindowState     =   2  'Maximized
   Begin VB.Frame fra 
      Height          =   1380
      Index           =   0
      Left            =   285
      TabIndex        =   0
      Top             =   105
      Width           =   11295
      Begin VB.Frame fra 
         Caption         =   "Fecha Actualización: "
         Height          =   630
         Index           =   2
         Left            =   3000
         TabIndex        =   10
         Top             =   675
         Width           =   2880
         Begin MSComCtl2.DTPicker dtp 
            Height          =   315
            Index           =   1
            Left            =   1320
            TabIndex        =   12
            Top             =   210
            Width           =   1500
            _ExtentX        =   2646
            _ExtentY        =   556
            _Version        =   393216
            CalendarTitleBackColor=   -2147483646
            CalendarTitleForeColor=   -2147483639
            Format          =   60030977
            CurrentDate     =   37950
         End
      End
      Begin VB.Frame fra 
         Caption         =   "Fecha Consulta: "
         Height          =   615
         Index           =   1
         Left            =   75
         TabIndex        =   9
         Top             =   675
         Width           =   2880
         Begin MSComCtl2.DTPicker dtp 
            Height          =   315
            Index           =   0
            Left            =   1320
            TabIndex        =   11
            Top             =   210
            Width           =   1500
            _ExtentX        =   2646
            _ExtentY        =   556
            _Version        =   393216
            CalendarTitleBackColor=   -2147483646
            CalendarTitleForeColor=   -2147483639
            Format          =   60030977
            CurrentDate     =   37950
         End
      End
      Begin VB.Frame Frame1 
         Height          =   630
         Left            =   5940
         TabIndex        =   1
         Top             =   675
         Width           =   5265
         Begin VB.CommandButton cmd 
            Caption         =   "Guardar"
            Enabled         =   0   'False
            Height          =   375
            Index           =   5
            Left            =   135
            Picture         =   "frmPlanPago.frx":0000
            TabIndex        =   21
            Top             =   180
            Width           =   840
         End
         Begin VB.CommandButton cmd 
            Caption         =   "Editar"
            Height          =   375
            Index           =   4
            Left            =   990
            Picture         =   "frmPlanPago.frx":030A
            TabIndex        =   20
            Top             =   180
            Width           =   840
         End
         Begin VB.CommandButton cmd 
            Caption         =   "Cancelar"
            Enabled         =   0   'False
            Height          =   375
            Index           =   3
            Left            =   1830
            Picture         =   "frmPlanPago.frx":0614
            TabIndex        =   19
            Top             =   180
            Width           =   840
         End
         Begin VB.CommandButton cmd 
            Caption         =   "Imprimir"
            Height          =   375
            Index           =   1
            Left            =   3510
            Picture         =   "frmPlanPago.frx":091E
            TabIndex        =   18
            Top             =   180
            Width           =   840
         End
         Begin VB.CommandButton cmd 
            Caption         =   "Cerrar"
            Height          =   375
            Index           =   0
            Left            =   4350
            Picture         =   "frmPlanPago.frx":0C28
            TabIndex        =   17
            Top             =   180
            Width           =   840
         End
         Begin VB.CommandButton cmd 
            Caption         =   "Reiniciar"
            Height          =   375
            Index           =   2
            Left            =   2670
            Picture         =   "frmPlanPago.frx":1F22
            TabIndex        =   16
            Top             =   180
            Width           =   840
         End
      End
      Begin MSDataListLib.DataCombo dtc 
         DataField       =   "CodInm"
         Height          =   315
         Index           =   0
         Left            =   930
         TabIndex        =   2
         Top             =   300
         Width           =   1050
         _ExtentX        =   1852
         _ExtentY        =   556
         _Version        =   393216
         ListField       =   "CodInm"
         BoundColumn     =   "Nombre"
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo dtc 
         DataField       =   "Nombre"
         Height          =   315
         Index           =   1
         Left            =   2040
         TabIndex        =   3
         Top             =   300
         Width           =   4455
         _ExtentX        =   7858
         _ExtentY        =   556
         _Version        =   393216
         ListField       =   "Nombre"
         BoundColumn     =   "CodInm"
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo dtc 
         Height          =   315
         Index           =   2
         Left            =   7335
         TabIndex        =   4
         Top             =   285
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         ListField       =   "NombreBanco"
         BoundColumn     =   "NumCuenta"
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo dtc 
         Height          =   315
         Index           =   3
         Left            =   9165
         TabIndex        =   5
         Top             =   285
         Width           =   2010
         _ExtentX        =   3545
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         ListField       =   "NumCuenta"
         BoundColumn     =   "NombreBanco"
         Text            =   ""
      End
      Begin VB.Label lbl 
         Caption         =   "Inmueble:"
         Height          =   240
         Index           =   0
         Left            =   180
         TabIndex        =   7
         Top             =   345
         Width           =   1215
      End
      Begin VB.Label lbl 
         Caption         =   "Cuentas:"
         Height          =   240
         Index           =   1
         Left            =   6585
         TabIndex        =   6
         Top             =   315
         Width           =   1215
      End
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid Grid 
      Height          =   5520
      Left            =   285
      TabIndex        =   8
      Tag             =   "900|1200|4000|1500|1200|500|1790|0"
      Top             =   1620
      Width           =   11325
      _ExtentX        =   19976
      _ExtentY        =   9737
      _Version        =   393216
      Cols            =   6
      FixedCols       =   0
      BackColorFixed  =   -2147483646
      ForeColorFixed  =   -2147483639
      BackColorBkg    =   -2147483636
      GridColor       =   -2147483636
      GridColorFixed  =   -2147483636
      WordWrap        =   -1  'True
      AllowBigSelection=   0   'False
      HighLight       =   2
      GridLines       =   2
      GridLinesFixed  =   1
      GridLinesUnpopulated=   2
      SelectionMode   =   1
      AllowUserResizing=   1
      FormatString    =   "^Cheque|^Fecha Cheque|Beneficiario|>Monto|^Fecha Pago|^Act."
      _NumberOfBands  =   1
      _Band(0).Cols   =   6
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   0
   End
   Begin VB.Line Line1 
      Index           =   1
      X1              =   2475
      X2              =   3840
      Y1              =   7680
      Y2              =   7680
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      Caption         =   "0,00"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   4
      Left            =   2385
      TabIndex        =   15
      Top             =   7470
      Width           =   1440
   End
   Begin VB.Line Line1 
      Index           =   0
      X1              =   1470
      X2              =   2385
      Y1              =   7680
      Y2              =   7680
   End
   Begin VB.Label lbl 
      Alignment       =   2  'Center
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   3
      Left            =   1500
      TabIndex        =   14
      Top             =   7470
      Width           =   810
   End
   Begin VB.Label lbl 
      Caption         =   "Total Cheques:"
      Height          =   240
      Index           =   2
      Left            =   285
      TabIndex        =   13
      Top             =   7470
      Width           =   1215
   End
   Begin VB.Image img 
      Enabled         =   0   'False
      Height          =   150
      Index           =   0
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   165
   End
   Begin VB.Image img 
      Enabled         =   0   'False
      Height          =   150
      Index           =   1
      Left            =   195
      Top             =   0
      Visible         =   0   'False
      Width           =   165
   End
End
Attribute VB_Name = "frmPlanPago"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
    Dim rstPP(3) As New ADODB.Recordset
    Dim strOrigenDatos As String
    Dim blnPote As Boolean
    Dim blnEdit As Boolean
    Dim Filtro_Cadena As String
    
    '
    
    
    Private Sub Cmd_Click(Index As Integer)
    'variables locales
    Dim i As Integer
    '
    Select Case Index
        'cerrar formulario
        Case 0: Unload Me: Set frmPlanPago = Nothing
        
        'imprimir
        Case 1: MsgBox "Opción no disponible", vbInformation, App.ProductName
        
        Case 2  'reiniciar
            For i = 0 To 3: dtc(i) = ""
            Next
            Call Filtrar
            
        'cancelar
        Case 3
            cmd(0).Enabled = True
            cmd(1).Enabled = True
            cmd(2).Enabled = True
            cmd(4).Enabled = True
            cmd(5).Enabled = False
            cmd(3).Enabled = False
            cnnConexion.RollbackTrans
            blnEdit = False
            Call rtnBitacora("Cancelar Cambios Plan de Pagos")
            Call Filtrar
            
        'Editar información
        Case 4
            cmd(0).Enabled = False
            cmd(1).Enabled = False
            cmd(2).Enabled = False
            cmd(4).Enabled = False
            cmd(3).Enabled = True
            cmd(5).Enabled = True
            blnEdit = True
            cnnConexion.BeginTrans
            Call rtnBitacora("Editar Plan de Pagos")
            
            
        'guardar
        Case 5
            cmd(0).Enabled = True
            cmd(1).Enabled = True
            cmd(2).Enabled = True
            cmd(4).Enabled = True
            cmd(5).Enabled = False
            cmd(3).Enabled = False
            cnnConexion.CommitTrans
            blnEdit = False
            Call rtnBitacora("Guardar Cambios Plan de Pagos")
            Call Formatear_Grid
            
    End Select
    '
    End Sub


    Private Sub dtp_Change(Index As Integer)
    Select Case Index
        Case 0  'fecha de selección

            Grid.Visible = False
            Set Grid.DataSource = Nothing
            For i = 0 To 10000
            Next
            If rstPP(0).State = 1 Then rstPP(0).Close
            rstPP(0).Open "SELECT Ch.IDCheque, Format(Ch.FechaCheque,'dd/mm/yyyy'), Ch.Benefici" _
            & "ario, Sum(CD.Monto) as Total,Format(Ch.Fecha,'dd/mm/yyyy'), CD.CodInm, Ch.Cuenta" _
            & ",Ch.Fecha FROM Cheque as Ch INNER JOIN ChequeDetalle as CD ON Ch.IDCheque = CD.I" _
            & "DCheque WHERE Ch.Fecha <=#" & Format(dtp(0), "mm/dd/yyyy") & "# AND IDEstado=0 G" _
            & "ROUP BY Ch.IDCheque,Ch.FechaCheque, Ch.Beneficiario, Ch.Fecha,CD.CodInm,Ch.Cuent" _
            & "a  ORDER BY Ch.Fecha dESC", cnnConexion, adOpenKeyset, adLockOptimistic, adCmdText
            Set Grid.DataSource = rstPP(0)
            lbl(3) = rstPP(0).RecordCount
            Call centra_titulo(Grid, True)
            Call Formatear_Grid
            Grid.Visible = True

        Case 1  'fecha de actualización

    End Select
    End Sub


    Private Sub Form_Load()
    '
    'carga el formulario
    '
    dtp(0).Value = Date
    dtp(1).Value = Date
    img(0).Picture = LoadResPicture("Unchecked", vbResBitmap)
    img(1).Picture = LoadResPicture("Checked", vbResBitmap)
    '
    'abre los origenes de datos
    rstPP(0).Open "SELECT Ch.IDCheque, Format(Ch.FechaCheque,'dd/mm/yyyy') as FCheque, Ch.Benef" _
    & "iciario, Sum(CD.Monto) as Total,Format(Ch.Fecha,'dd/mm/yyyy') as FPago,'', Ch.Cuenta,Ch." _
    & "Fecha FROM Cheque as Ch INNER JOIN ChequeDetalle as CD ON Ch.IDCheque =CD.IDCheque WHERE" _
    & " Ch.Fecha <=Date() AND Ch.IDEstado=0 GROUP BY Ch.IDCheque,Ch.FechaCheque, Ch.Beneficiari" _
    & "o, Ch.Fecha,Ch.Cuenta ORDER BY Ch.FEcha DESC", cnnConexion, adOpenKeyset, _
    adLockOptimistic, adCmdText
    lbl(3) = rstPP(0).RecordCount
    '
    rstPP(1).Open "SELECT * FROM Inmueble ORDER BY CodInm", cnnConexion, adOpenKeyset, _
    adLockReadOnly, adCmdText
    '
    rstPP(2).Open "SELECT * FROM Inmueble ORDER BY Nombre", cnnConexion, adOpenStatic, _
    adLockReadOnly, adCmdText
    '
    Set dtc(0).RowSource = rstPP(1)
    Set dtc(1).RowSource = rstPP(2)
    'configura la presentación del grid
    Set Grid.DataSource = rstPP(0)
    Set Grid.FontFixed = LetraTitulo(LoadResString(527), 7.5, , True)
    Set Grid.Font = LetraTitulo(LoadResString(528), 8)
    Call centra_titulo(Grid, True)
    Call Formatear_Grid
    
    '
    End Sub

    
    Private Sub Form_Unload(Cancel As Integer)
    'cierra los objetos
    On Error Resume Next
    For i = 0 To 3
        rstPP(i).Close
        Set rstPP(i) = Nothing
    Next
    '
    End Sub

    '---------------------------------------------------------------------------------------------
    '   Rutina: Inf_Inmueble
    '
    '   Busca la información de las cuentas del inmueble
    '
    '---------------------------------------------------------------------------------------------
    Private Sub Inf_Inmueble()
    'variables locales
    Dim strSql As String
    '
    With rstPP(1)
        
        .Find "CodInm='" & dtc(0) & "'"
        If .EOF Then
            .MoveFirst
            .Find "CodInm='" & dtc(0) & "'"
            If .EOF Then
                MsgBox "Inmueble no registrado", vbInformation, App.EXEName
                Exit Sub
            End If
        End If
        dtc(2) = ""
        dtc(3) = ""
        'Grid.Rows = 2
        'Call rtnLimpiar_Grid(Grid)
        If !Caja = sysCodCaja Then
            strOrigenDatos = gcPath & "\" & suscodinm & "\inm.mdb"
            blnPote = True
        Else
            strOrigenDatos = gcPath & !Ubica & "inm.mdb"
            blnPote = False
        End If
    End With
    'cuentas del inmueble
    strSql = "SELECT Bancos.*, Cuentas.* FROM Bancos INNER JOIN Cuentas ON Bancos.IDBanco = Cue" _
    & "ntas.IDBanco;"
    If Dir(strOrigenDatos) <> "" Then
        If rstPP(3).State = 1 Then rstPP(3).Close
        rstPP(3).Open strSql, cnnOLEDB + strOrigenDatos, adOpenStatic, adLockReadOnly, _
        adCmdText
        '
        Set dtc(2).RowSource = rstPP(3)
        Set dtc(3).RowSource = rstPP(3)
        Filtro_Cadena = ""
        With rstPP(3)
            If Not .EOF And Not .BOF Then
                .MoveFirst
                Do
                    Filtro_Cadena = Filtro_Cadena & IIf(Filtro_Cadena = "", "", " or ") & "Cuenta ='" & .Fields("NumCuenta") & "'"
                    .MoveNext
                Loop Until .EOF
            End If
        End With
    Else
        Set dtc(2).RowSource = Nothing
        Set dtc(3).RowSource = Nothing
    End If
    '
    End Sub

    Private Sub dtc_Click(Index As Integer, Area As Integer)
    '
    If Area = 2 Then
    
        Select Case Index
        
            Case 0, 1 'codinm y nombre inm
                If Index = 0 Then dtc(1) = dtc(0).BoundText
                If Index = 1 Then dtc(1) = dtc(0).BoundText
                Call Inf_Inmueble
                
                
            Case 2, 3 'banco y cuenta
                If Index = 2 Then dtc(3) = dtc(2).BoundText
                If Index = 3 Then dtc(2) = dtc(3).BoundText
                Call rtnLimpiar_Grid(Grid)

        End Select
        Call Filtrar
    End If
    
    End Sub

    Private Sub Filtrar()
    'variables locales
    Dim strFiltro As String
    '
    Set Grid.DataSource = Nothing
    If dtc(0) <> "" And dtc(3) = "" Then
        strFiltro = Filtro_Cadena
    End If
    If dtc(3) <> "" Then
        strFiltro = strFiltro & IIf(strFiltro = "", " ", " AND ") & "Cuenta ='" & dtc(3) & "'"
    End If
    rstPP(0).Filter = strFiltro
    Set Grid.DataSource = rstPP(0)
    lbl(3) = rstPP(0).RecordCount
    Call centra_titulo(Grid, True)
    Call Formatear_Grid
    End Sub


    Private Sub Formatear_Grid()
    'variables locales
    Dim i As Integer
    '
    With Grid
        .Visible = False
        lbl(4) = 0
        For i = 1 To .Rows - 1
            .TextMatrix(i, 0) = Format(.TextMatrix(i, 0), "000000")
            lbl(4) = Format(CCur(lbl(4)) + CCur(.TextMatrix(i, 3)), "#,##0.00")
            .TextMatrix(i, 3) = Format(.TextMatrix(i, 3), "#,##0.00 ")
            .TextMatrix(i, 5) = ""
            .Col = 5: .Row = i
            Set .CellPicture = img(0)
            .CellPictureAlignment = flexAlignCenterCenter
        Next
        .Visible = True
    End With
    '
    End Sub

    Private Sub grid_Click()
    'variables locales
    '
    ' sale de la rutina si no se está editando
    ' o si no tiene el permiso necesario
    If Not blnEdit Or gcNivel > nuSUPERVISOR Then Exit Sub
    
    '
    With Grid
        .Col = 5
        
        Set .CellPicture = IIf(.CellPicture = img(0), img(1), img(0))
        If .CellPicture = img(1) Then
            .TextMatrix(.Row, 7) = .TextMatrix(.Row, 4)
            .TextMatrix(.Row, 4) = dtp(1)
        Else
            .TextMatrix(.Row, 4) = .TextMatrix(.Row, 7)
        End If
        .Col = 0
        cnnConexion.Execute "UPDATE Cheque SET Fecha='" & .TextMatrix(.Row, 4) & _
        "' WHERE IDCheque=" & .Text
        Call rtnBitacora("Actualizar Fecha de Pago Cheque " & .Text)
        
    End With
    '
    End Sub

    
