VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Begin VB.Form FrmTgasto 
   Caption         =   "Tabla de Gastos"
   ClientHeight    =   510
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4935
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   HasDC           =   0   'False
   Icon            =   "FrmTgasto.frx":0000
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   510
   ScaleWidth      =   4935
   Tag             =   "1"
   WindowState     =   2  'Maximized
   Begin ComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   480
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   4935
      _ExtentX        =   8705
      _ExtentY        =   847
      ButtonWidth     =   714
      ButtonHeight    =   688
      Appearance      =   1
      ImageList       =   "ImageList1"
      _Version        =   327682
      BeginProperty Buttons {0713E452-850A-101B-AFC0-4210102A8DA7} 
         NumButtons      =   12
         BeginProperty Button1 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Top"
            Object.ToolTipText     =   "Primer Registro"
            Object.Tag             =   ""
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Back"
            Object.ToolTipText     =   "Registro Anterior"
            Object.Tag             =   ""
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Next"
            Object.ToolTipText     =   "Siguiente Registro"
            Object.Tag             =   ""
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "End"
            Object.ToolTipText     =   "Último Registro"
            Object.Tag             =   ""
            ImageIndex      =   4
         EndProperty
         BeginProperty Button5 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "New"
            Object.ToolTipText     =   "Nuevo Registro"
            Object.Tag             =   ""
            ImageIndex      =   5
         EndProperty
         BeginProperty Button6 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Enabled         =   0   'False
            Key             =   "Save"
            Object.ToolTipText     =   "Guardar Registro"
            Object.Tag             =   ""
            ImageIndex      =   6
         EndProperty
         BeginProperty Button7 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Enabled         =   0   'False
            Key             =   "Find"
            Object.ToolTipText     =   "Buscar Registro"
            Object.Tag             =   ""
            ImageIndex      =   7
         EndProperty
         BeginProperty Button8 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Enabled         =   0   'False
            Key             =   "Undo"
            Object.ToolTipText     =   "Cancelar Registro"
            Object.Tag             =   ""
            ImageIndex      =   8
         EndProperty
         BeginProperty Button9 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Delete"
            Object.ToolTipText     =   "Eliminar Registro"
            Object.Tag             =   ""
            ImageIndex      =   9
         EndProperty
         BeginProperty Button10 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Enabled         =   0   'False
            Key             =   "Edit1"
            Object.ToolTipText     =   "Editar Registro"
            Object.Tag             =   ""
            ImageIndex      =   10
         EndProperty
         BeginProperty Button11 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Print"
            Object.ToolTipText     =   "Imprimir"
            Object.Tag             =   ""
            ImageIndex      =   11
         EndProperty
         BeginProperty Button12 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Close"
            Object.ToolTipText     =   "Salir"
            Object.Tag             =   ""
            ImageIndex      =   12
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin VB.Frame fragasto 
      Height          =   7080
      Index           =   0
      Left            =   45
      TabIndex        =   0
      Top             =   525
      Width           =   11550
      Begin VB.Frame fragasto 
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   1455
         Index           =   3
         Left            =   270
         TabIndex        =   7
         Top             =   5505
         Width           =   9435
         Begin VB.TextBox txtGasto 
            Height          =   315
            Index           =   3
            Left            =   2100
            TabIndex        =   15
            Top             =   120
            Width           =   1000
         End
         Begin VB.OptionButton optBusca 
            Caption         =   "&Codigo"
            Height          =   300
            Index           =   0
            Left            =   1005
            TabIndex        =   14
            Top             =   75
            Value           =   -1  'True
            Width           =   930
         End
         Begin VB.OptionButton optBusca 
            Caption         =   "&Titulo"
            Height          =   300
            Index           =   1
            Left            =   1005
            TabIndex        =   13
            Top             =   495
            Width           =   1035
         End
         Begin VB.Frame fragasto 
            Caption         =   "Ordenar por:"
            Height          =   780
            Index           =   2
            Left            =   2100
            TabIndex        =   8
            Top             =   615
            Width           =   7260
            Begin VB.OptionButton optBusca 
               Caption         =   "Gastos Fijos"
               Height          =   195
               Index           =   2
               Left            =   1995
               TabIndex        =   12
               Tag             =   "Fijo DESC, CodGasto"
               Top             =   345
               Width           =   1635
            End
            Begin VB.OptionButton optBusca 
               Caption         =   "Gastos Comunes"
               Height          =   195
               Index           =   3
               Left            =   3645
               TabIndex        =   11
               Tag             =   "comun DESC, CodGAsto"
               Top             =   345
               Width           =   1635
            End
            Begin VB.OptionButton optBusca 
               Caption         =   "Código del Gasto"
               Height          =   195
               Index           =   4
               Left            =   300
               TabIndex        =   10
               Tag             =   "CodGasto"
               Top             =   345
               Value           =   -1  'True
               Width           =   1635
            End
            Begin VB.OptionButton optBusca 
               Caption         =   "Cuentas de Fondo"
               Height          =   195
               Index           =   5
               Left            =   5445
               TabIndex        =   9
               Tag             =   "Fondo DESC, CodGasto"
               Top             =   345
               Width           =   1635
            End
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Buscar por:"
            Height          =   195
            Left            =   90
            TabIndex        =   16
            Top             =   135
            Width           =   810
         End
      End
      Begin VB.Frame fragasto 
         Height          =   1275
         Index           =   1
         Left            =   9915
         TabIndex        =   3
         Top             =   5625
         Width           =   1500
         Begin VB.TextBox txtGasto 
            DataField       =   "Cuotas"
            Height          =   315
            Index           =   0
            Left            =   120
            TabIndex        =   6
            Top             =   480
            Width           =   1200
         End
         Begin VB.CheckBox chk 
            Caption         =   "Coletilla"
            DataField       =   "Incrementa"
            Height          =   300
            Left            =   150
            TabIndex        =   4
            Top             =   855
            Width           =   1200
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Nº Cuotas:"
            Height          =   195
            Left            =   120
            TabIndex        =   5
            Top             =   195
            Width           =   1185
         End
      End
      Begin MSDataGridLib.DataGrid DataGrid1 
         Height          =   4950
         Left            =   225
         TabIndex        =   2
         Top             =   360
         Width           =   10905
         _ExtentX        =   19235
         _ExtentY        =   8731
         _Version        =   393216
         AllowUpdate     =   -1  'True
         AllowArrows     =   -1  'True
         BorderStyle     =   0
         HeadLines       =   2
         RowHeight       =   15
         WrapCellPointer =   -1  'True
         RowDividerStyle =   4
         FormatLocked    =   -1  'True
         AllowAddNew     =   -1  'True
         AllowDelete     =   -1  'True
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "CATALOGO DE GASTOS"
         ColumnCount     =   7
         BeginProperty Column00 
            DataField       =   "CodGasto"
            Caption         =   "Código"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   8202
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   "Titulo"
            Caption         =   "Descripción"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   8202
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column02 
            DataField       =   "fijo"
            Caption         =   "Fijo"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   5
               Format          =   ""
               HaveTrueFalseNull=   1
               TrueValue       =   "S"
               FalseValue      =   "N"
               NullValue       =   ""
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   8202
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column03 
            DataField       =   "Comun"
            Caption         =   "Común"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   5
               Format          =   ""
               HaveTrueFalseNull=   1
               TrueValue       =   "S"
               FalseValue      =   "N"
               NullValue       =   ""
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   8202
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column04 
            DataField       =   "alicuota"
            Caption         =   "Alícutota"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   5
               Format          =   ""
               HaveTrueFalseNull=   1
               TrueValue       =   "S"
               FalseValue      =   "N"
               NullValue       =   ""
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   8202
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column05 
            DataField       =   "Fondo"
            Caption         =   "Fondo"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   5
               Format          =   ""
               HaveTrueFalseNull=   1
               TrueValue       =   "S"
               FalseValue      =   "N"
               NullValue       =   ""
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   8202
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column06 
            DataField       =   "MontoFijo"
            Caption         =   "Monto Fijo"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   "#,##0.00 ;(#,##0.00) "
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   8202
               SubFormatType   =   0
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            SizeMode        =   1
            BeginProperty Column00 
               Alignment       =   2
               ColumnWidth     =   810,142
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   5460,095
            EndProperty
            BeginProperty Column02 
               Alignment       =   2
               ColumnWidth     =   420,095
            EndProperty
            BeginProperty Column03 
               Alignment       =   2
               ColumnWidth     =   645,165
            EndProperty
            BeginProperty Column04 
               Alignment       =   2
               ColumnWidth     =   854,929
            EndProperty
            BeginProperty Column05 
               Alignment       =   2
               ColumnWidth     =   599,811
            EndProperty
            BeginProperty Column06 
               Alignment       =   1
               ColumnWidth     =   1379,906
            EndProperty
         EndProperty
      End
   End
   Begin ComctlLib.ImageList ImageList1 
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   20
      ImageHeight     =   20
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   12
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmTgasto.frx":000C
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmTgasto.frx":018E
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmTgasto.frx":0310
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmTgasto.frx":0492
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmTgasto.frx":0614
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmTgasto.frx":0796
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmTgasto.frx":0918
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmTgasto.frx":0A9A
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmTgasto.frx":0C1C
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmTgasto.frx":0D9E
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmTgasto.frx":0F20
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmTgasto.frx":10A2
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "FrmTgasto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
    Dim WithEvents rstTgasto As ADODB.Recordset
Attribute rstTgasto.VB_VarHelpID = -1
    Dim cnnTGasto As New ADODB.Connection
    Dim strCodIVA, strCodGA As String
    '
    Private Sub DataGrid1_AfterColUpdate(ByVal ColIndex As Integer)
    '
    Select Case ColIndex
        Case 2  'Edita la Columna de Gastos Fijos
    '   ---------------------
            If DataGrid1.Columns(2).Value = -1 Then
                DataGrid1.Columns(6) = 0
                Call rtnBitacora("Actualizar a Gasto Fijo Código: " & rstTgasto("CodGasto"))
                If gcNivel > nuAdministrador Then
                    Call Mensaje("Usuario: " & gcUsuario & ". Actualizó Gasto: " & _
                    rstTgasto("CodGasto") & " como fijo. Inmueble: " & gcCodInm)
                End If
            Else
                DataGrid1.Columns(6) = ""
                Call rtnBitacora("Actualizar a Gasto No Fijo Código: " & DataGrid1.Columns(0))
                If gcNivel > nuAdministrador Then
                    Call Mensaje("Usuario: " & gcUsuario & ". Actualizó Gasto: " & _
                    rstTgasto("CodGasto") & " como No Fijo. Inmueble: " & gcCodInm)
                End If
                '
'                If DataGrid1.Columns(0) = strCodGA Then
'                    'cnnTGasto.Execute "UPDATE Tgastos Set Fijo=False,MontoFijo=0 WHERE CodG" _
'                    & "asto='" & strCodIVA & "'"
'                End If
                '
            End If
            'cnnTGasto.Execute "UPDATE Tgastos SET Freg=Date(),Usuario='" & gcUsuario & _
            "' WHERE CodGasto='" & DataGrid1.Columns(0) & "';"

        Case 6  'Monto Gastos Fijos
    '   ---------------------
'            With DataGrid1
'                'Calcula el IVA a los honorarios administrativos si estan configurados como fijos
'                If Not .Columns(6) = "" And .Columns(2) = "S" And .Columns(0) = strCodGA _
'                    And gnIva > 0 Then
'                    cnnTGasto.Execute "UPDATE Tgastos Set Fijo=True,Comun=True,Alicuota=True,Mo" _
'                    & "ntoFijo=Clng(" & CLng(.Columns(6)) & " * " & gnIva & "/100) WHERE CodGas" _
'                    & "to='" & strCodIVA & "';"
'                End If
                '
'            End With
            '
    End Select
    '
    End Sub


    Private Sub DataGrid1_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))    'Convierte todo en mayúsculas
    End Sub

    Private Sub DataGrid1_KeyUp(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    If KeyCode = 113 And DataGrid1.Col = 1 Then
        DataGrid1.SelStart = Len(DataGrid1.Text)
    End If
    End Sub

    '---------------------------------------------------------------------------------------------
    Private Sub Form_Load() '
    '---------------------------------------------------------------------------------------------
    '
    Set rstTgasto = New ADODB.Recordset
    cnnTGasto.Open cnnOLEDB & mcDatos
    rstTgasto.CursorLocation = adUseClient
    'rstTgasto.ConnectionString = cnnTGasto.ConnectionString
    'rstTgasto.Refresh
    rstTgasto.Open "TGastos", cnnTGasto, adOpenKeyset, adLockOptimistic, adCmdTable
    rstTgasto.Sort = "CodGasto"
    '
    Set DataGrid1.DataSource = rstTgasto
    Set txtGasto(0).DataSource = rstTgasto
    Set chk.DataSource = rstTgasto
    Set DataGrid1.HeadFont = LetraTitulo(LoadResString(527), 7.5, True)
    Set DataGrid1.Font = LetraTitulo(LoadResString(528), 8)
    strCodIVA = FrmAdmin.objRst!CodIVA
    strCodGA = FrmAdmin.objRst!CodGastoAdmin
   ' DataGrid1.Columns(1).DataFormat = CheckBox
    '
    End Sub

    '---------------------------------------------------------------------------------------------
    Private Sub Form_Resize()   '
    '---------------------------------------------------------------------------------------------
    '
    If WindowState <> vbMinimized Then
        fragasto(0).Height = ScaleHeight - (Toolbar1.Height * 2)
        fragasto(0).Top = (ScaleHeight - fragasto(0).Height) / 2 '+ (Toolbar1.Height / 2)
        fragasto(0).Left = ScaleWidth / 2 - fragasto(0).Width / 2
        DataGrid1.Height = fragasto(0).Height - fragasto(3).Height - DataGrid1.Top - 360
        fragasto(1).Top = fragasto(0).Height - fragasto(1).Height - 180
        fragasto(3).Top = fragasto(1).Top - 120
    End If
    '
    End Sub

    '---------------------------------------------------------------------------------------------
    Private Sub Form_Unload(Cancel As Integer)  '
    '---------------------------------------------------------------------------------------------
    On Error Resume Next
    rstTgasto.Close
    Set rstTgasto = Nothing
    cnnTGasto.Close
    Set cnnTGasto = Nothing
    End Sub

    '---------------------------------------------------------------------------------------------
    Private Sub OptBusca_Click(Index As Integer)    '
    '---------------------------------------------------------------------------------------------
    Select Case Index
    '
        Case 0  'Busqueda por codigo
    '   ---------------------
        txtGasto(3).Width = 1000
        txtGasto(3) = ""
        '
        Case 1  'Busqueda por titulo
    '   ---------------------
        txtGasto(3).Width = 7305
        txtGasto(3) = ""
        '
        Case 2, 3, 5, 4 'Filtrar
    '   ---------------------
        MousePointer = vbHourglass
        'rstTGasto.Close
        'rstTGasto.Source = "SELECT * FROM Tgastos ORDER BY " & optBusca(Index).Tag
        'rstTGasto.Open
        'Set DataGrid1.DataSource = rstTGasto
        'rstTGasto.Filter = optBusca(Index).Tag
        'rstTGasto.Requery
        rstTgasto.Sort = OptBusca(Index).Tag
        MousePointer = vbDefault
        
        'Case 4
        'MousePointer = vbHourglass
        'rstTGasto.Filter = ""
        'rstTGasto.Requery
        'rstTGasto.Sort = "CodGasto"
        'MousePointer = vbDefault
        '
    End Select
    '
    End Sub

    Private Sub rstTgasto_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, _
    ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
    '
    'si no es fin o inicio de archivo
    If Not rstTgasto.EOF And Not rstTgasto.BOF Then
        fragasto(1).Visible = _
        IIf(IsNull(rstTgasto.Fields("Fondo")), False, rstTgasto.Fields("Fondo"))
    End If
    '
    End Sub

    '---------------------------------------------------------------------------------------------
    Private Sub Toolbar1_ButtonClick(ByVal Button As ComctlLib.Button)  '
    '---------------------------------------------------------------------------------------------
    '
    With rstTgasto
    '
        Select Case Button.Index
         '
            Case 1  'Primer Registro
        '   -------------------------------
            Case 2  'Registro Previo
        '   -------------------------------
            Case 3  'Siguiente Registro
        '   -------------------------------
            Case 4  'Último Registro
        '   -------------------------------
            Case 5  'Agregar registro
        '   -------------------------------
                .AddNew
                .Fields("SaldoActual") = 0
                .Fields("Saldo1") = 0
                .Fields("Saldo2") = 0
                .Fields("Saldo3") = 0
                .Fields("Saldo4") = 0
                .Fields("Saldo5") = 0
                .Fields("Saldo6") = 0
                .Fields("Saldo7") = 0
                .Fields("Saldo8") = 0
                .Fields("Saldo9") = 0
                .Fields("Saldo10") = 0
                .Fields("Saldo11") = 0
                .Fields("Saldo12") = 0
                .Fields("Freg") = Date
                .Fields("Usuario") = gcUsuario
                .Fields("Incrementa") = False
                .Fields("Cuotas") = 0
                .Fields("Van") = 0
                '----------------
    
            Case 9  'Eliminar Registro
        '   -------------------------------
                If Respuesta(LoadResString(526)) Then
                    .Delete
                    .MoveNext
                    If .EOF Then .MoveLast
                    MsgBox " Registro Eliminado ... ", vbInformation, App.ProductName
                End If
    
            Case 10 ' Editar Registro
        '   -------------------------------
            Case 12 'Descargar Formulario
        '   -------------------------------
                If .EditMode = adEditInProgress Then .Update
                Unload Me
    
            Case 11 'Imprimir Registro
        '   -------------------------------
                mcTitulo = "Catálogo de Conceptos de Gastos"
                mcReport = "ListaGas.Rpt"
                mcOrdCod = "+{Tgastos.CodGasto}"
                mcOrdAlfa = "+{Tgastos.Titulo}"
                If OptBusca(2) Then 'gastos fijos
                    mcCrit = "{Tgastos.Fijo}=True"
                ElseIf OptBusca(3) Then 'gastos comunes
                    mcCrit = "{Tgastos.Comun}=True"
                ElseIf OptBusca(4) Then '
                    mcCrit = ""
                ElseIf OptBusca(5) Then 'gastos de fondo
                    mcCrit = "{Tgastos.Fondo}=True"
                End If
                FrmReport.Show
        End Select
        '
    End With
    '
    End Sub


    '---------------------------------------------------------------------------------------------
    Private Sub txtGasto_Change(Index As Integer)   '
    '---------------------------------------------------------------------------------------------
    'variables locales
    Dim Criterio As String
    '
    If Index = 3 And txtGasto(3) <> "" Then
        If OptBusca(0).Value = True Then
             Criterio = "CodGasto Like '" & txtGasto(3) & "*'"
        Else
            Criterio = "Titulo Like '" & txtGasto(3) & "*'"
        End If
        rstTgasto.MoveFirst
        rstTgasto.Find Criterio
    End If
    '
    End Sub

    '---------------------------------------------------------------------------------------------
    Private Sub txtGasto_KeyPress(Index As Integer, KeyAscii As Integer)    '
    '---------------------------------------------------------------------------------------------
    '
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then
    '
        Select Case Index
            Case 0
                'txtGasto(1).SetFocus
            Case 1
                If txtGasto(1) <> "" Then
                    txtGasto(1) = Format(txtGasto(1), "#,##0.00")
                End If
                txtGasto(2).SetFocus

        End Select
    '
    Else
        If Index = 0 Then Call Validacion(KeyAscii, "0123456789")
        
    End If
    '
    End Sub

'    '---------------------------------------------------------------------------------------------
'    Private Sub rtnBuscaGasto(campo, valor As String) '
'    '---------------------------------------------------------------------------------------------
'    '
'    With rstTgasto
'    '
'        .MoveFirst
'        .Find campo & "  Like '" & valor & "*'"
'    End With
'    '
'    End Sub

    '---------------------------------------------------------------------------------------------
    '   Funcion:    Filtro_Recordset
    '
    '   Devuelve un nuevo ADODB.Recordset filtrado según parametros del usuario
    '---------------------------------------------------------------------------------------------
    Private Function Filtro_Recordset(rstTemp As ADODB.Recordset, Optional strFilter As String) _
    As ADODB.Recordset
    '
    rstTemp.Close
    If strFilter = "" Then
        rstTemp.Open "Tgastos", rstTemp.ActiveConnection, adOpenKeyset, adLockOptimistic, adCmdTable
    Else
        rstTemp.Open "SELECT * FROM Tgastos WHERE " & strFilter, rstTemp.ActiveConnection, adOpenKeyset, adLockOptimistic, adCmdText
    End If
    Set Filtro_Recordset = rstTemp
    '
    End Function

    '---------------------------------------------------------------------------------------------
    '   Rutina:     Mensaje
    '
    '   Envia un mensaje a los supervisores del cambio que está efectuando
    '
    '---------------------------------------------------------------------------------------------
    Private Sub Mensaje(strMsg As String)
    'variables locales
    Dim m As Long
    'si dispone de RealpoPup lo utiliza como medio para solicitar la autorización
    'a un supervisor
    If Dir("C:\Archivos de programa\RealPopup\RealPopup.exe") <> "" Then
        m = Shell("C:\Archivos de programa\RealPopup\RealPopup -send Condominio " & _
        Chr(34) & strMsg & Chr(34) & " -NOACTIVATE")
    End If
    '
    End Sub
