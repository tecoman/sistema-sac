VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmIntDes 
   Caption         =   "Intereses Descontados"
   ClientHeight    =   1155
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   2475
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   1155
   ScaleWidth      =   2475
   WindowState     =   2  'Maximized
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid FLEX 
      Height          =   4065
      Left            =   345
      TabIndex        =   9
      Tag             =   "1200|800|7500|1500|0|0"
      Top             =   3945
      Width           =   11340
      _ExtentX        =   20003
      _ExtentY        =   7170
      _Version        =   393216
      Cols            =   6
      FixedCols       =   0
      BackColorFixed  =   -2147483646
      ForeColorFixed  =   -2147483639
      BackColorSel    =   65280
      ForeColorSel    =   -2147483646
      BackColorBkg    =   -2147483636
      GridColor       =   -2147483633
      SelectionMode   =   1
      AllowUserResizing=   1
      MousePointer    =   99
      GridLineWidthFixed=   1
      FormatString    =   "Fecha |Apto. |Descripción |Monto |1|2"
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
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FontWidthFixed  =   4
      MouseIcon       =   "frmIntDes.frx":0000
      _NumberOfBands  =   1
      _Band(0).Cols   =   6
   End
   Begin VB.Frame fraIntDes 
      Height          =   3405
      Index           =   0
      Left            =   345
      TabIndex        =   0
      Top             =   300
      Width           =   11370
      Begin VB.Frame fraIntDes 
         Caption         =   "Imprimir a:"
         Height          =   1140
         Index           =   3
         Left            =   5880
         TabIndex        =   15
         Top             =   2025
         Width           =   2190
         Begin VB.OptionButton Opt 
            Caption         =   "Ventana"
            Height          =   315
            Index           =   6
            Left            =   360
            TabIndex        =   17
            Tag             =   "Deducciones.FecReg"
            Top             =   315
            Value           =   -1  'True
            Width           =   1215
         End
         Begin VB.OptionButton Opt 
            Caption         =   "Impresora"
            Height          =   315
            Index           =   7
            Left            =   360
            TabIndex        =   16
            Tag             =   "MovimientoCaja.AptoMovimientoCaja"
            Top             =   705
            Width           =   1215
         End
      End
      Begin VB.Frame fraIntDes 
         Caption         =   "Ordenar por:"
         Height          =   2085
         Index           =   2
         Left            =   3495
         TabIndex        =   10
         Top             =   1095
         Width           =   2190
         Begin VB.OptionButton Opt 
            Caption         =   "Monto"
            Height          =   315
            Index           =   5
            Left            =   360
            TabIndex        =   14
            Tag             =   "Deducciones.Monto"
            Top             =   1620
            Width           =   1215
         End
         Begin VB.OptionButton Opt 
            Caption         =   "Descripción"
            Height          =   315
            Index           =   4
            Left            =   360
            TabIndex        =   13
            Tag             =   "Deducciones.Titulo"
            Top             =   1185
            Width           =   1215
         End
         Begin VB.OptionButton Opt 
            Caption         =   "Apartamento"
            Height          =   315
            Index           =   3
            Left            =   360
            TabIndex        =   12
            Tag             =   "MovimientoCaja.AptoMovimientoCaja"
            Top             =   750
            Width           =   1215
         End
         Begin VB.OptionButton Opt 
            Caption         =   "Fecha"
            Height          =   315
            Index           =   2
            Left            =   360
            TabIndex        =   11
            Tag             =   "Deducciones.FecReg"
            Top             =   315
            Value           =   -1  'True
            Width           =   1215
         End
      End
      Begin VB.CommandButton cmd 
         Caption         =   "Imprimir"
         Height          =   885
         Index           =   1
         Left            =   8640
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   2130
         Width           =   1185
      End
      Begin VB.CommandButton cmd 
         Caption         =   "Salir"
         Height          =   885
         Index           =   0
         Left            =   9825
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   2130
         Width           =   1185
      End
      Begin VB.Frame fraIntDes 
         Caption         =   "Selección de Registros:"
         Height          =   2070
         Index           =   1
         Left            =   240
         TabIndex        =   4
         Top             =   1110
         Width           =   2865
         Begin VB.OptionButton Opt 
            Caption         =   "&Entre el "
            Height          =   495
            Index           =   1
            Left            =   180
            TabIndex        =   6
            Top             =   885
            Width           =   1005
         End
         Begin VB.OptionButton Opt 
            Caption         =   "&Todos"
            Height          =   495
            Index           =   0
            Left            =   180
            TabIndex        =   5
            Top             =   300
            Value           =   -1  'True
            Width           =   1215
         End
         Begin MSComCtl2.DTPicker dtp 
            Height          =   315
            Index           =   1
            Left            =   1230
            TabIndex        =   20
            Top             =   1500
            Width           =   1350
            _ExtentX        =   2381
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
            Format          =   21430273
            CurrentDate     =   37804
         End
         Begin MSComCtl2.DTPicker dtp 
            Height          =   315
            Index           =   0
            Left            =   1230
            TabIndex        =   21
            Top             =   990
            Width           =   1350
            _ExtentX        =   2381
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
            Format          =   21430273
            CurrentDate     =   37804
         End
      End
      Begin MSDataListLib.DataCombo dtc 
         Bindings        =   "frmIntDes.frx":0162
         Height          =   315
         Index           =   0
         Left            =   1230
         TabIndex        =   2
         Top             =   420
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   556
         _Version        =   393216
         ListField       =   "CodInm"
         BoundColumn     =   "Nombre"
         Text            =   ""
         Object.DataMember      =   "CmdListCod"
      End
      Begin MSDataListLib.DataCombo dtc 
         Bindings        =   "frmIntDes.frx":017D
         Height          =   315
         Index           =   1
         Left            =   2460
         TabIndex        =   3
         Top             =   420
         Width           =   8610
         _ExtentX        =   15187
         _ExtentY        =   556
         _Version        =   393216
         ListField       =   "Nombre"
         BoundColumn     =   "CodInm"
         Text            =   ""
         Object.DataMember      =   "CmdListNombre"
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
         Height          =   315
         Index           =   2
         Left            =   6675
         TabIndex        =   19
         Top             =   1125
         Width           =   1605
      End
      Begin VB.Label lbl 
         Caption         =   "TOTAL:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   1
         Left            =   5895
         TabIndex        =   18
         Top             =   1170
         Width           =   705
      End
      Begin VB.Label lbl 
         Caption         =   "Inmueble"
         Height          =   315
         Index           =   0
         Left            =   315
         TabIndex        =   1
         Top             =   480
         Width           =   855
      End
   End
End
Attribute VB_Name = "frmIntDes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmd_Click(Index As Integer)
MousePointer = vbHourglass
If Index = 0 Then
    Me.Hide
ElseIf Index = 1 Then 'Imprimir el reporte de intereses descontados
    Call Report
    'MsgBox "No disponible"
End If
MousePointer = vbDefault
End Sub

Private Sub dtc_Click(Index As Integer, Area As Integer)
If Area = 2 Then
    If Index = 0 Then dtc(1) = dtc(0).BoundText
    If Index = 1 Then dtc(0) = dtc(1).BoundText
    If dtc(1).MatchedWithList And dtc(0).MatchedWithList Then Call View
    'Else
        'Call rtnLimpiar_Grid(Flex)
    'End If
End If
'
End Sub

Private Sub dtc_KeyPress(Index As Integer, KeyAscii As Integer)
'variables locales
Dim strCrit$
'
KeyAscii = Asc(UCase(Chr(KeyAscii)))
If Index = 0 Then Call Validacion(KeyAscii, "1234567890")
If KeyAscii = 13 Then
    If Index = 0 Then strCrit = "='" & dtc(Index) & "'"
    If Index = 1 Then strCrit = " LIKE '%" & dtc(Index) & "%'"
    With FrmAdmin.objRst
        .MoveFirst
        .Find .Fields(Index).Name & strCrit
        If Not .EOF And Not .BOF Then
            dtc(0) = .Fields(0)
            dtc(1) = .Fields(1)
            Call View
        End If
    End With
End If
End Sub


Private Sub dtp_Change(Index As Integer)
If dtp(0) <= dtp(1) Then Call View
End Sub

Private Sub Form_Load()
cmd(0).Picture = LoadResPicture("Salir", vbResIcon)
cmd(1).Picture = LoadResPicture("Print", vbResIcon)
'
dtp(0).Value = DateAdd("yyyy", -10, Date)
dtp(1).Value = Date
Set Flex.FontFixed = LetraTitulo(LoadResString(527), 9, , True)
Set Flex.Font = LetraTitulo(LoadResString(528), 8)
Flex.RowHeight(0) = 315
Call centra_titulo(Flex, True)
End Sub

Private Sub Opt_Click(Index As Integer)
'
If Opt(0) Then
    dtp(0).Enabled = False
    dtp(1).Enabled = False
Else
    dtp(0).Enabled = True
    dtp(1).Enabled = True
End If
'
Call View


'
End Sub

    '---------------------------------------------------------------------------------------------
    '   Rutina:     Report
    '
    '   Imprime el reporte de interes descontados
    '---------------------------------------------------------------------------------------------
    Private Sub Report()
    'variables locales
    Dim ctlReport As ctlReport
    Dim strSQL As String
    
    '
    On Error GoTo salir
    '
    strSQL = "SELECT * FROM DetFAct IN '" & gcPath & "\" & dtc(0) & "\inm.mdb' WHERE CodGAsto " _
    & "IN (SELECT CodDedInt FROM Inmueble WHERE CodInm='" & dtc(0) & "')"
    
    If Opt(1) Then
        strSQL = strSQL & " AND (Fecha >=#" & Format(dtp(0), "mm/dd/yy") & "# and Fecha <= #" _
        & Format(dtp(1), "mm/dd/yy") & "#)"
    End If
    Call rtnGenerator(gcPath & "\sac.mdb", strSQL, "qdfIntDes")
    Set ctlReport = New ctlReport
    '
    With ctlReport
        .OrigenDatos(0) = gcPath & "\sac.mdb"
        .Reporte = gcReport + "cxc_intdes.rpt"
        .Formulas(0) = "sUBTITULO='" & dtc(0) & " - " & dtc(1) & "'"
        'Destino
        If Opt(6) Then
            .Salida = crPantalla
            .TituloVentana = "Intreses Descontados..."
        Else
            .Salida = crImpresora
        End If
        Call rtnBitacora("Imprimir reporte descuento de intereses inm." & dtc(0))
        'errlocal = .PrintReport
        .Imprimir
    End With
    Set ctlReport = Nothing
salir:
'    If errlocal <> 0 Then
'        Call rtnBitacora("Se produjo el error " & ctlReport.LastErrorNumber)
'        MsgBox ctlReport.LastErrorString, , "Error " & ctlReport.LastErrorNumber
'    End If
    End Sub
    
    '---------------------------------------------------------------------------------------------
    '   Rutina:     View
    '
    '   Muestra en pantalla la información requerida por el usuario
    '---------------------------------------------------------------------------------------------
    Private Sub View()
    'variables locales
    Dim rstIntDes As New ADODB.Recordset
    Dim strSQL As String
    
    '
    MousePointer = vbArrowHourglass
    
    strSQL = "SELECT * FROM DetFAct IN '" & gcPath & "\" & dtc(0) & "\inm.mdb' WHERE CodGAsto " _
    & "IN (SELECT CodDedInt FROM Inmueble WHERE CodInm='" & dtc(0) & "')"
    rstIntDes.Open strSQL, cnnConexion, adOpenStatic, adLockReadOnly, adCmdText
    
    If Opt(1) Then
        rstIntDes.Filter = "Fecha >= #" & dtp(0) & "# AND Fecha <= #" & dtp(1) & "#"
    End If
    '
    If Not rstIntDes.EOF And Not rstIntDes.BOF Then
    
        Flex.Rows = rstIntDes.RecordCount + 1
        rstIntDes.MoveFirst: lbl(2) = "0,00"
        '
        With Flex
        '
            'Set .DataSource = rstIntDes
            'Set FLEX.Recordset = rstIntDes
            'Call centra_titulo(FLEX, True)
            
            .ColAlignment(0) = flexAlignCenterCenter
            .ColAlignment(1) = flexAlignCenterCenter
            .ColAlignment(3) = flexAlignRightCenter
            Do
                .TextMatrix(rstIntDes.AbsolutePosition, 0) = rstIntDes("Fecha")
                .TextMatrix(rstIntDes.AbsolutePosition, 1) = rstIntDes("Codigo")
                .TextMatrix(rstIntDes.AbsolutePosition, 2) = rstIntDes("Detalle")
                .TextMatrix(rstIntDes.AbsolutePosition, 3) = Format(rstIntDes("Monto"), "#,##0.00 ")
                lbl(2) = Format(CCur(lbl(2)) + rstIntDes("Monto"), "#,##0.00")
                rstIntDes.MoveNext
                
            Loop Until rstIntDes.EOF
            
            '
        End With
    Else
        Call rtnLimpiar_Grid(Flex)
    End If
    '
    rstIntDes.Close
    Set rstIntDes = Nothing
    '
    MousePointer = vbDefault
    End Sub

'    '---------------------------------------------------------------------------------------------
'    '   Función:        strSQl
'    '
'    '   Devuelve la cadena SQL que obtiene el conjunto de registros solicitados
'    '   por el usuario según los parámtros establecidos
'    '---------------------------------------------------------------------------------------------
'    Function strSQL(Optional Comodin$) As String
'    'variables locales
'    Dim dat1 As Date, dat2 As Date
'    Dim strFecha$, strOrden$, strMonto$, i%
'    '----------
'    If Comodin = "" Then Comodin = "*"
'    If Comodin = "*" Then
'        strFecha = "Deducciones.FecReg"
'        strMonto = "Deducciones.Monto"
'    Else
'        strFecha = "Format(Deducciones.FecReg,'dd/mm/yyyy') as Fecha"
'        strMonto = "Format(Deducciones.Monto,'#,##0.00 ')"
'    End If
'    '
'    dat1 = Format(dtp(0), "mm/dd/yyyy")
'    dat2 = Format(dtp(1), "mm/dd/yyyy")
'    For i = 2 To 5
'        If opt(i).Value = True Then strOrden = opt(i).Tag: Exit For
'    Next i
'    '
'    strSQL = "SELECT " & strFecha & ",MovimientoCaja.aptoMovimientoCaja as Apto, Deducciones.ti" _
'    & "tulo," & strMonto & ",MovimientoCaja.InmuebleMovimientoCaja,Deducci" _
'    & "ones.CodGasto FROM (MovimientoCaja INNER JOIN Periodos ON MovimientoCaja.IDRecibo = Peri" _
'    & "odos.IDRecibo) INNER JOIN Deducciones ON Periodos.IDPeriodos = Deducciones.IDPeriodos Wh" _
'    & "ere Deducciones.Titulo Like '" & Comodin & "INT" & Comodin & "' AND MovimientoCaja.Fecha" _
'    & "MovimientoCaja Between #" & dat1 & "# AND #" & dat2 & "# AND MovimientoCaja.InmuebleMovi" _
'    & "mientoCaja='" & dtc(0) & "' ORDER BY " & strOrden
'    '
'    End Function
