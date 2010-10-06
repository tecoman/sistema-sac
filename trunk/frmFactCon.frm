VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmFactCon 
   Caption         =   "Consecutivos"
   ClientHeight    =   3135
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   4680
   ControlBox      =   0   'False
   FillColor       =   &H80000018&
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   3135
   ScaleWidth      =   4680
   Tag             =   "1"
   WindowState     =   2  'Maximized
   Begin VB.Frame fraApago 
      Height          =   1245
      Index           =   0
      Left            =   7770
      TabIndex        =   8
      Top             =   6555
      Width           =   3780
      Begin VB.CommandButton cmd 
         Caption         =   "Guardar"
         Height          =   945
         Index           =   2
         Left            =   1290
         Style           =   1  'Graphical
         TabIndex        =   11
         Tag             =   "Guardar"
         Top             =   210
         Width           =   1170
      End
      Begin VB.CommandButton cmd 
         Caption         =   "Salir"
         Height          =   945
         Index           =   1
         Left            =   2475
         Style           =   1  'Graphical
         TabIndex        =   10
         Tag             =   "Salir"
         Top             =   210
         Width           =   1170
      End
      Begin VB.CommandButton cmd 
         Caption         =   "Imprimir"
         Height          =   945
         Index           =   0
         Left            =   105
         Style           =   1  'Graphical
         TabIndex        =   9
         Tag             =   "Print"
         Top             =   210
         Width           =   1170
      End
   End
   Begin VB.Frame fraApago 
      Caption         =   "Filtrar:"
      Height          =   1245
      Index           =   10
      Left            =   465
      TabIndex        =   3
      Top             =   6555
      Width           =   7230
      Begin VB.TextBox Txt 
         Height          =   285
         Index           =   2
         Left            =   5970
         TabIndex        =   16
         Top             =   870
         Width           =   1065
      End
      Begin VB.TextBox Txt 
         Height          =   285
         Index           =   1
         Left            =   5970
         TabIndex        =   15
         Top             =   570
         Width           =   1065
      End
      Begin VB.TextBox Txt 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
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
         Height          =   285
         Index           =   0
         Left            =   5970
         Locked          =   -1  'True
         TabIndex        =   13
         Text            =   "0"
         Top             =   270
         Width           =   1065
      End
      Begin MSDataListLib.DataCombo dtc 
         Height          =   315
         Index           =   0
         Left            =   960
         TabIndex        =   4
         ToolTipText     =   "Banco"
         Top             =   315
         Width           =   3015
         _ExtentX        =   5318
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         ListField       =   "NombreBanco"
         BoundColumn     =   "NumCuenta"
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo dtc 
         Height          =   315
         Index           =   1
         Left            =   960
         TabIndex        =   5
         ToolTipText     =   "Número de Cuenta"
         Top             =   690
         Width           =   3015
         _ExtentX        =   5318
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         ListField       =   "NumCuenta"
         BoundColumn     =   "NombreBanco"
         Text            =   ""
      End
      Begin VB.Label LblAPago 
         Alignment       =   1  'Right Justify
         Caption         =   "Fecha Cheque:"
         Height          =   255
         Index           =   4
         Left            =   4290
         TabIndex        =   17
         Top             =   885
         Width           =   1515
      End
      Begin VB.Label LblAPago 
         Alignment       =   1  'Right Justify
         Caption         =   "Cod. Inmueble"
         Height          =   255
         Index           =   3
         Left            =   4770
         TabIndex        =   14
         Top             =   600
         Width           =   1035
      End
      Begin VB.Label LblAPago 
         Alignment       =   1  'Right Justify
         Caption         =   "Total Cheques:"
         Height          =   255
         Index           =   0
         Left            =   4455
         TabIndex        =   12
         Top             =   315
         Width           =   1350
      End
      Begin VB.Label LblAPago 
         Caption         =   "Cuenta:"
         Height          =   255
         Index           =   1
         Left            =   255
         TabIndex        =   7
         Top             =   720
         Width           =   705
      End
      Begin VB.Label LblAPago 
         Caption         =   "Banco:"
         Height          =   255
         Index           =   2
         Left            =   255
         TabIndex        =   6
         Top             =   345
         Width           =   705
      End
   End
   Begin TabDlg.SSTab Ficha 
      Height          =   8010
      Left            =   195
      TabIndex        =   0
      Top             =   210
      Width           =   11550
      _ExtentX        =   20373
      _ExtentY        =   14129
      _Version        =   393216
      Tabs            =   2
      TabHeight       =   520
      TabCaption(0)   =   "Pendientes"
      TabPicture(0)   =   "frmFactCon.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "grid(0)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Verificados"
      TabPicture(1)   =   "frmFactCon.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "grid(1)"
      Tab(1).ControlCount=   1
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid grid 
         Height          =   5655
         Index           =   0
         Left            =   285
         TabIndex        =   1
         Tag             =   "900|1500|1200|4000|1200|1200|600"
         Top             =   525
         Width           =   11085
         _ExtentX        =   19553
         _ExtentY        =   9975
         _Version        =   393216
         BackColor       =   16777215
         Cols            =   7
         FixedCols       =   0
         BackColorFixed  =   -2147483646
         ForeColorFixed  =   -2147483639
         BackColorSel    =   65280
         ForeColorSel    =   -2147483646
         BackColorBkg    =   -2147483636
         GridColor       =   -2147483636
         AllowBigSelection=   0   'False
         FocusRect       =   0
         HighLight       =   2
         GridLinesFixed  =   1
         SelectionMode   =   1
         MergeCells      =   2
         MousePointer    =   99
         GridLineWidthFixed=   1
         FormatString    =   "^Cheque |>Monto |^Fecha |Beneficiario |Beneficiario |Beneficiario |Verif."
         BandDisplay     =   1
         RowSizingMode   =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FontWidthFixed  =   0
         MouseIcon       =   "frmFactCon.frx":0038
         _NumberOfBands  =   1
         _Band(0).Cols   =   7
         _Band(0).GridLineWidthBand=   1
         _Band(0).TextStyleBand=   0
         _Band(0).TextStyleHeader=   0
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid grid 
         Height          =   5760
         Index           =   1
         Left            =   -74715
         TabIndex        =   2
         Tag             =   "900|1500|1200|4000|1200|1200|600"
         Top             =   525
         Width           =   11085
         _ExtentX        =   19553
         _ExtentY        =   10160
         _Version        =   393216
         BackColor       =   16777215
         Cols            =   7
         FixedCols       =   0
         BackColorFixed  =   -2147483646
         ForeColorFixed  =   -2147483639
         BackColorSel    =   65280
         ForeColorSel    =   -2147483646
         BackColorBkg    =   -2147483636
         GridColor       =   -2147483636
         AllowBigSelection=   0   'False
         FocusRect       =   0
         HighLight       =   2
         GridLinesFixed  =   1
         SelectionMode   =   1
         MergeCells      =   2
         MousePointer    =   99
         GridLineWidthFixed=   1
         FormatString    =   "^Cheque |>Monto |^Fecha |Beneficiario |Beneficiario |Beneficiario |Verif."
         BandDisplay     =   1
         RowSizingMode   =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FontWidthFixed  =   0
         MouseIcon       =   "frmFactCon.frx":019A
         _NumberOfBands  =   1
         _Band(0).Cols   =   7
         _Band(0).GridLineWidthBand=   1
         _Band(0).TextStyleBand=   0
         _Band(0).TextStyleHeader=   0
      End
   End
   Begin VB.Image Img 
      Height          =   255
      Index           =   1
      Left            =   255
      Tag             =   "Checked"
      Top             =   255
      Width           =   375
   End
   Begin VB.Image Img 
      Height          =   255
      Index           =   0
      Left            =   630
      Tag             =   "UnChecked"
      Top             =   225
      Width           =   375
   End
End
Attribute VB_Name = "frmFactCon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'-------------------------------------------------------------------------------------------------
'   Módulo Facturación      #22/07/2003#
'   Ventana verificación de consecutivos
'
'-------------------------------------------------------------------------------------------------
'Variables Publicas a nivel de módulo
Dim rstFC(1) As ADODB.Recordset
Dim rstCta As New ADODB.Recordset
Const Moneda$ = "#,##0.00 "
Dim Nivel&

Private Sub Cmd_Click(Index As Integer)
Select Case Index
    Case 0  'imprimir
    If dtc(0) <> "" And dtc(1) <> "" Then Call Printer_Report
    'Salir
    Case 1: Unload Me: Set frmFactCon = Nothing
    'Guardar cambios
    Case 2
        If Respuesta("¿Desea guardar todos los cambios realizados?") Then
            cnnConexion.CommitTrans 'procesa la transacciòn
        Else
            cnnConexion.RollbackTrans   'deshace los cambios efectuados
        End If
        Call Muestra_Datos(Ficha.Tab)
    '
End Select

End Sub

Private Sub dtc_Click(Index As Integer, Area As Integer)
Dim i As Integer
If Area = 2 Then
    If Index = 0 Then
        dtc(1) = dtc(0).BoundText
    Else
        dtc(0) = dtc(1).BoundText
    End If
    If dtc(1) <> "" Then
        'For i = 0 To 28000 'retardo
        'Next i
        'rstFC(0).Filter = "Cuenta='" & dtc(1) & "'"
        Call Muestra_Datos(Ficha.Tab)
    End If
End If
End Sub



Private Sub Ficha_Click(PreviousTab As Integer)
Static Cta(1) As String   'variables locales
If Cta(Ficha.Tab) <> dtc(1) Then Call Muestra_Datos(Ficha.Tab)
Cta(Ficha.Tab) = dtc(1)
txt(0) = rstFC(Ficha.Tab).RecordCount
End Sub

Private Sub Form_Load()
'
Dim strSql(1) As String, strCondicion(1) As String, i%
'Abre el objeto ADODB.Recordset
rstCta.Open "SELECT Cuentas.NumCuenta, Bancos.NombreBanco FROM Bancos INNER JOIN Cuentas ON Ban" _
& "cos.IDBanco = Cuentas.IDBanco WHERE Cuentas.Inactiva=False", cnnOLEDB + mcDatos, _
adOpenKeyset, adLockOptimistic, adCmdText
'
For i = 0 To 1
    '
    Set rstFC(i) = New ADODB.Recordset
    'Asigna la propiedad picture al objeto img
    img(i).Picture = LoadResPicture(img(i).Tag, vbResBitmap)
    '
    strCondicion(i) = "(SELECT NumCuenta FROM Cuentas IN '" & mcDatos & "') AND Verif=" & _
    IIf(i = 0, "False", "True")
    '
    strSql(i) = "SELECT TOP 400 Cheque.IDCheque,Cheque.FechaCheque,Cheque.Beneficiario,Cheque.Concepto," _
    & "Cheque.Verif,ChequeDetalle.Cargado,ChequeDetalle.CodInm,ChequeDetalle.CodGasto,ChequeDet" _
    & "alle.Detalle,ChequeDetalle.Monto,Cheque.Cuenta,Cheque.Fecha,Cheque.Hora FROM Cheque INNE" _
    & "R JOIN ChequeDetalle ON Cheque.IDCheque=ChequeDetalle.IDCheque WHERE Cheque.Cuenta IN " _
    & strCondicion(i) & " ORDER BY Cheque.IDCheque DESC UNION SELECT TOP 400 IDCheque,FechaCheque,Beneficiario,'ANULADO',Verif,Fecha," _
    & "CodInm,'',Concepto,Monto,Cuenta,Fecha,Hora FROM ChequeAnulado WHERE Cuenta IN " & _
    strCondicion(i) & " ORDER BY IDCheque DESC;"
    '
    'rstFC(i).MaxRecords = 10
    rstFC(i).Open strSql(i), cnnConexion, adOpenForwardOnly, adLockOptimistic, adCmdText
    rstFC(i).Sort = "FechaCheque,IDCheque"
    '
    Call centra_titulo(grid(i), True)
    grid(i).MergeRow(0) = True
    Set dtc(i).RowSource = rstCta
    '
Next i
'
For i = 0 To 2: cmd(i).Picture = LoadResPicture(cmd(i).Tag, vbResIcon)
Next
'abre una transacción
cnnConexion.BeginTrans
'
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
rstFC(0).Close
rstFC(1).Close
rstCta.Close
Set rstFC(0) = Nothing
Set rstFC(1) = Nothing
Set rstCta = Nothing
'si existe una transacción activa
cnnConexion.RollbackTrans
End Sub

'-------------------------------------------------------------------------------------------------
'   Rutina: Muestra_datos
'
'   Muestra el conjunto de registros del recodset
'-------------------------------------------------------------------------------------------------
Private Sub Muestra_Datos(Optional GridIndex%)
'Variables locales
Dim Total@, i&, C&, r&, strF$, Reg&

On Error Resume Next
MousePointer = vbHourglass
grid(GridIndex).Visible = False
grid(GridIndex).Rows = 2
Call rtnLimpiar_Grid(grid(GridIndex))
'
With rstFC(GridIndex)
    '
    .Filter = 0
    If txt(1) <> "" Then
        strF = "Cuenta='" & dtc(1) & "' and CodInm='" & txt(1) & "'"
        If txt(2) <> "" Then strF = strF & " AND FechaCheque >=#" & txt(2) & "#"
    Else
        strF = "Cuenta='" & dtc(1) & "'"
        If txt(2) <> "" Then strF = strF & " AND FechaCHeque >=#" & txt(2) & "#"
    End If
    rstFC(GridIndex).Filter = strF
    .Sort = "FechaCheque,IDCheque"
    'Txt(0) = .RecordCount
    If Not .EOF Or Not .BOF Then
    .MoveFirst: i = 1
    Do
        grid(GridIndex).MergeRow(i) = True
        If C <> !IDCheque Then   'agrega el encabezado del cheque
        '
            'If !IDcheque = 661098 Then Stop
            Reg = Reg + 1
            If Total > 0 Then
                r = IIf(r = 0, i, r)
                grid(GridIndex).TextMatrix(r, 1) = Format(Total, Moneda)
                Total = 0
            End If
            Call Marcar_Linea(i, GridIndex) 'sombrea la linea
            grid(GridIndex).TextMatrix(i, 0) = !IDCheque
            grid(GridIndex).TextMatrix(i, 2) = Format(!FechaCheque, "dd/mm/yyyy")
            Call Textgrid(5, i, !Beneficiario, GridIndex)
            grid(GridIndex).Row = i
            grid(GridIndex).Col = 6
            '
            Set grid(GridIndex).CellPicture = img(GridIndex)
            grid(GridIndex).CellPictureAlignment = flexAlignCenterCenter
            r = i
            'Call Marcar_Linea(I)
            If !Concepto <> "ANULADO" Then
                i = 1 + i
                If grid(GridIndex).Rows = i Then grid(GridIndex).AddItem ("")
                grid(GridIndex).MergeRow(i) = True
                Call Textgrid1(i, !Cargado, !Monto, !CodInm & " " & !codGasto, GridIndex)
                Call Textgrid(6, i, !Detalle, GridIndex)
            Else
                grid(GridIndex).TextMatrix(i, 1) = Format(!Monto, Moneda)
            End If
            '
        Else
            Call Textgrid1(i, !Cargado, !Monto, !CodInm & " " & !codGasto, GridIndex)
            Call Textgrid(6, i, !Detalle, GridIndex)
        End If
        '
        Total = !Monto + Total
        C = !IDCheque
        .MoveNext: i = i + 1: grid(GridIndex).AddItem ("")
    Loop Until .EOF
    '
    Else
        MsgBox "Cuenta sin movimientos", vbInformation, App.ProductName
    End If
    If Total > 0 Then grid(GridIndex).TextMatrix(r, 1) = Format(Total, Moneda)
End With
txt(0) = Reg
grid(GridIndex).Visible = True
MousePointer = vbDefault
'
End Sub

Private Sub Textgrid(K%, Row&, Texto$, Optional Indice%)
Dim i%
grid(Indice).ColAlignment(3) = flexAlignLeftCenter
If rstFC(Indice)!Concepto = "ANULADO" Then
    grid(Indice).TextMatrix(Row, K) = "(ANULADO)"
    K = K - 1
End If
For i = 3 To K: grid(Indice).TextMatrix(Row, i) = Texto
Next i
'
End Sub

Private Sub Textgrid1(Row&, Cargado$, Monto@, Cuenta$, Optional Indice%)
With grid(Indice)
    .TextMatrix(Row, 0) = Format(Cargado, "mm-yyyy")
    .TextMatrix(Row, 1) = Format(Monto, Moneda)
    .TextMatrix(Row, 2) = Cuenta
End With
End Sub

'-------------------------------------------------------------------------------------------------
'   Rutina: Marcar_linea
'
'   Entradas:   Row(entero) Nº de Línea, Opcional Indice(Propiedad Index) del
'               grid
'
'   Le da formato de Negritas y color de fondo a la fila determinada
'   por la variable Row
'-------------------------------------------------------------------------------------------------
Private Sub Marcar_Linea(Row&, Optional Indice%)
'
With grid(Indice)
    '
    .Row = Row
    .Col = 0
    .ColSel = .Cols - 1
    .FillStyle = flexFillRepeat
    .CellBackColor = &HFFFF80   'Color de Fondo
    .CellFontBold = True    'Letras negritas
    .FillStyle = flexFillSingle
    '
End With
'
End Sub


Private Sub grid_Click(Index As Integer)
'
Dim NumCheque&  'variables locales
Dim Si_No$
'
On Error GoTo salir
'
'If Index = 0 Then
    If grid(Index).ColSel = 6 And grid(Index).RowSel > 0 And Not grid(Index).TextMatrix(grid(Index).RowSel, 0) = "" Then
        'grid(Index).Row = grid(Index).RowSel
        grid(Index).Col = 6
        NumCheque = grid(Index).TextMatrix(grid(Index).Row, 0)
        If grid(Index).CellPicture = img(0) Then
            Set grid(Index).CellPicture = img(1)
            Si_No = "True"
        Else
            Set grid(Index).CellPicture = img(0)
            Si_No = "false"
        End If
        '
        cnnConexion.Execute "UPDATE Cheque SET Verif=" & Si_No & " WHERE IDCheque=" & NumCheque
        cnnConexion.Execute "UPDATE ChequeAnulado SET Verif=" & Si_No & " WHERE IDCheque=" & _
        NumCheque
        '
    End If
'End If
'
salir:
If Err.Number <> 0 Then
    MsgBox "Error " & Err.Number & vbCrLf & Err.Description, vbExclamation, App.ProductName
    Err.Clear
End If
'
End Sub

Private Sub txt_KeyPress(Index As Integer, KeyAscii As Integer)
'
If Index = 1 Then Call Validacion(KeyAscii, "0123456789")
If Index = 2 Then
    If KeyAscii = 45 Then KeyAscii = 47
    Call Validacion(KeyAscii, "0123456789/")
End If
'
If KeyAscii = 13 Then
    '
    If Index = 2 Then
        '
        If txt(2) <> "" Then
            If Not IsDate(txt(2)) Then
                 MsgBox "Introdujo una fecha no válida...", vbInformation, App.ProductName
                Exit Sub
            End If
        End If
    End If
    If dtc(0) = "" And dtc(1) = "" Then
        MsgBox "Seleccione la cuenta a consultar....", vbInformation, App.ProductName
        Exit Sub
    End If
    Call Muestra_Datos(Ficha.Tab)
End If
'
End Sub

'-------------------------------------------------------------------------------------------------
'   Rutina Print_Report
'
'   Imprime el reporte de los datos visualizados en pantalla
'-------------------------------------------------------------------------------------------------
Private Sub Printer_Report()
'varaibles locales
Dim strSql As String
Dim rpReporte As ctlReport
'genera la consulta
strSql = IIf(Ficha.Tab = 0, " AND Verif=False", " AND Verif=True")
If txt(1) <> "" Then strSql = strSql & " AND Codinm='" & txt(1) & "'"
If txt(2) <> "" Then strSql = strSql & " AND C.FechaCheque>=#" & Format(CDate(txt(2)), "mm/dd/yy") & "#"
'
strSql = "SELECT C.IDcheque,Sum(CD.Monto) AS Monto, C.FechaCheque,C.Beneficiario, C.Banco, C.Cuen" _
& "ta,C.Concepto FROM Cheque AS C INNER JOIN ChequeDetalle AS CD ON C.IDCheque = CD.IDCheque WH" _
& "ERE C.Banco='" & dtc(0) & "' AND C.Cuenta='" & dtc(1) & "'" & strSql & " GROUP BY C.IDCheque" _
& ", C.FechaCheque, C.Beneficiario, C.Banco,C.Cuenta,C.Concepto UNION SELECT IDCheque,Monto,Fec" _
& "haCheque,Beneficiario,Banco,Cuenta,'ANULADO' FROM ChequeAnulado as C WHERE C.Banco='" & _
dtc(0) & "' AND C.Cuenta='" & dtc(1) & "'" & strSql & " ORDER BY C.IDcheque DESC;"
'
Call rtnGenerator(gcPath & "\sac.mdb", strSql, "QfConsecutivoCondominio")
'
'Call clear_Crystal(FrmAdmin.rptReporte)
'
Set rpReporte = New ctlReport
With rpReporte
    .Reporte = gcReport + "consecutivos.rpt"
    .OrigenDatos(0) = gcPath & "\sac.mdb"
    .Formulas(0) = "SubTitulo='BANCO " & dtc(0) & " - Cuenta Nº: " & dtc(1) & " [" _
    & Ficha.TabCaption(Ficha.Tab) & "]'"
    If txt(1) <> "" Then .FormuladeSeleccion = "{ChequeDetalle.CodInm}='" & txt(1) & "'"
    If Respuesta(LoadResString(537)) Then
        '.Destination = crptToPrinter
        .Salida = crImpresora
    Else
        '.Destination = crptToWindow
        .Salida = crPantalla
    End If
    Call rtnBitacora("Printer Consecutivos.." & dtc(0) & "/" & dtc(1))
    'errlocal = .PrintReport
    .Imprimir
    '
'    If errlocal <> 0 Then   'si ocurre un error lo anota en bitácora
'        '
'        Call rtnBitacora("Error al imprimir el reporte. Error: " & .LastErrorNumber)
'        MsgBox "Error al imprimir el reporte." & .LastErrorString, vbExclamation, "Error " _
'        & .LastErrorNumber
'        '
'    End If
    '
End With
Set rpReporte = Nothing
'
End Sub
