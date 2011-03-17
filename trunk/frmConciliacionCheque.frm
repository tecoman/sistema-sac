VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmConciliacionCheque 
   Caption         =   "Conciliación Cheque"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   Tag             =   "1"
   WindowState     =   2  'Maximized
   Begin VB.Frame fraLBanco 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   1440
      Index           =   0
      Left            =   540
      TabIndex        =   11
      Top             =   6780
      Width           =   10305
      Begin VB.PictureBox pic 
         AutoRedraw      =   -1  'True
         Height          =   465
         Left            =   135
         ScaleHeight     =   405
         ScaleWidth      =   6345
         TabIndex        =   12
         Top             =   510
         Visible         =   0   'False
         Width           =   6405
      End
      Begin VB.CommandButton cmdLBanco 
         Caption         =   "Conciliar"
         Height          =   1020
         Index           =   3
         Left            =   6735
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   240
         Width           =   1170
      End
      Begin VB.CommandButton cmdLBanco 
         Caption         =   "Salir"
         Height          =   1020
         Index           =   1
         Left            =   9120
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   240
         Width           =   1170
      End
      Begin VB.CommandButton cmdLBanco 
         Caption         =   "Imprimir"
         Height          =   1020
         Index           =   0
         Left            =   7935
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   240
         Width           =   1170
      End
   End
   Begin VB.Frame fraLBanco 
      Height          =   915
      Index           =   1
      Left            =   495
      TabIndex        =   0
      Top             =   150
      Width           =   10290
      Begin VB.TextBox txt 
         Height          =   315
         Left            =   7170
         Locked          =   -1  'True
         TabIndex        =   5
         Top             =   345
         Width           =   2490
      End
      Begin VB.CommandButton cmdLBanco 
         Caption         =   "..."
         Height          =   315
         Index           =   2
         Left            =   9675
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   345
         Width           =   345
      End
      Begin MSDataListLib.DataCombo dtcCuentas 
         DataField       =   "NumCuenta"
         Height          =   315
         Index           =   0
         Left            =   3360
         TabIndex        =   3
         Top             =   360
         Width           =   2805
         _ExtentX        =   4948
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         ListField       =   "NumCuenta"
         BoundColumn     =   "NombreBanco"
         Text            =   "DataCombo1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDataListLib.DataCombo dtcCuentas 
         DataField       =   "NombreBanco"
         Height          =   315
         Index           =   1
         Left            =   945
         TabIndex        =   2
         Top             =   360
         Width           =   2400
         _ExtentX        =   4233
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         ListField       =   "NombreBanco"
         BoundColumn     =   "NumCuenta"
         Text            =   "DataCombo1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label lbl 
         BackStyle       =   0  'Transparent
         Caption         =   "Edo.Cta.:"
         Height          =   300
         Index           =   1
         Left            =   6435
         TabIndex        =   4
         Top             =   405
         Width           =   930
      End
      Begin VB.Label lbl 
         BackStyle       =   0  'Transparent
         Caption         =   "Cuenta:"
         Height          =   300
         Index           =   0
         Left            =   105
         TabIndex        =   1
         Top             =   390
         Width           =   930
      End
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid gridLBanco 
      Height          =   5295
      Left            =   525
      TabIndex        =   10
      Top             =   1365
      Width           =   10320
      _ExtentX        =   18203
      _ExtentY        =   9340
      _Version        =   393216
      Cols            =   5
      RowHeightMin    =   100
      BackColorFixed  =   -2147483646
      ForeColorFixed  =   -2147483639
      BackColorBkg    =   -2147483636
      GridColor       =   -2147483633
      ScrollTrack     =   -1  'True
      GridLineWidthFixed=   2
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
      FontWidthFixed  =   0
      _NumberOfBands  =   1
      _Band(0).Cols   =   5
      _Band(0).GridLineWidthBand=   1
   End
   Begin MSComDlg.CommonDialog dlg 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Image img 
      Height          =   240
      Index           =   0
      Left            =   420
      Picture         =   "frmConciliacionCheque.frx":0000
      Top             =   1095
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image img 
      Height          =   240
      Index           =   1
      Left            =   900
      Picture         =   "frmConciliacionCheque.frx":01D8
      Top             =   1095
      Visible         =   0   'False
      Width           =   240
   End
End
Attribute VB_Name = "frmConciliacionCheque"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rstLBanco(1) As New ADODB.Recordset
Dim cnnLBanco As New ADODB.Connection
Dim procesado As Boolean

Private Sub cmdLBanco_Click(Index As Integer)

Select Case Index
    
    Case 0 ' Imprimir
        imprimirReporteConciliacion
    
    Case 1 ' Cerrar formulario
        Unload Me
        Set frmConciliacionCheque = Nothing
        
    Case 2 ' Cargar archivo excel
        
        On Error Resume Next
        dlg.CancelError = True
        dlg.Filter = "Estado de Cuenta (*.xls)|*.xls"
        dlg.FilterIndex = 1
        dlg.DialogTitle = "Estado de Cuenta Formato excel"
        dlg.ShowOpen
        txt.Text = dlg.FileTitle
        txt.Tag = dlg.FileName
             
    Case 3 ' Ejecutar la conciliacion
        conciliacion
End Select
'
End Sub

Private Sub DtcCuentas_Click(Index As Integer, Area As Integer)
If Area = 2 Then
    dtcCuentas(IIf(Index = 0, 1, 0)) = dtcCuentas(Index).BoundText
    mostrarChequesSinConciliar (dtcCuentas(0).Text)
End If
End Sub

Private Sub Form_Load()
Dim strSQL$    'variables locales
cnnLBanco.Open cnnOLEDB + mcDatos   'abre la conexón al orígen de datos
'
strSQL = "SELECT Cuentas.*, Bancos.NombreBanco FROM Bancos INNER JOIN " & _
        "Cuentas ON Bancos.IDBanco" _
        & "= Cuentas.IDBanco;"
rstLBanco(1).Open strSQL, cnnLBanco, adOpenStatic, adLockReadOnly, adCmdText
'
Set dtcCuentas(0).RowSource = rstLBanco(1)
Set dtcCuentas(1).RowSource = rstLBanco(1)
img(0).Picture = LoadResPicture("UnChecked", vbResBitmap)
img(1).Picture = LoadResPicture("Checked", vbResBitmap)
gridLBanco.FormatString = "^Número|Beneficiario|^Fecha|Monto|Conciliado"
gridLBanco.Tag = "1000|4500|1200|1200|1000"
Call centra_titulo(gridLBanco, True)

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If procesado Then
    If Respuesta("¿Desea completar esta conciliación?" & vbCrLf & _
    "Presione SI para registrar los cheques conciliados.") Then
        cnnConexion.CommitTrans: procesado = False
    Else
        cnnConexion.RollbackTrans: procesado = False
    End If
End If
End Sub

Private Sub Form_Resize()
gridLBanco.Height = Me.Height - gridLBanco.Top - fraLBanco(0).Height - fraLBanco(1).Height - fraLBanco(1).Top
fraLBanco(0).Top = 150 + gridLBanco.Top + gridLBanco.Height
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set rstLBanco(0) = Nothing
Set rstLBanco(1) = Nothing
cnnLBanco.Close
Set ObjCnn = Nothing
End Sub


Private Sub conciliacion()
If Not procesado Then
    cnnConexion.BeginTrans: procesado = True
End If
On Error GoTo Salir:
' validamos que se haya seleccionado un archivo
If Trim(txt.Text) = "" Then
    MsgBox "Seleccione la ubicación del archivo que contiene los movimientos bancarios", vbInformation, App.ProductName
    Exit Sub
End If

' validamos que el archivo existe en la ubicación seleccionada
If (Dir(Trim(txt), vbArchive) = "") Then
    MsgBox "Verifique que el archivo seleccionado existe en esta ubicación: " & txt, vbInformation, App.ProductName
    Exit Sub
End If
For n = 0 To 3
    cmdLBanco(n).Enabled = False
Next

' cargamos en memoria el archivo de movimientos
Dim conexion As ADODB.Connection
Dim rst As ADODB.Recordset
Dim hoja As String
Dim Cheque As String, Monto As Double
Set conexion = New ADODB.Connection
Set rst = New ADODB.Recordset

conexion.Open "Provider=Microsoft.Jet.OLEDB.4.0;" & _
    "Data Source=" & txt & _
    ";Extended Properties=""Excel 8.0;HDR=Yes;"""

Set rst = conexion.OpenSchema(adSchemaTables)
If Not (rst.EOF And rst.BOF) Then
    rst.MoveFirst
    hoja = rst("TABLE_NAME")
End If
rst.Close

If hoja = "" Then
    MsgBox "El Libro " & txt & " no tiene hojas de datos", vbCritical, App.ProductName
    Exit Sub
End If

Set rst = New ADODB.Recordset
rst.Open "Select * from [" & hoja & "]", conexion, adOpenStatic, adLockReadOnly


rst.Filter = "Descripción Like 'CH%'"

If Not (rst.EOF And rst.BOF) Then
    rst.MoveFirst
    pic.Visible = True
    Do
        Call UpdateStatus(pic, CLng(rst.AbsolutePosition * 100 / rst.RecordCount) + 1, 2)
        If Not Left(rst("descripción"), 6) = "CH.DEV" Then
            Cheque = Right(rst("referencia"), 6)
            Monto = CDbl(rst("monto") * -1)
            marcarChequeConciliado Cheque, Monto
        End If
        rst.MoveNext
    Loop Until rst.EOF
    pic.Visible = False
End If
rst.Close
Set rst = Nothing
conexion.Close
Set conexion = Nothing
Salir:

For n = 0 To 3
    cmdLBanco(n).Enabled = True
Next
If Err.Description <> "" Then
    Dim msg As String
    msg = Err.Description
    'If Err.Description = "No se pudo descifrar el archivo." Then
        msg = msg & vbCrLf & "El archivo no se puede leer porque el libro, " & _
        "o la hoja esta protegido con contraseña." & vbCrLf & _
        "- Elimine la contraseña, o bien" & vbCrLf & _
        "- Abra el archivo " & txt & vbCrLf & _
        "desde excel, e intente nuevamente la operación."
        
    'End If
    pic.Visible = False
    MsgBox msg, vbCritical, App.ProductName
    If procesado Then cnnConexion.RollbackTrans: procesado = False
End If
End Sub

Private Sub mostrarChequesSinConciliar(Cuentas As String)
Dim rst As ADODB.Recordset
Dim I As Integer

Set rst = ejecutar_procedure("procChequeConciliacion", Cuentas, 0)
Call rtnLimpiar_Grid(gridLBanco)
If Not (rst.EOF And rst.BOF) Then
    gridLBanco.Rows = rst.RecordCount + 1
    rst.MoveFirst
    I = 1
    Do
        gridLBanco.TextMatrix(I, 0) = Format(rst("IDCheque"), "000000")
        gridLBanco.TextMatrix(I, 1) = rst("Beneficiario")
        gridLBanco.TextMatrix(I, 2) = rst("FechaCheque")
        gridLBanco.TextMatrix(I, 3) = Format(rst("Monto"), "#,##0.00")
        gridLBanco.Row = I
        gridLBanco.Col = 4
        gridLBanco.CellPictureAlignment = flexAlignCenterCenter
        Set gridLBanco.CellPicture = img(0)
        I = I + 1
        rst.MoveNext
    Loop Until rst.EOF
End If
rst.Close
Set rst = Nothing
End Sub

Private Sub marcarChequeConciliado(Cheque As String, Monto As Double)
Dim I As Integer, n As Integer
Dim sql As String
With gridLBanco
    For I = 1 To .Rows - 1
        ' validamos la coincidencia de número y monto del cheque
        If .TextMatrix(I, 0) = Cheque And CDbl(.TextMatrix(I, 3)) = Monto Then
            .Col = 4
            .Row = I
            Set gridLBanco.CellPicture = img(1)
            resaltarLinea I
            sql = "update cheque set conciliado = -1,fechaConciliado=date() where idcheque= " & CLng(Cheque)
            cnnConexion.Execute sql, n
            If n > I Then rtnBitacora n & " Cheque #" & Cheque & " conciliado."
            Exit For
        End If
    Next
End With
End Sub

Private Sub UpdateStatus(pic As PictureBox, ByVal sngPercent As Single, _
    Optional ByVal fBorderCase)
    Dim strPercent As String
    Dim intX As Integer
    Dim intY As Integer
    Dim intWidth As Integer
    Dim intHeight As Integer
    
    If sngPercent > 100 Then sngPercent = 100
    If IsMissing(fBorderCase) Then fBorderCase = False
    
    Const colBackground = &HFFFFFF  ' white
    Const colForeground = &HFF00&   'verde manzana

    pic.ForeColor = vbBlack
    pic.BackColor = colBackground
    
    '
    Dim intPercent
    intPercent = sngPercent
    
    If intPercent = 0 Then
        If Not fBorderCase Then
            intPercent = 1
        End If
    ElseIf intPercent = 100 Then
        If Not fBorderCase Then
            intPercent = 99
        End If
    End If
    
    strPercent = Format$(intPercent) & "%"
    intWidth = pic.TextWidth(strPercent)
    intHeight = pic.TextHeight(strPercent)

    intX = pic.Width / 2 - intWidth / 2
    intY = (pic.Height / 2) - (intHeight / 2) - 35

    If sngPercent > 0 Then
        pic.Line (0, 0)-(sngPercent * pic.Width / 100, pic.Height), colForeground, BF
    Else
        pic.Line (0, 0)-(pic.Width, pic.Height), colForeground, BF
    End If
    pic.DrawMode = 13 ' Copy Pen
    '
    pic.CurrentX = intX
    pic.CurrentY = intY
    pic.Print strPercent

    pic.Refresh
End Sub


Private Sub resaltarLinea(Linea As Integer)
Const conciliado = &HC0FFFF ' amarillo
Const anulado = &H8080FF    'rojo
gridLBanco.Col = 1
gridLBanco.Row = Linea
gridLBanco.ColSel = gridLBanco.Cols - 1
gridLBanco.FillStyle = flexFillRepeat
gridLBanco.CellBackColor = IIf(gridLBanco.TextMatrix(Linea, 1) = "ANULADO", anulado, conciliado)
gridLBanco.FillStyle = flexFillSingle
gridLBanco.Col = 1
End Sub


Private Sub imprimirReporteConciliacion()
'
Dim strSQL$
Dim rpReporte As ctlReport
Dim rst As ADODB.Recordset
'
'Call clear_Crystal(ctlReport)
Set rpReporte = New ctlReport
Set rst = ModGeneral.ejecutar_procedure("procChequeConciliacion", Me.dtcCuentas(0), 0)
With rpReporte
    .Reporte = gcReport + "conciliacion-cheque.rpt"
    '.OrigenDatos(0) = gcPath & "\sac.mdb"
    .TituloVentana = "Cheque - Conciliación"
    .Salida = crPantalla
    .Imprimir , rst
    Call rtnBitacora("Print Libro Banco Inm.:" & gcCodInm)
    
End With
Set rpReporte = Nothing
'

End Sub
