VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmTransito 
   Caption         =   "Depósitos en Transito"
   ClientHeight    =   6855
   ClientLeft      =   1170
   ClientTop       =   2145
   ClientWidth     =   8190
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6855
   ScaleWidth      =   8190
   WindowState     =   2  'Maximized
   Begin VB.Frame fraTansito 
      Caption         =   "Buscar Depósito por:"
      Height          =   1050
      Left            =   690
      TabIndex        =   3
      Top             =   6630
      Width           =   7050
      Begin VB.CommandButton cmdRCG 
         Caption         =   "Buscar"
         Height          =   630
         Index           =   2
         Left            =   5460
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   270
         Width           =   1455
      End
      Begin VB.TextBox txtBusca 
         Height          =   315
         Left            =   3465
         TabIndex        =   6
         Top             =   428
         Width           =   1635
      End
      Begin VB.OptionButton optBusca 
         Caption         =   "Número"
         Height          =   330
         Index           =   1
         Left            =   2100
         TabIndex        =   5
         Top             =   420
         Value           =   -1  'True
         Width           =   1275
      End
      Begin VB.OptionButton optBusca 
         Caption         =   "Monto"
         Height          =   330
         Index           =   0
         Left            =   360
         TabIndex        =   4
         Top             =   420
         Width           =   1275
      End
   End
   Begin VB.CommandButton cmdRCG 
      Caption         =   "&Imprimir"
      Height          =   975
      Index           =   0
      Left            =   8160
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   6720
      Width           =   1455
   End
   Begin VB.CommandButton cmdRCG 
      Caption         =   "&Salir"
      Height          =   975
      Index           =   1
      Left            =   9720
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   6720
      Width           =   1455
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid gridTsto 
      Height          =   5625
      Left            =   690
      TabIndex        =   0
      Top             =   645
      Width           =   10560
      _ExtentX        =   18627
      _ExtentY        =   9922
      _Version        =   393216
      Cols            =   6
      FixedCols       =   0
      BackColorFixed  =   -2147483646
      ForeColorFixed  =   -2147483639
      ForeColorSel    =   16777215
      FocusRect       =   2
      HighLight       =   2
      MergeCells      =   2
      FormatString    =   "Caja del |Banco | Nº Depósito | Monto | Fecha | Verificado"
      _NumberOfBands  =   1
      _Band(0).Cols   =   6
      _Band(0).GridLinesBand=   2
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   0
   End
   Begin VB.Image Img 
      Height          =   195
      Index           =   0
      Left            =   0
      Stretch         =   -1  'True
      Top             =   225
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Image Img 
      Height          =   195
      Index           =   1
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
End
Attribute VB_Name = "frmTransito"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim objRst As New ADODB.Recordset
'(((C.IDDeposito)Is Null) AND
Const strSql$ = "SELECT C.FechaMov,C.Banco, C.Ndoc,C.Monto,Inmueble.Caja,C.FechaDoc,C.IDDeposit" _
& "o FROM TDFCheques  AS C INNER JOIN Inmueble ON C.CodInmueble = Inmueble.CodInm WHER" _
& "E  (((C.IDDeposito) Is Null) AND ((C.Fpago)='DEPOSITO')) UNION SELECT TDFCheques.FechaMov,D.Banco, D.IDDeposito, Sum(TDFCheques." _
& "Monto) AS SumaDeMonto, D.Caja, D.Fecha,D.Confirmado FROM TDFDepositos AS D INNER JOIN TDFChe" _
& "ques ON D.IDDeposito = TDFCheques.IDDeposito Where D.Confirmado =0 And D.Banco > '0' GROUP B" _
& "Y TdfCheques.FechaMov,D.Banco, D.IDDeposito, D.Caja, D.Fecha,D.Conf" _
& "irmado ORDER BY Inmueble.Caja,C.FechaMov DESC"
Const Modena$ = "#,##0.00 ; (#,##0.00) "

Private Sub cmdRCG_Click(Index As Integer)
'
Dim Columna As Integer
Dim I As Integer
Static k%
Select Case Index
    Case 0
        MousePointer = vbArrowHourglass
        Call rtnGenerator(gcPath & "\sac.mdb", strSql, "Depen")
        mcTitulo = "Reporte Depósitos Pendientes"
        mcReport = "caja_depen.Rpt"
        mcOrdAlfa = "+{Depen.Banco}"
        mcOrdCod = "+{Depen.FechaMov}"
        FrmReport.Show
        MousePointer = vbDefault
    Case 1: Unload Me   'Cerrar el formulario
    Case 2  'Buscar
        Columna = IIf(OptBusca(0), 3, 2)
        With gridTsto
            .SetFocus
            k = IIf(Columna = 3, IIf(k >= .Rows, 1, k) + 1, 2)
            For I = k To .Rows - 1
                
                If Trim(.TextMatrix(I, Columna)) = Trim(txtBusca) Then
                    .Col = Columna
                    .Row = I
                    .TopRow = I
                    Exit For
                End If
            Next
            k = I
        End With
End Select
'
End Sub

Private Sub Form_Load()
'
Dim J As Integer    'variables locales
Dim Caja As String * 2
'
img(0).Picture = LoadResPicture("Unchecked", vbResBitmap)
img(1).Picture = LoadResPicture("Checked", vbResBitmap)
cmdRCG(0).Picture = LoadResPicture("Print", vbResIcon)
cmdRCG(1).Picture = LoadResPicture("Salir", vbResIcon)
cmdRCG(2).Picture = LoadResPicture("Buscar", vbResIcon)
'----------
Call config_grid
objRst.CursorLocation = adUseClient
objRst.Open strSql, cnnConexion, adOpenKeyset, adLockOptimistic
'
With objRst
'
    If Not .EOF Or Not .BOF Then
    '
        .MoveFirst: J = 1
        gridTsto.Rows = .RecordCount + 1
        
        Do
            If Caja <> !Caja Then
                Caja = !Caja
                J = SubTitulo(J, Caja)
            End If
            gridTsto.TextMatrix(J, 0) = !FechaMov
            gridTsto.TextMatrix(J, 1) = !Banco
            gridTsto.TextMatrix(J, 2) = !NDoc
            gridTsto.TextMatrix(J, 3) = Format(!Monto, Modena)
            gridTsto.TextMatrix(J, 4) = !FechaDoc
            gridTsto.Row = J
            gridTsto.Col = 5
            Set gridTsto.CellPicture = img(0)
            gridTsto.CellPictureAlignment = flexAlignCenterCenter
            .MoveNext: J = J + 1
            
        Loop Until .EOF
        '
    End If
    '
    .Close
End With
Set objRst = Nothing
'
End Sub

Sub config_grid()
'
Dim I As Integer    'variables locales
'
With gridTsto
    '
    Set .FontFixed = LetraTitulo(LoadResString(527), 7.5, True)
    Set .Font = LetraTitulo(LoadResString(528), 8)
    .RowHeight(0) = 315
'    .Row = 0
'    For i = 0 To .Cols - 1
'        .Col = i
'        .CellAlignment = flexAlignCenterCenter
'    Next i
    Call centra_titulo(gridTsto)
    .ColWidth(0) = 1200 'Fecha del movimiento
    .ColWidth(1) = 3000  'Banco
    .ColWidth(2) = 2000  'Cuenta
    .ColWidth(3) = 1200  'Monto
    .ColWidth(4) = 1200  'Fecha
    '
End With
'
End Sub

Function SubTitulo(Fila As Integer, strNCaja As String) As Integer
'
Dim k As Integer    'Variables locales
'
With gridTsto
    .MergeRow(Fila) = True
    .Row = Fila
    For k = 0 To .Cols - 1
        .Col = k
        .TextMatrix(Fila, k) = "CAJA Nº " & strNCaja
        .CellFontBold = True
        .CellForeColor = vbTitleBarText
        .CellBackColor = vbActiveTitleBar
        .CellAlignment = flexAlignCenterCenter
    Next
    .AddItem ("")
End With
SubTitulo = Fila + 1
'
End Function

Private Sub Form_Unload(Cancel As Integer)
Set frmTransito = Nothing
End Sub

Private Sub gridTsto_Click()
'
Dim strSQLvalue  'variables locales
'
MousePointer = vbHourglass
With gridTsto
    '
    If .ColSel = .Cols - 1 Then 'si hace click en la última columna
        .Col = .Cols - 1
        .Row = .RowSel
        If .CellPicture = img(0) Then
            Set .CellPicture = img(1)
            strSQLvalue = -1
        Else
            Set .CellPicture = img(0)
            strSQLvalue = 0
        End If
        .CellPictureAlignment = flexAlignCenterCenter
        .Refresh
        '
        cnnConexion.Execute "UPDATE TDFCheques SET IDDeposito='" & strSQLvalue & "' WHERE Banco & " _
        & "Ndoc & 'DEPOSITO' & Monto  ='" & .TextMatrix(.Row, 1) & .TextMatrix(.Row, 2) _
        & "DEPOSITO" & CCur(.TextMatrix(.Row, 3)) & "';"
        '
        cnnConexion.Execute "UPDATE TDfDepositos SET Confirmado=" & strSQLvalue & " WHERE Banco & " _
        & "IDDeposito & 'DEPOSITO' ='" & .TextMatrix(.Row, 1) & .TextMatrix(.Row, 2) _
        & "DEPOSITO';"
    '
    End If
    '
End With
MousePointer = vbDefault
'
End Sub

Private Sub gridTsto_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 And gridTsto.Col = gridTsto.Cols - 1 Then gridTsto_Click
End Sub

Private Sub txtBusca_KeyPress(KeyAscii As Integer)
Call Validacion(KeyAscii, "0123456789,.")
If KeyAscii = 13 Then Call cmdRCG_Click(2)
End Sub
