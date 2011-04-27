VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmchequecontrol 
   Caption         =   "Control Cheques Emitidos"
   ClientHeight    =   7695
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12420
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7695
   ScaleWidth      =   12420
   WindowState     =   2  'Maximized
   Begin VB.Frame fra 
      Height          =   1275
      Index           =   0
      Left            =   240
      TabIndex        =   1
      Top             =   120
      Width           =   11670
      Begin VB.TextBox txt 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   0
         Left            =   9120
         TabIndex        =   5
         Text            =   "0,00"
         Top             =   300
         Width           =   2460
      End
      Begin VB.Frame Frame1 
         Height          =   555
         Left            =   8220
         TabIndex        =   2
         Top             =   630
         Width           =   3360
         Begin VB.CommandButton cmd 
            Caption         =   "Guardar"
            Height          =   375
            Index           =   1
            Left            =   75
            Picture         =   "frmchequecontrol.frx":0000
            TabIndex        =   4
            Top             =   135
            Width           =   1545
         End
         Begin VB.CommandButton cmd 
            Caption         =   "Salir"
            Height          =   375
            Index           =   0
            Left            =   1650
            TabIndex        =   3
            Top             =   135
            Width           =   1545
         End
      End
      Begin MSDataListLib.DataCombo dtc 
         DataField       =   "CodInm"
         Height          =   315
         Index           =   0
         Left            =   1020
         TabIndex        =   6
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
         Left            =   2130
         TabIndex        =   7
         Top             =   300
         Width           =   6825
         _ExtentX        =   12039
         _ExtentY        =   556
         _Version        =   393216
         ListField       =   "Nombre"
         BoundColumn     =   "CodInm"
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo dtc 
         Height          =   315
         Index           =   2
         Left            =   1020
         TabIndex        =   8
         Top             =   795
         Width           =   2355
         _ExtentX        =   4154
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
         Left            =   3435
         TabIndex        =   9
         Top             =   795
         Width           =   2220
         _ExtentX        =   3916
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         ListField       =   "NumCuenta"
         BoundColumn     =   "NombreBanco"
         Text            =   ""
      End
      Begin MSMask.MaskEdBox MskFecha 
         Bindings        =   "frmchequecontrol.frx":030A
         DataField       =   "FRecep"
         Height          =   315
         Index           =   0
         Left            =   6825
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   810
         Width           =   1260
         _ExtentX        =   2223
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   12
         Format          =   "dd/MM/yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label lbl 
         Caption         =   "&Fecha:"
         Height          =   240
         Index           =   2
         Left            =   6075
         TabIndex        =   13
         Top             =   840
         Width           =   855
      End
      Begin VB.Label lbl 
         Caption         =   "&Cuentas:"
         Height          =   240
         Index           =   1
         Left            =   180
         TabIndex        =   12
         Top             =   825
         Width           =   1215
      End
      Begin VB.Label lbl 
         Caption         =   "&Inmueble:"
         Height          =   240
         Index           =   0
         Left            =   180
         TabIndex        =   11
         Top             =   345
         Width           =   1215
      End
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid Grid 
      Height          =   4755
      Left            =   240
      TabIndex        =   0
      Tag             =   "400|1000|1200|5000|1500|1200|600|400"
      Top             =   1560
      Width           =   11715
      _ExtentX        =   20664
      _ExtentY        =   8387
      _Version        =   393216
      Cols            =   8
      FixedCols       =   0
      BackColorFixed  =   -2147483646
      ForeColorFixed  =   -2147483639
      BackColorBkg    =   -2147483636
      GridColor       =   -2147483633
      WordWrap        =   -1  'True
      AllowBigSelection=   0   'False
      FocusRect       =   2
      HighLight       =   2
      GridLinesFixed  =   1
      FormatString    =   " |^Cheque|^Fecha Cheque|Beneficiario|>Monto|^Firma 1|^Firma 2|"
      _NumberOfBands  =   1
      _Band(0).Cols   =   8
      _Band(0).GridLinesBand=   1
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   0
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
   Begin VB.Menu chequecontrol 
      Caption         =   "Cheque Control"
      Visible         =   0   'False
      Begin VB.Menu cargado 
         Caption         =   "Detalle Cargado"
      End
   End
End
Attribute VB_Name = "frmchequecontrol"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rstcontrol(2) As ADODB.Recordset
Attribute rstcontrol.VB_VarHelpID = -1

Private Sub cargado_Click()
'muestra el cargado del cheque
Dim rst As ADODB.Recordset
Dim sql As String, IDCheque As Integer

IDCheque = grid.TextMatrix(grid.RowSel, 1)

MsgBox IDCheque, vbInformation, App.ProductName
End Sub

Private Sub cmd_Click(Index As Integer)
Select Case Index
    Case 0
        Unload Me
        Set frmchequecontrol = Nothing
    Case 1
        Call firmar_cheques
        
End Select
End Sub

Private Sub dtc_Click(Index As Integer, Area As Integer)
If Area = 2 Then
    
    Select Case Index
    
        Case 0, 1 'codinm y nombre inm
            If Index = 0 Then dtc(1) = dtc(0).BoundText
            If Index = 1 Then dtc(1) = dtc(0).BoundText
            Call rtnLimpiar_Grid(grid)
            Call Inf_Inmueble
            Call Listar
            
            
        Case 2, 3 'banco y cuenta
            If Index = 2 Then dtc(3) = dtc(2).BoundText
            If Index = 3 Then dtc(2) = dtc(3).BoundText
            'MousePointer = vbHourglass
            Call rtnLimpiar_Grid(grid)
            Call Listar
            'MousePointer = vbDefault
    End Select
        
End If
End Sub


Private Sub Form_Load()
Set rstcontrol(0) = New ADODB.Recordset
Set rstcontrol(1) = New ADODB.Recordset
Set rstcontrol(2) = New ADODB.Recordset

img(0).Picture = LoadResPicture("Unchecked", vbResBitmap)
img(1).Picture = LoadResPicture("Checked", vbResBitmap)
    
sql = "SELECT * FROM Inmueble WHERE Inactivo=False ORDER BY CodInm"
rstcontrol(0).Open sql, cnnConexion, adOpenStatic, adLockReadOnly, adCmdText

sql = "SELECT * FROM Inmueble WHERE Inactivo=False ORDER BY Nombre"
rstcontrol(1).Open sql, cnnConexion, adOpenStatic, adLockReadOnly, adCmdText

Set dtc(0).RowSource = rstcontrol(0)
Set dtc(1).RowSource = rstcontrol(1)
Call centra_titulo(grid, True)
Call Listar
End Sub

Private Sub Inf_Inmueble()
'
Dim strOrigenDatos As String

With rstcontrol(0)
    '
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
    If !Caja = sysCodCaja Then
        strOrigenDatos = gcPath & "\" + sysCodInm + "\inm.mdb"
    Else
        strOrigenDatos = gcPath & !Ubica & "inm.mdb"
    End If
    '
    Call CUENTA_INMUEBLE(strOrigenDatos)
    '
End With
'
End Sub

Private Sub CUENTA_INMUEBLE(strOrigenDatos As String)
'variables locales
Dim strSQL As String
'
strSQL = "SELECT Bancos.*, Cuentas.* FROM Bancos INNER JOIN Cuentas ON Bancos.IDBanco = Cue" _
& "ntas.IDBanco WHERE Cuentas.Inactiva=False;"

If Dir(strOrigenDatos) <> "" Then
    If rstcontrol(2).State = 1 Then rstcontrol(2).Close
   rstcontrol(2).Open strSQL, cnnOLEDB + strOrigenDatos, adOpenStatic, adLockReadOnly, adCmdText
    '
    Set dtc(2).RowSource = rstcontrol(2)
    Set dtc(3).RowSource = rstcontrol(2)
'    If rstcontrol(2).RecordCount > 0 Then
'        dtc(2).Text = rstcontrol(2)("NombreBanco")
'        Call dtc_Click(2, 2)
'    End If
Else
    Set dtc(2).RowSource = Nothing
    Set dtc(3).RowSource = Nothing
End If

'
End Sub

Private Sub Listar()
Dim strSQL As String, strFiltro As String
Dim rst As ADODB.Recordset
Dim I As Long

If dtc(0) <> "" Then
    
    If dtc(2) <> "" Then
        strFiltro = " AND C.Cuenta='" & dtc(3) & "'"
    Else
        If Not (rstcontrol(2).BOF And rstcontrol(2).EOF) Then
            rstcontrol(2).MoveFirst
            Do
                If strFiltro <> "" Then strFiltro = strFiltro & " OR "
                strFiltro = strFiltro & "C.Cuenta='" & rstcontrol(2)("NumCuenta") & "'"
                rstcontrol(2).MoveNext
            Loop Until rstcontrol(2).EOF
            strFiltro = "AND (" & strFiltro & ") "
        End If
        
    End If
End If
    If IsDate(MskFecha(0)) Then
        strFiltro = strFiltro & " AND C.FechaCheque =#" & Format(MskFecha(0), "mm/dd/yy") & "#"
    End If
strSQL = "SELECT C.IDCheque, C.FechaCheque, C.Beneficiario, C.Banco," & _
         "C.Cuenta, Sum(CD.Monto) as Total, C.Firma1, C.Firma2 " & _
         "FROM Cheque as C INNER JOIN ChequeDetalle as CD " & _
         "ON C.Clave = CD.Clave " & _
         "WHERE (C.Firma1=false or C.Firma2=false) " & strFiltro & _
         "GROUP BY C.IDCheque, C.FechaCheque, C.Beneficiario, " & _
         "C.Banco, C.Cuenta, C.Firma1, C.Firma2 "

Set rst = New ADODB.Recordset

rst.Open strSQL, cnnConexion, adOpenKeyset, adLockOptimistic, adCmdText

If Not (rst.EOF And rst.BOF) Then
    I = 1
    Call rtnLimpiar_Grid(grid)
    grid.Rows = rst.RecordCount + 1
    Do
        grid.TextMatrix(I, 0) = I
        grid.TextMatrix(I, 1) = Format(rst("IDCheque"), "000000")
        grid.Row = I
        grid.Col = 1
        grid.CellFontUnderline = True
        grid.CellForeColor = &HFF0000
'        Grid.CellForeColor = &H80000012
'        Grid.CellFontUnderline = False
        grid.TextMatrix(I, 2) = rst("FechaCheque")
        grid.TextMatrix(I, 3) = rst("Beneficiario")
        grid.TextMatrix(I, 4) = Format(rst("Total"), "#,##0.00")
        grid.TextMatrix(I, 7) = I
        grid.Col = 5
        grid.Row = I
        Set grid.CellPicture = img(IIf(rst("Firma1"), 1, 0))
        grid.CellPictureAlignment = flexAlignCenterCenter
        grid.Col = 6
        Set grid.CellPicture = img(IIf(rst("Firma2"), 1, 0))
        grid.CellPictureAlignment = flexAlignCenterCenter
        rst.MoveNext
        If I Mod 2 = 0 Then Marcar_Linea grid, &H80000018
        I = I + 1
        DoEvents
        
        If I > 1000 Then
            MsgBox "Demasiados registros, no se pueden mostrar todos", vbInformation, Me.Caption
            rst.Close
            Set rst = Nothing
            grid.Rows = 1001
            Exit Sub
        End If
        
    Loop Until rst.EOF
    
End If
rst.Close
Set rst = Nothing

End Sub


Private Sub Form_Unload(Cancel As Integer)
Set rstcontrol(0) = Nothing
Set rstcontrol(1) = Nothing
Set rstcontrol(2) = Nothing
End Sub

Private Sub grid_Click()
With grid
    If .RowSel >= 1 Then
        If .ColSel = 5 Or .ColSel = 6 Then
            .Col = .ColSel
            Set .CellPicture = IIf(.CellPicture = img(0), img(1), img(0))
            .CellPictureAlignment = flexAlignCenterCenter
        End If
    End If
End With
End Sub

Private Sub firmar_cheques()
Dim strSQL As String, msg As String
Dim firma1 As Integer, firma2 As Integer
Dim I As Long
Dim acum As Long

MousePointer = vbHourglass
cmd(1).Enabled = False
cmd(0).Enabled = False
With grid
    For I = 1 To grid.Rows - 1
        If (IsNumeric(.TextMatrix(I, 1))) Then
            .Col = 5
            .Row = I
            firma1 = IIf(.CellPicture = img(1), -1, 0)
            .Col = 6
            firma2 = IIf(.CellPicture = img(1), -1, 0)
            
            strSQL = "update cheque set Firma1=" & firma1 & _
                    ", firma2 = " & firma2 & _
                    " WHERE IDCheque = " & .TextMatrix(I, 1)
            cnnConexion.Execute strSQL, r
            Call rtnBitacora("Cheque " & .TextMatrix(I, 1) & "Firma1: " & _
                            firma1 & ", Firma2: " & firma2)
            acum = acum + r
            DoEvents
        End If
    Next
    If acum > 0 Then
        msg = "Se han actualizado " & acum & " registros."
    Else
        msg = "Ningún registro actualizado."
    End If
    MsgBox msg, vbInformation, Me.Caption
End With
cmd(1).Enabled = True
cmd(0).Enabled = True
MousePointer = vbDefault
End Sub

Private Sub Grid_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 Then
    PopupMenu chequecontrol
End If
End Sub


Private Sub MskFecha_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii = 13 Then
    Call rtnLimpiar_Grid(grid)
    Call Listar
End If
End Sub
