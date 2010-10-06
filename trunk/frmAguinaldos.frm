VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmAguinaldos 
   Caption         =   "Aguinaldos"
   ClientHeight    =   3420
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10740
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   3420
   ScaleWidth      =   10740
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmd 
      Caption         =   "&Salir"
      Height          =   1275
      Index           =   0
      Left            =   10200
      TabIndex        =   17
      Top             =   7680
      Width           =   1215
   End
   Begin TabDlg.SSTab tab 
      Height          =   9285
      Left            =   360
      TabIndex        =   0
      Top             =   480
      Width           =   11670
      _ExtentX        =   20585
      _ExtentY        =   16378
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Asignación"
      TabPicture(0)   =   "frmAguinaldos.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "fra(2)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Consulta"
      TabPicture(1)   =   "frmAguinaldos.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "fra(0)"
      Tab(1).ControlCount=   1
      Begin VB.Frame fra 
         Height          =   8295
         Index           =   2
         Left            =   360
         TabIndex        =   8
         Top             =   600
         Width           =   10935
         Begin VB.Frame fra 
            Caption         =   "Procesar Nomina: "
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1680
            Index           =   3
            Left            =   540
            TabIndex        =   29
            Top             =   6330
            Width           =   2835
            Begin VB.TextBox txt 
               Height          =   315
               Left            =   1230
               TabIndex        =   31
               Top             =   330
               Width           =   1350
            End
            Begin VB.CommandButton cmd 
               Cancel          =   -1  'True
               Caption         =   "&Cerrar Aguinaldos"
               Enabled         =   0   'False
               Height          =   675
               Index           =   4
               Left            =   255
               TabIndex        =   30
               Top             =   795
               Width           =   2370
            End
            Begin VB.Label lbl 
               Caption         =   "Contraseña:"
               Height          =   255
               Index           =   10
               Left            =   285
               TabIndex        =   32
               Top             =   330
               Width           =   855
            End
         End
         Begin VB.CommandButton cmd 
            Caption         =   "&Facturar"
            Height          =   1275
            Index           =   2
            Left            =   6480
            TabIndex        =   16
            Top             =   6275
            Width           =   1215
         End
         Begin VB.CommandButton cmd 
            Caption         =   "&Guardar"
            Height          =   1275
            Index           =   1
            Left            =   7965
            TabIndex        =   15
            Top             =   6275
            Width           =   1215
         End
         Begin MSDataListLib.DataCombo dtc 
            Height          =   315
            Index           =   0
            Left            =   1335
            TabIndex        =   10
            Top             =   540
            Width           =   1170
            _ExtentX        =   2064
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo dtc 
            Height          =   315
            Index           =   1
            Left            =   2520
            TabIndex        =   11
            Top             =   525
            Width           =   6900
            _ExtentX        =   12171
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
         End
         Begin MSFlexGridLib.MSFlexGrid Flex 
            Height          =   3615
            Left            =   540
            TabIndex        =   14
            Tag             =   "Caja"
            Top             =   2295
            Width           =   7860
            _ExtentX        =   13864
            _ExtentY        =   6376
            _Version        =   393216
            Cols            =   8
            FixedCols       =   0
            RowHeightMin    =   285
            BackColorFixed  =   -2147483646
            ForeColorFixed  =   -2147483639
            BackColorSel    =   65280
            ForeColorSel    =   0
            GridColor       =   -2147483645
            WordWrap        =   -1  'True
            AllowBigSelection=   0   'False
            FocusRect       =   2
            HighLight       =   2
            ScrollBars      =   2
            AllowUserResizing=   1
            BorderStyle     =   0
            MousePointer    =   99
            FormatString    =   "CodEmp|Nombre y Apellido|Cargo|Sueldo|BonoNoc|Días s/Ley|Días Adic.|Bs. Aguinaldos"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MouseIcon       =   "frmAguinaldos.frx":0038
         End
         Begin MSMask.MaskEdBox mskRemesa 
            Height          =   315
            Left            =   1335
            TabIndex        =   19
            Top             =   1530
            Width           =   1170
            _ExtentX        =   2064
            _ExtentY        =   556
            _Version        =   393216
            MaxLength       =   7
            Format          =   "mm-yyyy"
            Mask            =   "##-####"
            PromptChar      =   "_"
         End
         Begin MSDataListLib.DataCombo dtc 
            Height          =   315
            Index           =   2
            Left            =   3720
            TabIndex        =   20
            Top             =   1050
            Width           =   1170
            _ExtentX        =   2064
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo dtc 
            Height          =   315
            Index           =   3
            Left            =   4920
            TabIndex        =   21
            Top             =   1050
            Width           =   4500
            _ExtentX        =   7938
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo dtc 
            Height          =   315
            Index           =   4
            Left            =   3720
            TabIndex        =   22
            Top             =   1530
            Width           =   1170
            _ExtentX        =   2064
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo dtc 
            Height          =   315
            Index           =   5
            Left            =   4920
            TabIndex        =   23
            Top             =   1530
            Width           =   4500
            _ExtentX        =   7938
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0,00"
            Height          =   285
            Index           =   9
            Left            =   9540
            TabIndex        =   27
            Top             =   1530
            Width           =   1245
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0,00"
            Height          =   300
            Index           =   8
            Left            =   9540
            TabIndex        =   26
            Top             =   1050
            Width           =   1245
         End
         Begin VB.Label lbl 
            Caption         =   "Cod.Gasto(2):"
            Height          =   255
            Index           =   7
            Left            =   2640
            TabIndex        =   25
            Top             =   1590
            Width           =   1215
         End
         Begin VB.Label lbl 
            Caption         =   "Cod.Gasto(1):"
            Height          =   255
            Index           =   6
            Left            =   2640
            TabIndex        =   24
            Top             =   1110
            Width           =   1215
         End
         Begin VB.Label lbl 
            Caption         =   "Cargado:"
            Height          =   255
            Index           =   5
            Left            =   495
            TabIndex        =   18
            Top             =   1590
            Width           =   690
         End
         Begin VB.Label lbl 
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Index           =   4
            Left            =   1335
            TabIndex        =   13
            Top             =   1050
            Width           =   1170
         End
         Begin VB.Label lbl 
            Caption         =   "Año:"
            Height          =   255
            Index           =   3
            Left            =   495
            TabIndex        =   12
            Top             =   1110
            Width           =   1215
         End
         Begin VB.Label lbl 
            Caption         =   "Inmueble:"
            Height          =   255
            Index           =   2
            Left            =   495
            TabIndex        =   9
            Top             =   555
            Width           =   1215
         End
      End
      Begin VB.Frame fra 
         Height          =   8295
         Index           =   0
         Left            =   -74640
         TabIndex        =   1
         Top             =   600
         Width           =   10935
         Begin VB.CommandButton cmd 
            Caption         =   "&Imprimir"
            Height          =   1275
            Index           =   3
            Left            =   7965
            TabIndex        =   28
            Top             =   6805
            Width           =   1215
         End
         Begin VB.Frame fra 
            Caption         =   "Filtros:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1935
            Index           =   1
            Left            =   240
            TabIndex        =   2
            Top             =   6120
            Width           =   5745
            Begin VB.ComboBox cmb 
               Height          =   315
               Index           =   1
               Left            =   1260
               Sorted          =   -1  'True
               Style           =   2  'Dropdown List
               TabIndex        =   7
               Top             =   765
               Width           =   1500
            End
            Begin VB.ComboBox cmb 
               Height          =   315
               Index           =   0
               Left            =   1275
               Style           =   2  'Dropdown List
               TabIndex        =   5
               Top             =   330
               Width           =   4140
            End
            Begin VB.Label lbl 
               Caption         =   "A partir de:"
               Height          =   255
               Index           =   1
               Left            =   240
               TabIndex        =   6
               Top             =   810
               Width           =   885
            End
            Begin VB.Label lbl 
               Caption         =   "Inmueble:"
               Height          =   255
               Index           =   0
               Left            =   240
               TabIndex        =   3
               Top             =   360
               Width           =   1215
            End
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid Grid 
            Height          =   5400
            Left            =   240
            TabIndex        =   4
            Top             =   360
            Width           =   10440
            _ExtentX        =   18415
            _ExtentY        =   9525
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
            FormatString    =   "Cod.Emp|Nombre y Apellido|Cargo|Sueldo Mensual|Fecha Ingreso|Dias Agui|Aguinal s/Ley|d/apro. 2004|d/apro. 2005"
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
End
Attribute VB_Name = "frmAguinaldos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim bEntrada As Boolean

Private Sub cmb_Click(Index As Integer)
Call mostrar_informacion(obtener_filtro(), CInt(cmb(1).Text))
End Sub

Private Sub cmd_Click(Index As Integer)
Select Case Index
    Case 0  'cerrar
        Unload Me
        Set frmAguinaldos = Nothing
    
    Case 1  'guardar
        If guardar_aguinaldos Then
           MsgBox "Los aguinaldos de " & dtc(1) & vbCrLf & _
           "han sido guardados con éxito", vbInformation, App.ProductName
        End If
    Case 2  'facturar
        If facturar_aguinaldos Then
            
            MsgBox "Los aguinaldos de " & dtc(1) & vbCrLf & _
            "han sido facturados con éxito", vbInformation, App.ProductName
            
         End If
    Case 3  'imprimir
      
    Case 4  'cerrar aguinaldos
        Me.tab.Enabled = False
        If cerrar_aguinaldos Then
            MsgBox "Aguinaldos " & lbl(4) & " cerrados con éxito."
            Call rtnBitacora("Aguinaldos " & lbl(4) & " cerrados con éxito.")
        End If
        Me.tab.Enabled = True
        
End Select

End Sub

Private Sub dtc_Click(Index As Integer, Area As Integer)
'
If Area = 2 Then
    Select Case Index
        
        Case 0, 1
            dtc(IIf(Index = 0, 1, 0)) = dtc(Index).BoundText
            busca_empleados (dtc(0).Text)
        
        Case 2, 3
            dtc(IIf(Index = 2, 3, 2)) = dtc(Index).BoundText
            
        Case 4, 5
            dtc(IIf(Index = 4, 5, 5)) = dtc(Index).BoundText
            
    End Select
End If
'
End Sub

Private Sub dtc_KeyPress(Index As Integer, KeyAscii As Integer)
'buscamos el inmueble por nombre o código
KeyAscii = Asc(UCase(Chr(KeyAscii)))

If Index = 0 Or Index = 2 Or Index = 4 Then Call Validacion(KeyAscii, "0123456789")

If KeyAscii = 13 Then
    Select Case Index
        Case 0, 1
            Dim StrCampo As String
            StrCampo = IIf(Index = 0, "CodInm", "Nombre")
            dtc(IIf(Index = 0, 1, 0)) = ""
            Call selecciona_inmueble(StrCampo, dtc(Index).Text)
            busca_empleados (dtc(0).Text)
        Case 2, 3
            dtc(IIf(Index = 2, 3, 2)) = dtc(Index).Text
            Call selecciona_gasto(IIf(Index = 2, "CodGasto", "Titulo"), dtc(Index), dtc(0).Text, 2, 3)
            
        Case 4, 5
            dtc(IIf(Index = 4, 5, 4)) = dtc(Index).Text
            Call selecciona_gasto(IIf(Index = 4, "CodGasto", "Titulo"), dtc(Index), dtc(0).Text, 4, 5)
            
    End Select
End If

End Sub

Private Sub dtc_LostFocus(Index As Integer)
If Index = 0 Or Index = 1 Then
    dtc(IIf(Index = 0, 1, 0)) = dtc(Index).BoundText
    busca_empleados (dtc(0).Text)
End If
End Sub

Private Sub Flex_EnterCell()
If Flex.Col = 6 Then bEntrada = True
End Sub

Private Sub flex_KeyDown(KeyCode As Integer, Shift As Integer)
Call mantenimiento_celda(Flex, 6, KeyCode, Shift)
End Sub

Private Sub Flex_KeyPress(KeyAscii As Integer)
If Flex.ColSel = 6 Then
    If KeyAscii = Asc(".") Then KeyAscii = Asc(",")
    Call Validacion(KeyAscii, "0123456789,")
    If InStr(Flex.Text, ",") And KeyAscii = Asc(",") Then KeyAscii = 0
    If KeyAscii = 8 Then Flex.Text = Left(Flex.Text, Len(Flex.Text) - 1)
    
    If KeyAscii > 26 Then Flex.Text = IIf(bEntrada, "", Flex.Text) & Chr(KeyAscii)
    
    If KeyAscii = 13 Then
        If Flex.Text <> "" Then
            If Not IsNumeric(Flex.Text) Then
                MsgBox "Monto no válido [" & Flex.Text & "]", vbExclamation, _
                App.ProductName
                Exit Sub
            End If
        End If
        Flex.Col = 6
    End If
    bEntrada = False
End If
End Sub

Private Sub Flex_LeaveCell()
If Flex.ColSel = 6 And Flex.RowSel > 0 And Flex.TextMatrix(Flex.RowSel, 5) <> "" Then
    Dim iDias As Double
    Dim cAguinaldo As Currency
    Dim I As Integer
    
    If Len(Flex.TextMatrix(Flex.RowSel, 6)) > 1 Then
        Flex.Text = CCur(Flex.TextMatrix(Flex.RowSel, 6))
    End If
    If Flex.TextMatrix(Flex.RowSel, 6) = "" Then Flex.TextMatrix(Flex.RowSel, 6) = "0"
    iDias = CDbl(Flex.TextMatrix(Flex.RowSel, 5)) + _
            CDbl(Flex.TextMatrix(Flex.RowSel, 6))
    cAguinaldo = iDias * (CCur(Flex.TextMatrix(Flex.RowSel, 3)) / 30)
    Flex.TextMatrix(Flex.RowSel, 7) = Format(cAguinaldo, "#,##0.00")
    cAguinaldo = 0
    For I = 1 To Flex.Rows - 1
        cAguinaldo = cAguinaldo + (CCur(Flex.TextMatrix(I, 6)) * CCur(Flex.TextMatrix(I, 3)) / 30)
    Next
    lbl(9).Caption = Format(cAguinaldo, "#,##0.00")
End If
End Sub

Private Sub Form_Load()
Dim strSQL As String
Dim I As Integer
Dim rst As ADODB.Recordset

strSQL = "SELECT TOP 2 Right(nom_inf.IDNomina,4) as ano " & _
         "FROM nom_inf " & _
         "WHERE (((Left([nom_inf].[IDNomina], 1)) = 3)) " & _
         "ORDER BY Right(nom_inf.IDNomina,4) DESC"
        
Call llenar_combo(cmb(1), listar_valores(strSQL, False), 0)

lbl(4) = cnnConexion.Execute(strSQL).Fields(0).Value + 1

strSQL = "SELECT CodInm, Nombre FROM Inmueble Where Inactivo =false"
Set rst = cnnConexion.Execute(strSQL)

dtc(0).ListField = "CodInm"
dtc(0).BoundColumn = "Nombre"
Set dtc(0).RowSource = rst

dtc(1).ListField = "Nombre"
dtc(1).BoundColumn = "CodInm"
Set dtc(1).RowSource = rst


strSQL = "SELECT CodInm & ' ' & Nombre FROM Inmueble ORDER BY CodInm"
Call llenar_combo(cmb(0), listar_valores(strSQL, True))
Call configurar_grid

End Sub

Private Sub Form_Resize()
Dim ancho()
Dim I As Integer

ancho = Array(600, 2500, 1500, 1000, 900, 600, 750, 750, 750, 750, 750, 750)
'centrar las fichas
Me.tab.Left = (Me.ScaleWidth - Me.tab.Width) / 2
Me.tab.Top = (Me.ScaleHeight - Me.tab.Height) / 2
'
With Grid
    .RowHeight(0) = 450
    .TextArray(0) = "Cod." & vbCrLf & "Emp"
    .TextArray(1) = "Nombre y" & vbCrLf & "Apellido"
    .TextArray(2) = "Cargo"
    .TextArray(3) = "Sueldo" & vbCrLf & "Mensual"
    .FillStyle = flexFillRepeat
    .Row = 0
    .RowSel = 1
    .Col = 0
    .ColSel = .Cols - 1
    .CellAlignment = flexAlignCenterCenter
    .FillStyle = flexFillSingle
    .Row = 1
    .ColAlignment(5) = flexAlignCenterCenter
    For I = 0 To .Cols - 1
        .ColWidth(I) = ancho(I)
    Next
End With
Call configurar_botones
End Sub

Private Sub mostrar_informacion(strFiltro As String, intYear As Integer)
Dim rstlocal As ADODB.Recordset
Dim Inm As String, sNombre As String
Dim E As Integer, I As Integer
Dim subTotal As Currency, SueldoNeto As Currency

Call habilitar_boton(False)

Set rstlocal = ejecutar_procedure("agui_year", 1, intYear)
rstlocal.Filter = strFiltro
rstlocal.Sort = "CodInm,CodEmp"
Call rtnLimpiar_Grid(Grid)
If Not (rstlocal.EOF And rstlocal.BOF) Then
    rstlocal.MoveFirst
    
    Grid.Cols = rstlocal.Fields.count - 1
    For E = 8 To rstlocal.Fields.count - 1
        Grid.TextArray(E - 1) = "D/Aprob." & vbCrLf & rstlocal.Fields(E).Name
    Next
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
            Grid.MergeRow(I) = False
            sNombre = rstlocal("Empleado")
            If sNombre = "" Then sNombre = rstlocal("Empleado")
            Grid.RowHeight(I) = IIf(Screen.Width / Screen.TwipsPerPixelX >= 1024, 250, 215)
            Grid.TextMatrix(I, 0) = rstlocal("CodEmp")
            Grid.TextMatrix(I, 1) = sNombre
            Grid.TextMatrix(I, 2) = rstlocal("NombreCargo")
            SueldoNeto = (rstlocal("Sueldo") * rstlocal("BonoNoc") / 100) + rstlocal("Sueldo")
            Grid.TextMatrix(I, 3) = Format(SueldoNeto, "#,##0.00 ")
            Grid.TextMatrix(I, 4) = rstlocal("Fingreso")
            Grid.TextMatrix(I, 5) = IIf(DateDiff("m", rstlocal("Fingreso"), "30/11/" & Year(Date)) >= 12, 15, 15 / 12 * DateDiff("m", rstlocal("FIngreso"), "30/11/" & Year(Date)))
            Grid.TextMatrix(I, 6) = Format(SueldoNeto / 30 * Grid.TextMatrix(I, 5), "#,##0.00")
            subTotal = subTotal + Grid.TextMatrix(I, 6)
            For E = 8 To rstlocal.Fields.count - 1
                Grid.ColAlignment(E - 1) = flexAlignCenterCenter
                Grid.TextMatrix(I, E - 1) = IIf(IsNull(rstlocal.Fields(E)), "--", rstlocal.Fields(E))
            Next
            Grid.Col = 0
            Grid.Row = I
            Grid.ColSel = Grid.Cols - 1
            Grid.RowSel = I
            Grid.FillStyle = flexFillRepeat
            Grid.CellFontName = "Arial Narrow"
            Grid.CellFontSize = IIf(Screen.Width / Screen.TwipsPerPixelX >= 1024, 8, 8)
            Grid.FillStyle = flexFillSingle
            Inm = rstlocal("CodInm")
            rstlocal.MoveNext
    Loop Until rstlocal.EOF
    I = I + 1
    Grid.AddItem ""
    Grid.TextMatrix(I, 6) = Format(subTotal, "#,##0.00")
    Grid.Row = I
    Grid.Col = 6
    Grid.CellFontBold = True
    Grid.RowHeight(I) = 200
    '
End If
Call habilitar_boton(True)
Set rstlocal = Nothing
End Sub

Private Sub llenar_combo(cmb As ComboBox, Items As String, Optional ElementoSeleccionado As Integer)
Dim I As Integer
Dim aElementos() As String
aElementos = Split(Items, ",")
For I = LBound(aElementos) To UBound(aElementos)
    cmb.AddItem aElementos(I)
Next
If Not IsEmpty(ElementoSeleccionado) Then cmb.ListIndex = ElementoSeleccionado
End Sub

Public Function obtener_filtro() As String
Dim strFiltro As String
Dim aInmueble() As String
If cmb(0).Text <> "" Then
    aInmueble = Split(cmb(0).Text, " ")
    strFiltro = "CodInm='" & aInmueble(0) & "'"
End If
obtener_filtro = strFiltro
End Function

Private Sub configurar_grid()
Dim ancho As Long
Dim I As Integer
ancho = (fra(2).Width - (Flex.Left * 2))
Flex.Width = ancho
Flex.ColWidth(0) = 0.1 * ancho  'codigo empleado
Flex.ColWidth(1) = 0.42 * ancho 'nombre y apellido (Cargo)
Flex.ColWidth(2) = 0.11 * ancho  'Fecha Ingreso
Flex.ColWidth(3) = 0.09 * ancho 'sueldo mensual
Flex.ColWidth(4) = 0            'bono nocturno
Flex.ColWidth(5) = 0.08 * ancho 'dias s/ley
Flex.ColWidth(6) = 0.08 * ancho 'dias adicionales
Flex.ColWidth(7) = 0.09 * ancho 'Bs. aguinaldos

Flex.TextMatrix(0, 0) = "Código" & vbCrLf & "Empleado"
Flex.TextMatrix(0, 2) = "Fecha de" & vbCrLf & "Ingreso"
Flex.TextMatrix(0, 5) = "Días" & vbCrLf & "s/Ley(1)"
Flex.TextMatrix(0, 6) = "Días" & vbCrLf & "Adic.(2)"
Flex.Row = 0
For I = 0 To Flex.Cols - 1
    Flex.Col = I
    Flex.CellAlignment = flexAlignCenterCenter
Next
Flex.ColAlignment(0) = flexAlignCenterCenter
Flex.ColAlignment(2) = flexAlignCenterCenter
Flex.RowHeight(0) = 1.2 * TextHeight(Flex.TextMatrix(0, 0))
End Sub

Private Sub busca_empleados(Codigo_Inmueble As String)
Dim rst As ADODB.Recordset
Dim cAguiSL As Currency, cAguiAD As Currency

Call habilitar_boton(False)

Set rst = ModGeneral.ejecutar_procedure("agui_temp", Codigo_Inmueble)

Call rtnLimpiar_Grid(Flex)

If Not (rst.EOF And rst.BOF) Then
    Dim I As Integer
    Flex.Rows = rst.RecordCount + 1
    Do
        'DoEvents
        Flex.RowHeight(rst.AbsolutePosition) = 250
        For I = 0 To rst.Fields.count - 1
            
            Flex.TextMatrix(rst.AbsolutePosition, I) = IIf(IsNull(rst(I)), 0, rst(I))
            If rst(I).Name = "Sueldo" Then
                Flex.TextMatrix(rst.AbsolutePosition, I) = Format(rst("Sueldo"), "#,##0.00")
            End If
        Next
        cAguiSL = cAguiSL + (rst("Sueldo") / 30 * rst("DiasSley"))
        cAguiAD = cAguiAD + (rst("Sueldo") / 30 * rst("DiasAgui"))
        Flex.TextMatrix(rst.AbsolutePosition, 7) = Format(rst("Sueldo") / 30 * (rst("DiasSley") + rst("DiasAgui")), "#,##0.00")
        rst.MoveNext
    Loop Until rst.EOF
    Flex.SetFocus
    Flex.Row = 1
    Flex.Col = 6
    '
    Call setCodGasto(Codigo_Inmueble)
End If
lbl(8).Caption = Format(cAguiSL, "#,##0.00")
lbl(9).Caption = Format(cAguiAD, "#,##0.00")
Call habilitar_boton(True)
cmd(2).Enabled = aguinaldo_no_cerrado(CLng("312" & lbl(4)), Codigo_Inmueble, lbl(4))
cmd(1).Enabled = cmd(2).Enabled

End Sub

Private Sub configurar_botones()
Dim ancho As Long
Dim I As Integer

cmd(0).Left = Me.tab.Left + fra(2).Left + Flex.Left + Flex.Width - cmd(0).Width
cmd(0).Top = Me.tab.Top + fra(2).Top + fra(2).Height - cmd(0).Height - 200
For I = 1 To 2
    cmd(I).Top = cmd(0).Top - fra(2).Top - Me.tab.Top - 5
    cmd(I).Left = Flex.Left + Flex.Width - (cmd(I).Width * I) - cmd(0).Width
Next

End Sub

Private Sub selecciona_inmueble(campo As String, valor As String)
Dim rst As ADODB.Recordset
Dim sql As String

sql = "SELECT * FROM Inmueble WHERE " & campo & " Like '%" & valor & "%'"

Set rst = cnnConexion.Execute(sql)
If Not (rst.EOF And rst.BOF) Then
    dtc(0) = rst("CodInm")
    dtc(1) = rst("Nombre")
    
End If
rst.Close
Set rst = Nothing

End Sub

Private Sub selecciona_gasto(cammpo As String, valor As String, _
                            CodInm As String, DC1 As Integer, DC2 As Integer)
Dim rst As ADODB.Recordset
Dim sql As String

sql = "SELECT * FROM TGastos in '" & gcPath & "\" & CodInm & "\inm.mdb' " & _
      "WHERE " & cammpo & " LIKE '%" & valor & "%'"

Set rst = cnnConexion.Execute(sql)
If Not (rst.EOF And rst.BOF) Then
    dtc(DC1) = rst("CodGasto")
    dtc(DC2) = rst("Titulo")
End If
rst.Close
Set rst = Nothing
End Sub
Private Function guardar_aguinaldos() As Boolean
Dim I As Integer, N As Integer
Dim sql As String

On Error GoTo salir

guardar_aguinaldos = False

'validamos los datos mínimos necesarios para actualizar aguinaldos
If Not dtc(0).MatchedWithList Then sql = "- Código de Inmueble." + vbCrLf
If Not dtc(0).MatchedWithList Then sql = sql + "- Nombre Inmueble." + vbCrLf
If Flex.TextMatrix(1, 0) = "" Then sql = sql + "- No existen empleados registrados."

If sql <> "" Then
    sql = "No se pudo completar esta operación." + vbCrLf + "Faltan datos:" + vbCrLf + vbCrLf + sql
    MsgBox sql, vbExclamation, App.ProductName
    Exit Function
End If
cnnConexion.BeginTrans
With Flex
    Call rtnBitacora("Guardando aguinaldos inm " & dtc(0) & "...")

    For I = 1 To .Rows - 1
       sql = "UPDATE Emp SET DiasAgui='" & .TextMatrix(I, 6) & _
            "' WHERE CodEmp=" & .TextMatrix(I, 0)
       cnnConexion.Execute sql, N
       Call rtnBitacora(.TextMatrix(I, 0) & " Dias:" & _
                        .TextMatrix(I, 6) & "Reg.: " & N)
    Next
    
End With
salir:
If Err <> 0 Then
    cnnConexion.RollbackTrans
    MsgBox Err.Description, vbCritical, _
    "Ocurrió un error durante el proceso"
Else
    cnnConexion.CommitTrans
    guardar_aguinaldos = True
End If
End Function

Private Function facturar_aguinaldos() As Boolean
Dim msg As String

facturar_aguinaldos = False
mskRemesa.PromptInclude = False

If mskRemesa.Text = "" Then
    msg = "Introduzca el período de facturación"
    MsgBox msg, vbCritical, App.ProductName
    Exit Function
End If
mskRemesa.PromptInclude = True
If Not IsDate("01/" & mskRemesa.Text) Then
    msg = "Introdujo un período de facturación inválido"
    mskRemesa.SetFocus
    mskRemesa.SelStart = 0
    mskRemesa.SelLength = Len(mskRemesa.Text)
    MsgBox msg, vbCritical, App.ProductName
    Exit Function
End If
If Not dtc(0).MatchedWithList Then
    msg = "El código de inmueble no corresponde a un elemento de la lista."
    dtc(0).SetFocus
    dtc(0).SelStart = 0
    dtc(0).SelLength = Len(dtc(0).Text)
    MsgBox msg, vbCritical, "Facturar Aguinaldos"
    Exit Function
End If
If Not dtc(1).MatchedWithList Then
    MsgBox "El nombre del inmueble no corresponde a un elemento de la lista."
    dtc(1).SetFocus
    dtc(1).SelStart = 0
    dtc(1).SelLength = Len(dtc(1).Text)
    MsgBox msg, vbCritical, "Facturar Aguinaldos"
    Exit Function
End If

msg = "Esta seguro de facturar los aguinaldos de " & vbCrLf & _
dtc(1).Text & "?"
If Respuesta(msg) = False Then Exit Function
Dim strSQL As String
Dim rst As ADODB.Recordset

'validamos que el periodo de facturación no se haya facturado
strSQL = "SELECT MAX(Periodo) FROM Factura in '" & gcPath & "\" & dtc(0).Text & "\inm.mdb' WHERE Fact Not Like 'CHD%';"
Set rst = cnnConexion.Execute(strSQL)
If Not IsNull(rst.Fields(0)) And rst.Fields(0) >= CDate("01/" & mskRemesa.Text) Then
    msg = "No se puede procesar. Período ya facturado." & vbCrLf & _
          "Ingrese un nuevo período e inténtelo nuevamente."
    MsgBox msg, vbCritical, "Proceso de Facturación"
    Exit Function
End If
'si llegamos a este punto comenzamos el proceso de facturacion
Dim strDoc As String, strIDNom As String, strBdInm As String
Dim cAgui(1) As Currency, cSueldoNeto As Currency, cBonoNoc As Currency
Dim I As Integer, j As Integer
Dim NumEmpleado As Long

strDoc = "AGU" & Trim(lbl(4))
strBdInm = gcPath & "\" & dtc(0).Text & "\inm.mdb"
strIDNom = "312" & lbl(4)
        
With Flex
    For I = 1 To .Rows - 1
        For j = 0 To 1
            cAgui(j) = cAgui(j) + _
                    (CCur(.TextMatrix(I, 3)) / 30 * CDbl(.TextMatrix(I, j + 5)))
        Next
        
        cBonoNoc = 1 + (CCur(.TextMatrix(I, 4)) / 100)
        cSueldoNeto = CLng(CCur(.TextMatrix(I, 3)) / cBonoNoc * 100) / 100
        cBonoNoc = cBonoNoc - 1
        cBonoNoc = CLng((cSueldoNeto * cBonoNoc) / 30 * CCur(.TextMatrix(I, 5)) * 100) / 100
        
        NumEmpleado = .TextMatrix(I, 0)
        
        
        'guardamos la información en Nom_Detalle
        If Not agui_registrado(CLng(strIDNom), NumEmpleado) Then
            strSQL = "INSERT INTO Nom_Detalle(IDNom,CodEmp,Sueldo,Dias_Trab," & _
                     "Dias_libres,Porc_BonoNoc,Bono_Noc) VALUES (" & strIDNom & _
                     "," & NumEmpleado & ",'" & cSueldoNeto & "','" & _
                     .TextMatrix(I, 5) & "','" & .TextMatrix(I, 6) & "','" & _
                     .TextMatrix(I, 4) & "','" & cBonoNoc & "')"
            cnnConexion.Execute strSQL
        
        Else
            strSQL = "UPDATE Nom_Detalle SET Dias_libres='" & .TextMatrix(I, 6) & _
                     "' WHERE IDNom=" & strIDNom & " AND CodEmp=" & NumEmpleado
            cnnConexion.Execute strSQL
        
        End If
    Next
    'guardamos la información de la facturación
    For j = 0 To 1
        If gasto_no_facturado(strDoc, dtc(IIf(j = 0, 2, 4)).Text, strBdInm) Then
            'eliminamos el cargado
            strSQL = "DELETE FROM AsignaGasto in '" & strBdInm & "' WHERE NDoc='" & _
            strDoc & "' AND CodGasto='" & dtc(IIf(j = 0, 2, 4)).Text & "'"
            
            cnnConexion.Execute strSQL
            If cAgui(j) > 0 Then
                strSQL = "INSERT INTO AsignaGasto(NDoc,CodGasto,Cargado," & _
                         "Descripcion,Fijo,Comun,alicuota,monto,usuario," & _
                         "fecha,hora) in '" & strBdInm & "' " & _
                         "VALUES('" & strDoc & "','" & dtc(IIf(j = 0, 2, 4)).Text & _
                         "','01/" & mskRemesa & "','" & dtc(IIf(j = 0, 3, 5)).Text & _
                         "',0,-1,-1,'" & cAgui(j) & "','" & _
                         gcUsuario & "',Date(),Time())"
                         
                cnnConexion.Execute strSQL
            End If
        End If
    Next
End With
facturar_aguinaldos = True
End Function

Private Function gasto_no_facturado(NDoc As String, CodigoGasto As String, _
ubicaBD As String)
Dim rst As ADODB.Recordset
Dim sql As String

sql = "SELECT NDoc FROM AsignaGasto in '" & ubicaBD & "' WHERE Ndoc='" & _
      NDoc & "' AND CodGasto='" & CodigoGasto & "' AND Cargado <= " & _
      "(SELECT MAX(Periodo) FROM Factura in '" & ubicaBD & "' WHERE Fact Not Like 'CHD%')"

Set rst = cnnConexion.Execute(sql)

gasto_no_facturado = rst.EOF And rst.BOF

rst.Close
Set rst = Nothing

End Function
Private Function agui_registrado(IDNomina As Long, IDEmpleado As Long) As Boolean
Dim rst As ADODB.Recordset
Dim sql As String
sql = "SELECT IDNom FROM Nom_Detalle WHERE IDNom=" & IDNomina & _
      " AND CodEmp=" & IDEmpleado

Set rst = cnnConexion.Execute(sql)
agui_registrado = Not (rst.EOF And rst.BOF)
Set rst = Nothing
End Function
Private Sub habilitar_boton(Estado As Boolean)
Dim I As Integer

For I = 0 To cmd.UBound
    cmd(I).Enabled = Estado
Next

End Sub

Private Sub setCodGasto(CodInm As String)
Dim strSQL As String
Dim I As Integer, j As Integer
Dim rstGas As ADODB.Recordset

'cargamos los valores para los combos
For I = 0 To 1
    For j = 2 To 4 Step 2
        'DoEvents
        strSQL = "SELECT CodGasto,Titulo FROM Tgastos in '" & _
        gcPath & "\" & CodInm & "\inm.mdb' WHERE Titulo<>'' and Left(Titulo,1)<>'*' ORDER BY " & I + 1
        dtc(j + I).Text = ""
        Set rstGas = cnnConexion.Execute(strSQL)
        dtc(j + I).ListField = IIf(I = 0, "CodGasto", "Titulo")
        dtc(j + I).BoundColumn = IIf(I = 0, "Titulo", "CodGasto")
        Set dtc(j + I).RowSource = rstGas
    Next
    Set rstGas = Nothing
Next
'mostramos los codigos de gastos de aguinaldos facturados
strSQL = "SELECT DISTINCT CodGasto,Descripcion FROM AsignaGasto in '" & _
         gcPath & "\" & CodInm & "\inm.mdb' WHERE Ndoc LIKE 'AGU%' AND Mid(CodGasto,3,1)<>2"
Set rstGas = cnnConexion.Execute(strSQL)
If Not (rstGas.EOF And rstGas.BOF) Then
    Do
        dtc(IIf(InStr(rstGas("Descripcion"), "JUNTA") = 0, 2, 4)) = rstGas("CodGasto")
        rstGas.MoveNext
    Loop Until rstGas.EOF
End If
Set rstGas = Nothing

If dtc(2).MatchedWithList Then dtc(3).Text = dtc(2).BoundText
If dtc(4).MatchedWithList Then dtc(5).Text = dtc(4).BoundText

End Sub

Private Function aguinaldo_no_cerrado(IDNom As Long, CodInm As String, Year As Integer)
Dim rst As ADODB.Recordset
Dim sql As String
Dim Reg As Integer

'validamos que la nomina de aguinaldos generales correspondiente
'a este periodo no haya sido cerrada
sql = "SELECT IDNomina FROM Nom_Inf WHERE IDnomina=" & IDNom
Set rst = cnnConexion.Execute(sql)
aguinaldo_no_cerrado = (rst.EOF And rst.BOF)
rst.Close
Set rst = Nothing

If aguinaldo_no_cerrado = False Then Exit Function

'validamos que los aguinaldos correspondientes al periodo y
'a este inmueble en especifico no haya sido cerrada
Set rst = ejecutar_procedure("agui_cerrado", IDNom, CodInm)
'aguinaldo_no_cerrado = (rst.EOF And rst.BOF)

If aguinaldo_no_cerrado = False Then Exit Function
If Not (rst.EOF And rst.BOF) Then
    If rst("DiasTrab") > 0 And rst("DiasLibres") > 0 Then
        Reg = 2
        rst.Close
        Set rst = Nothing
            
        Dim strUbica As String
        strUbica = gcPath & "\" & CodInm & "\inm.mdb"
    
        sql = "SELECT cargado FROM AsignaGasto in '" & strUbica & _
              "' WHERE Ndoc='AGU" & Year & "' AND cargado <= " & _
              "(SELECT MAX(periodo) FROM Factura in '" & strUbica & _
              "' WHERE Fact Not Like 'CHD%');"
        
        Set rst = cnnConexion.Execute(sql)
        aguinaldo_no_cerrado = rst.EOF And rst.BOF
        If Not (aguinaldo_no_cerrado) Then
            aguinaldo_no_cerrado = (rst.RecordCount < Reg)
        End If
        
            
    End If
End If
rst.Close
Set rst = Nothing

End Function

Private Function cerrar_aguinaldos() As Boolean
Dim sql As String, strD As String
Dim ID As Long
Dim rstNom As ADODB.Recordset
Dim subT As String

sql = txt.Text
subT = "Aguinaldos " & lbl(4)
If Trim(sql) = "" Then
    sql = "Introduzca su contraseña e intente nuevamete. " & vbCrLf & _
          "Si no la tiene, póngase en contacto con el " & vbCrLf & _
          "administrador del sistema."
            
    MsgBox sql, vbExclamation, "Nómina " & subT
    Exit Function
End If

If UCase(txt.Text) <> gcContraseña & lbl(4).Caption Then
    sql = "Introdujo una contraseña incorrecta."
    MsgBox sql, vbExclamation, "Nómina " & subT
    Exit Function
End If
habilitar_boton (False)

'validamos que todos los inmuebles tengan procesados
'como mínimo sus aguinaldos según ley
If Not aguinaldos_procesados() Then
    habilitar_boton (True)
    MsgBox "Faltan inmuebles por procesar los " & subT, vbInformation, App.ProductName
    Exit Function
End If
ID = "312" & Trim(lbl(4))
RtnConfigUtility True, Me.Caption, "Iniciando proceso...", "Cierre Nómina. Aguinaldos"
'comienza el proceso de cierre de la nomina de aguinaldos
'abrimos una transaccion
cnnConexion.BeginTrans

'agrega el registro en la tabla Nom_Inf (Nominas procesadas)
sql = "INSERT INTO Nom_Inf(IDNomina,Fecha," & _
"Usuario,Efectivo) VALUES (" & ID & ",Date()," & "'" & _
gcUsuario & "',Date());"

cnnConexion.Execute sql

'Set rstNom = New ADODB.Recordset
Set rstNom = ModGeneral.ejecutar_procedure("procNomina_Bco", ID)
'rstNom.Open "qdfNomina_bco", cnnConexion, adOpenKeyset, adLockOptimistic
rstNom.Filter = "Cuenta='' or Cuenta = Null"

If Not rstNom.EOF And Not rstNom.BOF Then
    '
    rstNom.MoveFirst
    Dim cppDetalle As String
    Dim aGasto() As String
    '
    Do
        RtnProUtility "Registrando Cuenta por Pagar (" & rstNom!CodInm & ")", _
            rstNom.AbsolutePosition * 6025 / rstNom.RecordCount
        'Call RtnProUtility("Cargando Cpp inm. " & rstNom("CodInm"), rstNom.AbsolutePosition * 6025 / rstNom.RecordCount)
        'GENERA LA FACTURA DE CPP Y ASIGNA EL GASTO PARA SACAR EL CHEQUE
        '
        DoEvents
        strD = FrmFactura.FntStrDoc
        'cppDetalle = DescripcionAguinaldo(ID, rstNom!CodEmp, rstNom!CodInm, _
                    rstNom!Nombre, rstNom!NombreCargo)
        cppDetalle = "CANC. " & subT & " " & rstNom!NombreCargo & " " & _
                    rstNom!Nombre & "(" & rstNom!CodInm & ")"
        '
        cnnConexion.Execute "INSERT INTO Cpp(Tipo,Ndoc,Fact,CodProv,Benef," & _
        "Detalle,Monto,Ivm,Total,FRecep,Fecr,Fven,CodInm,Moneda,Estatus," & _
        "Usuario,Freg) VALUES('NO','" & strD & "','" & Left(ID, 3) & _
        Right(ID, 2) & Format(rstNom.AbsolutePosition, "00") & "','" & _
        sysCodPro & "','" & rstNom("Name") & "' ,'" & cppDetalle & _
        "','" & rstNom("Neto") & "',0,'" & rstNom("Neto") & _
        "',Date(),Date(),'" & DateAdd("d", 30, Date) & "','" & _
        rstNom("CodInm") & "','BS','ASIGNADO','" & gcUsuario & "',Date())"
        '
        
        'ingresa el cargado
        '
        If rstNom("BsTrab") > 0 Then
            aGasto = Split(BuscaCodigoGasto(rstNom("CodInm"), lbl(4)), "|")
            cnnConexion.Execute "INSERT INTO Cargado(Ndoc,CodGasto,Detalle," & _
            "Periodo,Monto,Fecha,Hora,Usuario) IN '" & gcPath & "\" & _
            rstNom("CodInm") & "\inm.mdb' SELECT '" & strD & "','" & _
            aGasto(0) & "',Titulo,'" & aGasto(1) & "','" & rstNom("Neto") _
            & "',Date(),Time(),'" & gcUsuario & "' FROM Tgastos IN '" & gcPath & "\" & _
            rstNom("CodInm") & "\inm.mdb' WHERE CodGasto='" & aGasto(0) & "'"
            
        End If
        If rstNom("BsLib") > 0 Then
            aGasto = Split(BuscaCodigoGasto(rstNom("CodInm"), lbl(4), "JUNTA"), "|")
            cnnConexion.Execute "INSERT INTO Cargado(Ndoc,CodGasto,Detalle," & _
            "Periodo,Monto,Fecha,Hora,Usuario) IN '" & gcPath & "\" & _
            rstNom("CodInm") & "\inm.mdb' SELECT '" & strD & "','" & _
            aGasto(0) & "',Titulo,'" & aGasto(1) & "','" & rstNom("Neto") _
            & "',Date(),Time(),'" & gcUsuario & "' FROM Tgastos IN '" & gcPath & "\" & _
            rstNom("CodInm") & "\inm.mdb' WHERE CodGasto='" & aGasto(0) & "'"
        End If
        '--------------
        '
        rstNom.MoveNext
        '
    Loop Until rstNom.EOF
    Unload FrmUtility
    '
    Call rtnBitacora("Emitiendo listado de cheques")
    '----------------------------+
    'impresion listado de cheque |
    '----------------------------+
    'Dim crReporte As ctlReport
    Dim RepNom As String
    Dim sFormulas() As String
    Dim sOrigenDatos() As String
    Dim sParams()
    Dim NameFile As String
    Dim TituloReporte As String
    
    RepNom = gcPath & "\nomina\"
    sFormulas = Split("Subtitulo='" & subT & "'", "|")
    sOrigenDatos = Split("sac.mdb", "|")
    sParams = Array(ID)
    NameFile = RepNom & "CH" & ID & ".rpt"
    TituloReporte = "Listado de Cheques"
    '-------------------------------+
    '   listado de cheques          |
    '-------------------------------+
    Call Printer_ReporteNomina("nom_chq_agu.rpt", sOrigenDatos(), sFormulas(), _
                                sParams(), NameFile, TituloReporte)
    '-------------------------------+
    '   reporte general al banco    |
    '-------------------------------+
    NameFile = RepNom & "GB" & ID & ".rpt"
    TituloReporte = "Reporte General Banco"
    Call Printer_ReporteNomina("nom_gen_agu.rpt", sOrigenDatos(), sFormulas(), _
                                sParams(), NameFile, TituloReporte)
    '-------------------------------+
    '  reporte general por inmueble |
    '-------------------------------+
    NameFile = RepNom & "RGINMU" & ID & ".rpt"
    TituloReporte = "Reporte General por Inmueble"
    sParams = Array(ID, ID)
    Call Printer_ReporteNomina("nom_report_agu.rpt", sOrigenDatos(), sFormulas(), _
                            sParams(), NameFile, TituloReporte)
    '-------------------------------+
    '  reporte cuadre nómina        |
    '-------------------------------+
    NameFile = RepNom & "CN" & ID & ".rpt"
    TituloReporte = "Reporte Cuadre Nómina"
    sParams = Array(ID)
    Call Printer_ReporteNomina("nom_cuadre_agu.rpt", sOrigenDatos(), sFormulas(), _
                            sParams(), NameFile, TituloReporte)
    '-------------------------------+
    '  reporte nómina general       |
    '-------------------------------+
    NameFile = RepNom & "GE" & ID & ".rpt"
    TituloReporte = "Reporte Aguinaldos General"
    sParams = Array(ID, ID)
    Call Printer_ReporteNomina("nom_report_gen_agu.rpt", sOrigenDatos(), sFormulas(), _
                            sParams(), NameFile, TituloReporte)
    
End If
If Err = 0 Then
    cnnConexion.CommitTrans
    cerrar_aguinaldos = True
Else
    cnnConexion.RollbackTrans
    MsgBox Err.Description, vbCritical, "Error " & Err
End If
End Function

Private Function aguinaldos_procesados() As Boolean
Dim rst As ADODB.Recordset
Dim rstlocal As ADODB.Recordset
Dim sql As String
Dim NEmp As Long

'seleccionamos todos los edificios activos
sql = "SELECT CodInm,Nombre FROM Inmueble WHERE Inactivo=false and CodInm NOT IN ('" & _
    sysCodInm & "','8888')"
Set rst = cnnConexion.Execute(sql)

If Not (rst.EOF And rst.BOF) Then
    RtnConfigUtility True, Me.Caption, "Validación proceso de aguinaldos", "Proceso de Aguinaldos..."
    rst.Sort = "CodInm"
    '   hacemos un recorrido por cada inmueble validando
    '   a.- registro de aguinaldos tabla nom_detalle
    '   b.- registro de aguinaldos en asignagasto
    Do
        RtnProUtility "Verificando inmueble " & rst("CodInm"), rst.AbsolutePosition * 6015 / rst.RecordCount
        DoEvents
        sql = "SELECT Count(CodEmp) from Emp WHERE CodInm='" & rst("CodInm") & "' AND CodEstado<>1"
        Set rstlocal = cnnConexion.Execute(sql)
        If (rstlocal.EOF And rstlocal.BOF) Then GoTo Continuar
        If rstlocal.Fields(0) < 1 Then GoTo Continuar
        Set rstlocal = Nothing
        'seleccionamos los registros en nom_detalle
        sql = "SELECT Nom_Detalle.CodEmp " & _
              "FROM Emp LEFT JOIN Nom_Detalle ON Emp.CodEmp=Nom_Detalle.CodEmp " & _
              "WHERE Emp.CodInm='" & rst("CodInm") & _
              "' AND Nom_Detalle.IDNom=312" & lbl(4)
        Set rstlocal = cnnConexion.Execute(sql)
        If (rstlocal.EOF And rstlocal.BOF) Then
            Unload FrmUtility
            sql = "El Inmueble " & rst("CodInm") & " - " & rst("Nombre") & vbCrLf & _
                "esta 'ACTIVO', pero no aparecen procesados sus aguinaldos." & vbCrLf & _
                "Verifíquelo e inténtelo nuevamente."
            MsgBox sql, vbInformation, App.ProductName
            Exit Function
        End If
        Set rstlocal = Nothing
        
        'seleccionamos los registros en asignagasto
        sql = "SELECT Sum(Nom_Detalle.Dias_Trab) AS NDTra, " & _
              "Sum(Nom_Detalle.Dias_libres) AS NDLib " & _
              "FROM Emp INNER JOIN Nom_Detalle ON Emp.CodEmp=Nom_Detalle.CodEmp " & _
              "GROUP BY Emp.CodInm, Nom_Detalle.IDNom " & _
              "HAVING (((Emp.CodInm)='" & rst("CodInm") & _
              "') AND ((Nom_Detalle.IDNom)=312" & lbl(4) & "));"
        
        Set rstlocal = cnnConexion.Execute(sql)
        
        If (rstlocal.EOF And rstlocal.BOF) Then Exit Function
        NEmp = 0
        If rstlocal("NDTra") > 0 Then NEmp = NEmp + 1
        If rstlocal("NDLib") > 0 Then NEmp = NEmp + 1
        Set rstlocal = Nothing
        
        sql = "SELECT Count(NDoc) FROM AsignaGasto in '" & gcPath & "\" & _
        rst("CodInm") & "\inm.mdb' WHERE Ndoc='AGU" & lbl(4) & "'"
        Debug.Print rst("CodInm")
        Set rstlocal = cnnConexion.Execute(sql)
        If (rstlocal.EOF And rstlocal.BOF) Then Exit Function
        If rstlocal.Fields(0) < NEmp Then Exit Function
Continuar:
        rst.MoveNext
        
    Loop Until rst.EOF
    Unload FrmUtility
    aguinaldos_procesados = True
Else
    Exit Function
End If
End Function

Function BuscaCodigoGasto(CodInm As String, Periodo As String, Optional Tipo As String) As String
Dim strSQL As String
Dim rstGas As ADODB.Recordset

strSQL = "SELECT DISTINCT CodGasto,Cargado FROM AsignaGasto in '" & _
         gcPath & "\" & CodInm & "\inm.mdb' WHERE Ndoc LIKE 'AGU" & Periodo & _
         "%' AND Mid(CodGasto,3,1)<>2 AND Descripcion LIKE '%" & Tipo & "%'"
         
Set rstGas = cnnConexion.Execute(strSQL)

If Not (rstGas.EOF And rstGas.BOF) Then
    BuscaCodigoGasto = rstGas("CodGasto") & "|" & rstGas("Cargado")
End If

Set rstGas = Nothing

End Function

