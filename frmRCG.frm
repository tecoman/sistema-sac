VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmRCG 
   Caption         =   "Form1"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   ControlBox      =   0   'False
   DrawMode        =   12  'Nop
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdRCG 
      Caption         =   "&Salir"
      Height          =   975
      Index           =   1
      Left            =   8415
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   6480
      Width           =   1455
   End
   Begin VB.CommandButton cmdRCG 
      Caption         =   "&Imprimir"
      Height          =   975
      Index           =   0
      Left            =   6795
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   6495
      Width           =   1455
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   7755
      Left            =   465
      TabIndex        =   2
      Top             =   300
      Width           =   10890
      _ExtentX        =   19209
      _ExtentY        =   13679
      _Version        =   393216
      Tabs            =   2
      TabHeight       =   520
      TabCaption(0)   =   "Tab 0"
      TabPicture(0)   =   "frmRCG.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "fraRCG(0)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Lista"
      TabPicture(1)   =   "frmRCG.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "fraRCG(4)"
      Tab(1).Control(1)=   "fraRCG(3)"
      Tab(1).Control(2)=   "gridList"
      Tab(1).ControlCount=   3
      Begin VB.Frame fraRCG 
         Caption         =   "Ordenar por:"
         Height          =   1095
         Index           =   4
         Left            =   -70905
         TabIndex        =   26
         Top             =   6105
         Width           =   1980
         Begin VB.OptionButton optPrint 
            Caption         =   "Inmueble"
            Height          =   255
            Index           =   7
            Left            =   405
            TabIndex        =   28
            Tag             =   "Cobrador=True"
            Top             =   660
            Width           =   1455
         End
         Begin VB.OptionButton optPrint 
            Caption         =   "Fecha/Hora"
            Height          =   255
            Index           =   6
            Left            =   180
            TabIndex        =   27
            Top             =   285
            Value           =   -1  'True
            Width           =   1455
         End
      End
      Begin VB.Frame fraRCG 
         Caption         =   "Recibos emitidos:"
         Height          =   1095
         Index           =   3
         Left            =   -74595
         TabIndex        =   17
         Top             =   6120
         Width           =   3645
         Begin VB.OptionButton optPrint 
            Caption         =   "al Edificio"
            Height          =   540
            Index           =   8
            Left            =   1770
            TabIndex        =   30
            Tag             =   "Cobrador=True"
            Top             =   525
            Width           =   1770
         End
         Begin VB.OptionButton optPrint 
            Caption         =   "Otros"
            Height          =   255
            Index           =   4
            Left            =   1335
            TabIndex        =   25
            Tag             =   "Cobrador=True"
            Top             =   285
            Width           =   900
         End
         Begin VB.OptionButton optPrint 
            Caption         =   "Todos"
            Height          =   255
            Index           =   2
            Left            =   150
            TabIndex        =   24
            Top             =   285
            Value           =   -1  'True
            Width           =   1455
         End
         Begin VB.OptionButton optPrint 
            Caption         =   "Legal"
            Height          =   255
            Index           =   5
            Left            =   2370
            TabIndex        =   23
            Tag             =   "Cobrador=True"
            Top             =   285
            Width           =   1080
         End
         Begin VB.OptionButton optPrint 
            Caption         =   "al Cobrador"
            Height          =   255
            Index           =   3
            Left            =   375
            TabIndex        =   18
            Tag             =   "Cobrador=True"
            Top             =   660
            Width           =   1455
         End
      End
      Begin VB.Frame fraRCG 
         BorderStyle     =   0  'None
         Height          =   7185
         Index           =   0
         Left            =   450
         TabIndex        =   3
         Top             =   345
         Width           =   10095
         Begin VB.CheckBox Check1 
            Caption         =   "Enviado al Edificio"
            Height          =   255
            Index           =   2
            Left            =   6000
            MouseIcon       =   "frmRCG.frx":0038
            MousePointer    =   99  'Custom
            TabIndex        =   29
            Top             =   1155
            Width           =   1875
         End
         Begin VB.CheckBox Check1 
            Caption         =   "Legal"
            Height          =   255
            Index           =   1
            Left            =   4965
            MouseIcon       =   "frmRCG.frx":018A
            MousePointer    =   99  'Custom
            TabIndex        =   22
            Top             =   1155
            Width           =   1260
         End
         Begin VB.CheckBox Check1 
            Caption         =   "Emitido al Cobrador"
            Height          =   255
            Index           =   0
            Left            =   3015
            MouseIcon       =   "frmRCG.frx":02DC
            MousePointer    =   99  'Custom
            TabIndex        =   19
            Top             =   1155
            Width           =   1710
         End
         Begin VB.Frame fraRCG 
            Caption         =   "Facturas Pendientes:"
            Height          =   5535
            Index           =   1
            Left            =   720
            TabIndex        =   4
            Top             =   1515
            Width           =   8655
            Begin VB.TextBox txtRec 
               Alignment       =   2  'Center
               BackColor       =   &H00FFFFFF&
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   2325
               MaxLength       =   2
               TabIndex        =   9
               Top             =   4290
               Width           =   400
            End
            Begin VB.CheckBox check 
               Appearance      =   0  'Flat
               Caption         =   "Todos los propietarios.             recibo(s) pendiente(s)"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   -1  'True
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF0000&
               Height          =   255
               Left            =   360
               MouseIcon       =   "frmRCG.frx":042E
               MousePointer    =   99  'Custom
               TabIndex        =   8
               Top             =   4380
               Width           =   4575
            End
            Begin VB.Frame fraRCG 
               Caption         =   "Destino"
               Height          =   615
               Index           =   2
               Left            =   360
               TabIndex        =   5
               Top             =   4680
               Width           =   4575
               Begin VB.OptionButton optPrint 
                  Caption         =   "Impresora"
                  Height          =   255
                  Index           =   1
                  Left            =   2640
                  TabIndex        =   7
                  Top             =   240
                  Width           =   1455
               End
               Begin VB.OptionButton optPrint 
                  Caption         =   "Ventana"
                  Height          =   255
                  Index           =   0
                  Left            =   840
                  TabIndex        =   6
                  Top             =   240
                  Value           =   -1  'True
                  Width           =   1455
               End
            End
            Begin MSHierarchicalFlexGridLib.MSHFlexGrid gridFact 
               Height          =   3495
               Left            =   405
               TabIndex        =   10
               Tag             =   "2000|2000|2000|1000|0|0"
               Top             =   360
               Width           =   7815
               _ExtentX        =   13785
               _ExtentY        =   6165
               _Version        =   393216
               Cols            =   4
               FixedCols       =   0
               BackColorFixed  =   -2147483646
               ForeColorFixed  =   -2147483639
               GridColor       =   -2147483633
               FormatString    =   "Fact |Periodo |Monto |Imprimir"
               _NumberOfBands  =   1
               _Band(0).Cols   =   4
            End
            Begin VB.Image Img 
               Height          =   195
               Index           =   1
               Left            =   360
               Stretch         =   -1  'True
               Top             =   4080
               Width           =   195
            End
            Begin VB.Label lblRCG 
               Caption         =   "Seleccionar todos los recibos"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   -1  'True
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF0000&
               Height          =   315
               Index           =   2
               Left            =   600
               MouseIcon       =   "frmRCG.frx":0580
               MousePointer    =   99  'Custom
               TabIndex        =   11
               Top             =   4080
               Width           =   4335
            End
         End
         Begin MSDataListLib.DataCombo dtcRCG 
            Height          =   315
            Index           =   0
            Left            =   1680
            TabIndex        =   1
            Top             =   225
            Width           =   1200
            _ExtentX        =   2117
            _ExtentY        =   556
            _Version        =   393216
            ListField       =   "CodInm"
            BoundColumn     =   "Nombre"
            Text            =   ""
            Object.DataMember      =   ""
         End
         Begin MSDataListLib.DataCombo dtcRCG 
            Height          =   315
            Index           =   1
            Left            =   3000
            TabIndex        =   12
            Top             =   225
            Width           =   6195
            _ExtentX        =   10927
            _ExtentY        =   556
            _Version        =   393216
            ListField       =   "Nombre"
            BoundColumn     =   "CodInm"
            Text            =   ""
            Object.DataMember      =   ""
         End
         Begin MSDataListLib.DataCombo dtcRCG 
            Height          =   315
            Index           =   2
            Left            =   1680
            TabIndex        =   13
            Top             =   705
            Width           =   1200
            _ExtentX        =   2117
            _ExtentY        =   556
            _Version        =   393216
            ListField       =   "Codigo"
            BoundColumn     =   "Nombre"
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo dtcRCG 
            Height          =   315
            Index           =   3
            Left            =   3000
            TabIndex        =   14
            Top             =   705
            Width           =   6195
            _ExtentX        =   10927
            _ExtentY        =   556
            _Version        =   393216
            ListField       =   "Nombre"
            BoundColumn     =   "Codigo"
            Text            =   ""
         End
         Begin VB.Label lblRCG 
            Caption         =   "Inmueble:"
            Height          =   315
            Index           =   0
            Left            =   720
            TabIndex        =   0
            Top             =   225
            Width           =   1215
         End
         Begin VB.Label lblRCG 
            Caption         =   "Propieario:"
            Height          =   315
            Index           =   1
            Left            =   720
            TabIndex        =   15
            Top             =   705
            Width           =   1215
         End
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid gridList 
         Height          =   5175
         Left            =   -74595
         TabIndex        =   16
         Top             =   780
         Width           =   9915
         _ExtentX        =   17489
         _ExtentY        =   9128
         _Version        =   393216
         Cols            =   4
         FixedCols       =   0
         BackColorFixed  =   -2147483646
         ForeColorFixed  =   -2147483639
         GridColor       =   -2147483633
         FormatString    =   "Fact |Periodo |Monto |Imprimir"
         _NumberOfBands  =   1
         _Band(0).Cols   =   4
      End
   End
   Begin VB.Image Img 
      Height          =   255
      Index           =   0
      Left            =   840
      Top             =   480
      Width           =   375
   End
End
Attribute VB_Name = "frmRCG"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim strUbica As String
Dim rstPropietario(1) As New ADODB.Recordset
Dim cnnPropietario As New ADODB.Connection
Public Opcion As Integer

Private Sub check_Click()
If check.Value = Checked Then
    txtRec.Enabled = True
Else
    txtRec.Enabled = False
End If
End Sub


Private Sub cmdRCG_Click(Index As Integer)
'
Dim Fila As Integer 'Variables locales
Dim strTitulo As String
Dim rpReporte As ctlReport
'
Select Case Index
    '
    Case 0
    '
        cmdRCG(0).Enabled = False
        cmdRCG(1).Enabled = True
        MousePointer = vbHourglass
        If SSTab1.tab = 0 Then
            Call Printer_CancGastos 'Impresión cancelación de gastos
        Else    'hace otra vaina
            If optPrint(2).Value = True Then
                mcCrit = ""
                strTitulo = ""
            ElseIf optPrint(3).Value = True Then
                mcCrit = "{Recibos_Emision.Cobrador}=True"
                strTitulo = " al Cobrador"
            ElseIf optPrint(4).Value = True Then
                mcCrit = "{Recibos_Emision.Legal}=False and {Recibos_Emision.Cobrador}=False"
                strTitulo = " Otros"
            ElseIf optPrint(5).Value = True Then
                mcCrit = "{Recibos_Emision.Legal}=True"
                strTitulo = " Legal"
            End If
            Set rpReporte = New ctlReport
            With rpReporte
                .OrigenDatos(0) = gcPath & "\sac.mdb"
                .OrigenDatos(1) = gcPath & "\sac.mdb"
                .Reporte = gcReport + "cxc_cobrador.rpt"
                .FormuladeSeleccion = mcCrit
                .Formulas(0) = "Titulo='Emisión Recibos" & strTitulo & "'"
                .Imprimir
                Call rtnBitacora("Printer Relación CxC Cobrador")
                
            End With
            Set rpReporte = Nothing
        End If
        MousePointer = vbDefault
        cmdRCG(0).Enabled = True
        cmdRCG(1).Enabled = True
        '
    Case 1  'Cerrar Formulario
        With gridList
            Fila = 1
            .Col = .Cols - 1
            Do
                .Row = Fila
                If .CellPicture = img(1) Then
                    cnnConexion.Execute "DELETE * FROM Recibos_Emision WHERE CodInm & CodPro  &" _
                    & " Periodo & Fecha ='" & .TextMatrix(Fila, 0) & .TextMatrix(Fila, 1) & _
                    .TextMatrix(Fila, 3) & .TextMatrix(Fila, 6) & "'"
                End If
                Fila = Fila + 1
            Loop Until Fila = .Rows
        End With
        Unload Me
    '
End Select
'
End Sub


Private Sub dtcRCG_Change(Index As Integer)
'
Select Case Index
    '
    Case 0
        With FrmAdmin.objRst
            .Find "CodInm='" & dtcRCG(0) & "'"
            If .EOF Or .BOF Then
                .MoveFirst
                .Find "CodInm='" & dtcRCG(0) & "'"
                If .EOF Or .BOF Then Exit Sub
            End If
            strUbica = !Ubica
        End With
    If Not Dir(gcPath & strUbica & "inm.mdb") = "" Then
        If cnnPropietario.State = 1 Then cnnPropietario.Close
        cnnPropietario.Open cnnOLEDB & gcPath & strUbica & "inm.mdb"
        rstPropietario(0).Open "SELECT * FROM Propietarios WHERE Codigo Not Like 'U%';", _
        cnnPropietario, adOpenKeyset, adLockOptimistic
        Set dtcRCG(2).RowSource = rstPropietario(0)
        Set dtcRCG(3).RowSource = rstPropietario(0)
    End If
End Select
End Sub

Private Sub dtcRCG_Click(Index As Integer, Area As Integer)
If Area = 2 Then
    Call rtnLimpiar_Grid(gridFact)
    Select Case Index
        Case 0, 1
            dtcRCG(IIf(Index = 0, 1, 0)).Text = dtcRCG(Index).BoundText
            dtcRCG(2) = "": dtcRCG(3) = ""
            
        Case 2, 3
            dtcRCG(IIf(Index = 2, 3, 2)).Text = dtcRCG(Index).BoundText
            Call Busca_Facturas
            
    End Select
    '
End If
'
End Sub


Private Sub dtcRCG_KeyPress(Index As Integer, KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
If KeyAscii = 13 Then   'presionó enter
    Call rtnLimpiar_Grid(gridFact)
    Select Case Index
        Case 0, 1
            dtcRCG(IIf(Index = 0, 1, 0)).Text = dtcRCG(Index).BoundText
            dtcRCG(2) = "": dtcRCG(3) = ""
            dtcRCG(2).SetFocus
            
        Case 2, 3
            If Index = 2 And Len(dtcRCG(2)) = 2 Then dtcRCG(2) = "00" & dtcRCG(2)
            If Index = 3 Then
                With rstPropietario(0)
                    '.MoveNext
                    .Find "Nombre Like '%" & dtcRCG(3) & "%'"
                    If .EOF Or .BOF Then
                        .MoveFirst
                        .Find "Nombre Like '%" & dtcRCG(3) & "%'"
                    End If
                    dtcRCG(3) = IIf(.EOF, "", !Nombre)
                End With
            End If
            dtcRCG(IIf(Index = 2, 3, 2)).Text = dtcRCG(Index).BoundText
            If dtcRCG(Index).MatchedWithList Then Call Busca_Facturas
                
    End Select
End If
End Sub

Private Sub Form_Load()
cmdRCG(0).Picture = LoadResPicture("Print", vbResIcon)
cmdRCG(1).Picture = LoadResPicture("Salir", vbResIcon)
img(0).Picture = LoadResPicture("Unchecked", vbResBitmap)
img(1).Picture = LoadResPicture("checked", vbResBitmap)
'
If Opcion = 0 Then  'emisión cancalecion de gastos
    Me.Caption = "Emisión Cancelación de Gastos"
    'SSTab1.TabCaption(0) = Me.Caption
    fraRCG(1).Caption = "Facturas Pendientes: "
    If gcNivel >= nuSUPERVISOR Then 'si tiene nievel usuario
        SSTab1.TabEnabled(0) = False
        SSTab1.tab = 1
        optPrint(2).Enabled = False
        optPrint(4).Enabled = False
        optPrint(5).Enabled = False
        optPrint(3).Value = True
    End If
    Call Config_Grid1
Else
    Me.Caption = "Re-Impresión Cancelacion de Gastos"
    fraRCG(1).Caption = "Facturas Canceladas: "
    txtRec.Visible = False
    check.Visible = False
    SSTab1.TabVisible(1) = False
End If
SSTab1.TabCaption(0) = Me.Caption
Call Config_Grid
Set dtcRCG(0).RowSource = FrmAdmin.objRst
Set dtcRCG(1).RowSource = FrmAdmin.ObjRstNom
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set rstPropietario(0) = Nothing
Set cnnPropietario = Nothing
End Sub

Private Sub Busca_Facturas()
Dim strSQL As String    'variables locales

If dtcRCG(2) <> "" Then
    If Opcion = 0 Then  'Emisión de Cancelación de Gastos
        strSQL = "SELECT Fact,Format(Periodo,'MMM-YYYY'), Format(Saldo,'#,##0.00 '),'' as I FRO" _
        & "M Factura WHERE CodProp='" & dtcRCG(2) & "' AND Not IsNull(Fact) AND Saldo > 0 ORDER" _
        & " BY Periodo DESC;"
    Else    'Reimpresión cancelación de gastos
        strSQL = "SELECT Fact,Format(Periodo,'MMM-YYYY'), Format(Pagado,'#,##0.00 '),'' as R,Fo" _
        & "rmat(Periodo,'MM-YY'),Format(Freg,'ddmmyy') FROM Factura WHERE CodProp='" & _
        dtcRCG(2) & "' AND Not IsNull(Fact) AND Saldo = 0 ORDER BY Periodo DESC;"
    End If
    If rstPropietario(1).State = 1 Then rstPropietario(1).Close
    rstPropietario(1).Open strSQL, cnnPropietario, adOpenStatic, adLockReadOnly
    '
    With rstPropietario(1)
        If .RecordCount > 0 And Not .Fields(0) = "" Then
            Set gridFact.DataSource = rstPropietario(1)
            Call Config_Grid
        End If
    End With
    '
End If
'
End Sub


Private Sub Config_Grid()
With gridFact
    .FontFixed = LetraTitulo(LoadResString(527), 7.5, True)
    .Font = LetraTitulo(LoadResString(528), 7.5)
    If Opcion = 0 Then
        .FormatString = "Fact |Periodo |Monto |Imprimir"
    Else
        .FormatString = "Fact |Periodo |Monto |Re-Imprimir | |"
        .ColWidth(4) = 0
        .ColWidth(5) = 0
    End If
    Call centra_titulo(gridFact, True)
    .ColAlignment(0) = flexAlignCenterCenter
    .ColAlignment(3) = flexAlignCenterCenter
    .Col = 3
    For I = 1 To .Rows - 1
        .Row = I
        Set .CellPicture = img(0)
        .CellPictureAlignment = flexAlignCenterCenter
    Next I
End With
End Sub

Private Sub gridFact_Click()
With gridFact
    If .ColSel = 3 Then
        .Row = .RowSel
        Set .CellPicture = IIf(.CellPicture = img(0), img(1), img(0))
        .CellPictureAlignment = flexAlignCenterCenter
    End If
End With
End Sub


Private Sub gridList_Click()
'
'requiere un nivel mínimo de administrador
If gcNivel <= nuAdministrador Then
    With gridList
        If .ColSel = 7 Then
            .Row = .RowSel
            Set .CellPicture = IIf(.CellPicture = img(0), img(1), img(0))
            .CellPictureAlignment = flexAlignCenterCenter
        End If
    End With
End If
'
End Sub

Private Sub lblRCG_Click(Index As Integer)
If Index = 2 Then
    With gridFact
        .Col = 3
        For I = 1 To .Rows - 1
            .Row = I
            Set .CellPicture = img(1)
            .CellPictureAlignment = flexAlignCenterCenter
        Next I
    End With
End If
End Sub

Private Sub Printer_CancGastos()
'variables locales
Dim Salida As crSalida  'Variables locales
Dim objRst As New ADODB.Recordset
Dim FP(2) As String
Dim strSQL As String, Recibo As String, Destinatario As String
Dim curPago@, Filas&, N&

If optPrint(1).Value = True Then Salida = crImpresora

If check.Value = Checked Then   'Todos los propietarios
    If txtRec = "" Then         'Emision cancelacion de gastos
        strSQL = "SELECT F.Fact,F.Periodo,F.Saldo,F.CodProp,P.Nombre FROM Factura as F INNER JO" _
        & "IN Propietarios as P ON F.CodProp=P.Codigo WHERE Not IsNull(F.Fact) AND F.Saldo > 0 AND F.Fact Not Like 'C" _
        & "H%' ORDER BY F.CodProp,F.Periodo;"
    Else
        strSQL = "SELECT F.Fact,F.Periodo,F.Saldo,F.CodProp,P.Nombre FROM Factura as F INNER JO" _
        & "IN Propietarios as P ON F.CodProp=P.Codigo WHERE F.CodProp IN (SELECT CodProp FROM Factura WHERE Saldo  > " _
        & "0 GROUP BY CodProp HAVING Count(Saldo)<=(SELECT MesesMora FROM Inmueble IN '" & _
        gcPath & "\sac.mdb' WHERE CodInm='" & dtcRCG(0) & "')) AND NOT IsNull(F.Fact) AND F.Sal" _
        & "do>0 AND F.Periodo IN (SELECT DISTINCT TOP " & txtRec & " Periodo FROM Factura WHERE" _
        & " Fact Not Like 'CH%' ORDER BY Periodo DESC)"
        
    End If
    '
    With objRst
        .Open strSQL, cnnPropietario, adOpenStatic, adLockReadOnly
        If Not .EOF And Not .BOF Then    'si no es fin de archivo
            .MoveFirst
            Do
                
                Call Printer_Pago(!Fact, !Saldo, strUbica, dtcRCG(0), dtcRCG(1), !Fact, False, _
                1, Salida)
                '
                cnnConexion.Execute "INSERT INTO Recibos_Emision(CodInm,CodPro,Nombre,Perio" _
                & "do,Monto,ImpresoPor,Fecha,Verificado,Nfact,Edificio) VALUES ('" & _
                dtcRCG(0) & "','" & !CodProp & "','" & !Nombre & "','" & Format(!Periodo, "mmm-yyyy") & "','" & _
                !Saldo & "','" & gcUsuario & "',Date() &  ' ' & Time(),0,'" & !Fact & "',True);"
                        
                .MoveNext
                
            Loop Until .EOF
            '
            Call rtnBitacora("Emisión Canc. de Gastos: " & "En " & _
            IIf(Salida = 0, "Ventana ", "Impresora ") & dtcRCG(0) & " Todos los prop. " _
            & " Parám.: " & txtRec)
        Else
            MsgBox "No existen recibos dentro del criterio de selección. Si considera esto un e" _
            & "rror verifique si está establecido el parámetro 'Meses Vencidos para legal'.", _
            vbInformation, App.ProductName
        End If
    End With
    '
Else
    With gridFact
        .Col = 3: Filas = .Rows - 1
        For I = 1 To Filas
            .Row = I
            '
            If .CellPicture = img(1) And .TextMatrix(I, 0) <> "" Then
                
                If Opcion = 0 Then  'Emisión cancelación de gastos
                    
                    'valida que tenga alguna de las opciones seleccionadas
                    If Check1(0).Value = vbUnchecked And Check1(1).Value = vbUnchecked And Check1(2).Value = vbUnchecked Then
                        MsgBox "Debe marcar una de las opciones:" & vbCrLf & vbCrLf & "1.- Emitido al co" _
                        & "brador" & vbCrLf & "2.- Legal" & vbCrLf & "3.- Enviado al Edificio", _
                        vbInformation, App.ProductName
                        Exit Sub
                    End If
                    '
                    Call Printer_Pago(.TextMatrix(I, 0), CCur(.TextMatrix(I, 2)), strUbica, _
                    dtcRCG(0), dtcRCG(1), .TextMatrix(I, 0), False, 1, Salida)
                    '
                    cnnConexion.Execute "INSERT INTO Recibos_Emision(CodInm,CodPro,Nombre,Perio" _
                    & "do,Monto,ImpresoPor,Fecha,Verificado,Nfact,Cobrador,Legal,Edificio) VALU" _
                    & "ES ('" & dtcRCG(0) & "','" & dtcRCG(2) & "','" & dtcRCG(3) & "','" & _
                    .TextMatrix(.Row, 1) & "','" & .TextMatrix(.Row, 2) & "','" & gcUsuario & _
                    "',Date() & ' ' & Time(),0,'" & .TextMatrix(.Row, 0) & "'," & Check1(0) _
                    & "," & Check1(1) & "," & Check1(2) & ");"
                    '
                    For K = 0 To 2
                        '
                        If Check1(K).Value = vbChecked Then
                            Destinatario = Check1(K).Caption
                            Exit For
                        End If
                        '
                    Next
                    '
                    Call rtnBitacora("Emisión Canc. de Gastos: " & "En " & _
                    IIf(Salida = 0, "Ventana ", "Impresora ") & dtcRCG(0) & "/" & dtcRCG(2) _
                    & "/" & .TextMatrix(I, 1) & " " & Destinatario)
                    
                Else    'Reimpresión cancelación de gastos
                    '
                    objRst.Open "SELECT * FROM MovimientoCaja INNER JOIN Periodos ON Movimient" _
                    & "oCaja.IDRecibo = Periodos.IDRecibo WHERE MovimientoCaja.IDRecibo Like '" _
                    & Right(dtcRCG(0), 2) & dtcRCG(2) & .TextMatrix(I, 5) & "%' and Periodos.Pe" _
                    & "riodo ='" & Format("01/" & .TextMatrix(I, 1), "MM-YY") & "';", _
                    cnnConexion, adOpenKeyset, adLockOptimistic, adCmdText
                    '
                    With objRst
                    '
                        If Not .EOF And Not .BOF Then
                        .MoveLast
                        '
                            FP(0) = !FPago & " - " & !NumDocumentoMovimientoCaja & " - " & _
                            !BancoDocumentoMovimientoCaja
                            FP(1) = !Fpago1 & " - " & !NumDocumentoMovimientoCaja1 & " - " & _
                            !BancoDocumentoMovimientoCaja1
                            FP(2) = !Fpago2 & " - " & !NumDocumentoMovimientoCaja2 & " - " & _
                            !BancoDocumentoMovimientoCaja2
                        Else: FP(0) = "EFECTIVO"
                        End If
                        Recibo = .Fields("MovimientoCaja.IDRecibo")
                        curPago = !Monto - Total_Pagado(!IDPeriodos)
                        .Close
                        '
                    End With
                    '
                    'si el recibo está en la tabla de recibos enviar al edificio
                    'lo elimina
                    cnnConexion.Execute "DELETE FROM Recibos_Enviar WHERE IDRecibo='" & _
                    Recibo & "' AND Fact='" & .TextMatrix(I, 0) & "';", N
                    If N > 0 Then
                        Call rtnBitacora("Eliminado Recibo " & .TextMatrix(I, 0) & " de la lista recibos " _
                        & "por enviar.")
                    End If
                    '
                    Call Printer_Pago(.TextMatrix(I, 0), curPago, strUbica, _
                    dtcRCG(0), dtcRCG(1), Recibo, True, 2, _
                    Salida, FP(0), FP(1), FP(2))
                    '
                    
                    Call rtnBitacora("Printer Canc. de Gastos Inm: " & "En " & _
                    IIf(Salida = 0, "Ventana ", "Impresora ") & dtcRCG(0) & "/" & dtcRCG(2) _
                    & "/" & .TextMatrix(I, 1))
                    '
                    
                End If
                '
            End If
            
        Next
        '
    End With
    '
End If
End Sub

    Private Sub optPrint_Click(Index As Integer): If Index >= 2 Or Index <= 8 Then Call Config_Grid1
    End Sub


Private Sub txtRec_KeyPress(KeyAscii As Integer): Call Validacion(KeyAscii, "1234567890")
End Sub

Private Sub Config_Grid1()
'variables locales
Dim rstList As New ADODB.Recordset
Dim Y As Integer, I As Integer
'
With gridList
    
    .Visible = False
    .FontFixed = LetraTitulo(LoadResString(527), 7.5, True)
    .Font = LetraTitulo(LoadResString(528), 7.5)
    .FormatString = "CodInm. |Apto. |Nombre Propietario |Período |Monto |Impreso por |Fecha |" _
    & "Confirma"
    .ColWidth(0) = 600: .ColWidth(1) = 600
    .ColWidth(2) = 2500: .ColWidth(3) = 1000
    .ColWidth(4) = 1200: .ColWidth(5) = 1200
    .ColWidth(6) = 2000
    .ColWidth(7) = 400
    .ColAlignment(0) = flexAlignCenterCenter
    .ColAlignment(1) = flexAlignCenterCenter
    rstList.Open "Recibos_Emision", cnnConexion, adOpenStatic, adLockReadOnly, adCmdTable
    'filtrar recibos emitidos
    If optPrint(3).Value = True Then    'al cobrador
        rstList.Filter = "Cobrador=True"
    ElseIf optPrint(4).Value = True Then    'otros
        rstList.Filter = "Legal=False and Cobrador=False"
    ElseIf optPrint(5).Value = True Then    'legal
        rstList.Filter = "Legal=True"
    ElseIf optPrint(8).Value = True Then    'edificio
        rstList.Filter = "Edificio=True"
    End If
    If optPrint(6).Value = True Then 'ordenar por fecha/hora
        rstList.Sort = "Fecha"
    Else
        rstList.Sort = "CodInm,Fecha"
    End If
    '
    gridList.Rows = 2
    Call rtnLimpiar_Grid(gridList)
    
    If Not rstList.EOF Or Not rstList.BOF Then
        rstList.MoveFirst
        .Rows = rstList.RecordCount + 1
        Y = 1
        Do
            For I = 0 To 6
                .TextArray((Y * 8) + I) = IIf(I = 4, _
                Format(rstList.Fields(I), "#,##0.00 "), rstList.Fields(I))
            Next I
            .Row = Y
            .Col = .Cols - 1
            Set .CellPicture = IIf(rstList!Verificado, img(1), img(0))
            .CellPictureAlignment = flexAlignCenterCenter
            Y = Y + 1
            rstList.MoveNext
        Loop Until rstList.EOF
        '
    End If
    rstList.Close
    Set rstList = Nothing
    .Visible = True
End With
'
End Sub

'-------------------------------------------------------------------------------------------------
'
'   Funcion:    Total_Pagado
'
'   IDPer:  variable de tipo String identificador del periodo cancelado
'
'   Devuelve el real pagado por un propietario en un mes determinado
'   tomando en cuen la(s) deduccion(es)
'-------------------------------------------------------------------------------------------------
Private Function Total_Pagado(IDPer$) As Currency
Dim strSQL$ 'variables locales  'variables locales
Dim rstPag  As New ADODB.Recordset
strSQL = "SELECT Sum(Monto) as Total FROM Deducciones WHERE IDPeriodos='" & IDPer & "'"
Total_Pagado = 0
'
With rstPag
    '
    .CursorLocation = adUseClient
    .Open strSQL, cnnConexion, adOpenKeyset, adLockOptimistic, adCmdText
    If Not IsNull(!Total) Then Total_Pagado = !Total
    '
End With
rstPag.Close
Set rstPag = Nothing
'
End Function

