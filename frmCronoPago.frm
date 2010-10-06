VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmCronoPago 
   AutoRedraw      =   -1  'True
   Caption         =   "Cheques en Tránsito"
   ClientHeight    =   45
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   3060
   ControlBox      =   0   'False
   LinkTopic       =   "frmCronoPago"
   MDIChild        =   -1  'True
   ScaleHeight     =   11010
   ScaleWidth      =   15240
   WindowState     =   2  'Maximized
   Begin VB.Frame fra 
      Caption         =   "Disponibilidad:"
      Height          =   1395
      Index           =   3
      Left            =   9090
      TabIndex        =   24
      Top             =   6360
      Width           =   2565
      Begin VB.TextBox TXT 
         Alignment       =   1  'Right Justify
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
         Index           =   8
         Left            =   855
         Locked          =   -1  'True
         TabIndex        =   34
         Text            =   "0,00"
         Top             =   945
         Width           =   1575
      End
      Begin VB.TextBox TXT 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   6
         Left            =   870
         Locked          =   -1  'True
         TabIndex        =   26
         Text            =   "0,00"
         Top             =   315
         Width           =   1575
      End
      Begin VB.TextBox TXT 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   5
         Left            =   870
         Locked          =   -1  'True
         TabIndex        =   25
         Text            =   "0,00"
         Top             =   630
         Width           =   1575
      End
      Begin VB.Label lbl 
         Caption         =   "Saldo:"
         Height          =   240
         Index           =   11
         Left            =   135
         TabIndex        =   35
         Top             =   1020
         Width           =   855
      End
      Begin VB.Label lbl 
         Caption         =   "Deuda:"
         Height          =   240
         Index           =   8
         Left            =   120
         TabIndex        =   28
         Top             =   345
         Width           =   855
      End
      Begin VB.Label lbl 
         Caption         =   "Fondo:"
         Height          =   240
         Index           =   7
         Left            =   120
         TabIndex        =   27
         Top             =   705
         Width           =   855
      End
   End
   Begin VB.Frame fra 
      Caption         =   "Totales:"
      Height          =   1395
      Index           =   2
      Left            =   6675
      TabIndex        =   18
      Top             =   6345
      Width           =   2265
      Begin VB.TextBox TXT 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   7
         Left            =   840
         TabIndex        =   32
         Text            =   "0,00"
         Top             =   330
         Width           =   1305
      End
      Begin VB.TextBox TXT 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   4
         Left            =   840
         Locked          =   -1  'True
         TabIndex        =   22
         Text            =   "0,00"
         Top             =   960
         Width           =   1305
      End
      Begin VB.TextBox TXT 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   3
         Left            =   840
         Locked          =   -1  'True
         TabIndex        =   21
         Text            =   "0,00"
         Top             =   645
         Width           =   1305
      End
      Begin VB.Label lbl 
         Caption         =   "Disp.:"
         Height          =   240
         Index           =   10
         Left            =   120
         TabIndex        =   33
         Top             =   360
         Width           =   855
      End
      Begin VB.Label lbl 
         Caption         =   "X Pagar:"
         Height          =   240
         Index           =   6
         Left            =   120
         TabIndex        =   20
         Top             =   990
         Width           =   855
      End
      Begin VB.Label lbl 
         Caption         =   "Tránsito:"
         Height          =   240
         Index           =   5
         Left            =   120
         TabIndex        =   19
         Top             =   675
         Width           =   855
      End
   End
   Begin VB.Frame fra 
      Caption         =   "Filtrar:"
      Height          =   1395
      Index           =   1
      Left            =   375
      TabIndex        =   8
      Top             =   6345
      Width           =   6180
      Begin VB.CommandButton cmd 
         Caption         =   "Ver Todo"
         Height          =   330
         Index           =   2
         Left            =   4515
         Picture         =   "frmCronoPago.frx":0000
         TabIndex        =   39
         Top             =   1005
         Width           =   1545
      End
      Begin VB.OptionButton opt 
         Caption         =   "Impresora"
         Height          =   315
         Index           =   1
         Left            =   2280
         TabIndex        =   38
         Top             =   1013
         Width           =   1215
      End
      Begin VB.OptionButton opt 
         Caption         =   "Ventana"
         Height          =   315
         Index           =   0
         Left            =   1065
         TabIndex        =   36
         Top             =   1013
         Value           =   -1  'True
         Width           =   1215
      End
      Begin VB.TextBox TXT 
         Height          =   315
         Index           =   2
         Left            =   3300
         TabIndex        =   17
         Top             =   645
         Width           =   2775
      End
      Begin VB.TextBox TXT 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   1
         Left            =   990
         TabIndex        =   16
         Top             =   645
         Width           =   1215
      End
      Begin MSComCtl2.DTPicker DTP 
         Height          =   315
         Index           =   0
         Left            =   990
         TabIndex        =   29
         Top             =   285
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   556
         _Version        =   393216
         CalendarTitleBackColor=   -2147483646
         CalendarTitleForeColor=   -2147483639
         CheckBox        =   -1  'True
         Format          =   55967745
         CurrentDate     =   37950
      End
      Begin MSComCtl2.DTPicker DTP 
         Height          =   315
         Index           =   1
         Left            =   3270
         TabIndex        =   30
         Top             =   285
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   556
         _Version        =   393216
         CalendarTitleBackColor=   -2147483646
         CalendarTitleForeColor=   -2147483639
         CheckBox        =   -1  'True
         Format          =   55967745
         CurrentDate     =   37950
      End
      Begin VB.Label lbl 
         Caption         =   "Imprimir en:"
         Height          =   240
         Index           =   12
         Left            =   195
         TabIndex        =   37
         Top             =   1050
         Width           =   870
      End
      Begin VB.Label lbl 
         Caption         =   "Hasta:"
         Height          =   240
         Index           =   9
         Left            =   2670
         TabIndex        =   31
         Top             =   315
         Width           =   615
      End
      Begin VB.Label lbl 
         Caption         =   "&Beneficiario:"
         Height          =   240
         Index           =   4
         Left            =   2340
         TabIndex        =   11
         Top             =   675
         Width           =   1215
      End
      Begin VB.Label lbl 
         Caption         =   "Monto:"
         Height          =   240
         Index           =   3
         Left            =   210
         TabIndex        =   10
         Top             =   675
         Width           =   495
      End
      Begin VB.Label lbl 
         Caption         =   "Desde:"
         Height          =   240
         Index           =   2
         Left            =   210
         TabIndex        =   9
         Top             =   322
         Width           =   615
      End
   End
   Begin VB.Frame fra 
      Height          =   1275
      Index           =   0
      Left            =   360
      TabIndex        =   7
      Top             =   135
      Width           =   11295
      Begin VB.TextBox TXT 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   0
         Left            =   9465
         TabIndex        =   15
         Text            =   "0,00"
         Top             =   300
         Width           =   1560
      End
      Begin VB.Frame Frame1 
         Height          =   555
         Left            =   6255
         TabIndex        =   12
         Top             =   630
         Width           =   4800
         Begin VB.CheckBox chk 
            Caption         =   "Cheques Pagados"
            Height          =   375
            Left            =   60
            Style           =   1  'Graphical
            TabIndex        =   23
            Top             =   135
            Width           =   1545
         End
         Begin VB.CommandButton cmd 
            Caption         =   "Imprimir"
            Height          =   375
            Index           =   1
            Left            =   1635
            Picture         =   "frmCronoPago.frx":030A
            TabIndex        =   14
            Top             =   135
            Width           =   1545
         End
         Begin VB.CommandButton cmd 
            Caption         =   "Salir"
            Height          =   375
            Index           =   0
            Left            =   3210
            TabIndex        =   13
            Top             =   135
            Width           =   1545
         End
      End
      Begin MSDataListLib.DataCombo dtc 
         DataField       =   "CodInm"
         Height          =   315
         Index           =   0
         Left            =   1515
         TabIndex        =   1
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
         Left            =   2625
         TabIndex        =   2
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
         Left            =   1515
         TabIndex        =   4
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
         Left            =   3930
         TabIndex        =   5
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
      Begin VB.Label lbl 
         Caption         =   "&Cuentas:"
         Height          =   240
         Index           =   1
         Left            =   180
         TabIndex        =   3
         Top             =   825
         Width           =   1215
      End
      Begin VB.Label lbl 
         Caption         =   "&Inmueble:"
         Height          =   240
         Index           =   0
         Left            =   180
         TabIndex        =   0
         Top             =   345
         Width           =   1215
      End
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid Grid 
      Height          =   4755
      Left            =   390
      TabIndex        =   6
      Tag             =   "1000|1200|5300|1500|1200|600"
      Top             =   1500
      Width           =   11745
      _ExtentX        =   20717
      _ExtentY        =   8387
      _Version        =   393216
      Cols            =   7
      FixedCols       =   0
      BackColorFixed  =   -2147483646
      ForeColorFixed  =   -2147483639
      BackColorBkg    =   -2147483636
      GridColor       =   -2147483633
      WordWrap        =   -1  'True
      AllowBigSelection=   0   'False
      HighLight       =   2
      GridLinesFixed  =   1
      SelectionMode   =   1
      FormatString    =   "^Cheque|^Fecha Cheque|Beneficiario|>Monto|^Fecha Pago|^Pagar"
      _NumberOfBands  =   1
      _Band(0).Cols   =   7
      _Band(0).GridLinesBand=   1
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   0
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
   Begin VB.Image img 
      Enabled         =   0   'False
      Height          =   150
      Index           =   0
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   165
   End
End
Attribute VB_Name = "frmCronoPago"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
    'variables globales a nivel de módulo
    Dim rstCrono(3) As New ADODB.Recordset
    Dim strOrigenDatos As String
    'Dim blnPote As Boolean
    Public Opcion As Estado    'Tránsito / xEntregar / Entregados
    Enum Estado
        Transito = 0
        XEntregar
        Entregado
    End Enum
    
    Private Sub chk_Click()
    Call listar(CHK.Value)
    If gcNivel < nuSUPERVISOR Then Grid.Enabled = Not CHK.Value
    End Sub

    Private Sub cmd_Click(Index As Integer)
    Select Case Index
        Case 0  'cerrar el formulario
            Unload Me
            Set frmCronoPago = Nothing
        'imprimir
        Case 1: Call Print_report
        Case 2
            dtc(0) = "": dtc(1) = ""
            dtc(2) = "": dtc(3) = ""
            Call listar
    End Select
    End Sub

    Private Sub dtc_Click(Index As Integer, Area As Integer)
    '
    If Area = 2 Then
    
        Select Case Index
        
            Case 0, 1 'codinm y nombre inm
                If Index = 0 Then dtc(1) = dtc(0).BoundText
                If Index = 1 Then dtc(1) = dtc(0).BoundText
                If fra(3).Visible Then
                    With rstCrono(1)
                        .MoveFirst
                        .Find "CodInm ='" & dtc(0) & "'"
                        If Not .EOF Then
                            txt(6) = Format(.Fields("Deuda"), "#,##0.00")
                            txt(5) = Format(.Fields("FondoAct"), "#,##0.00")
                            txt(8) = Format(CCur(txt(5)) - CCur(txt(6)), "#,##0.00")
                        Else
                            txt(6) = "0,00"
                            txt(5) = "0,00"
                            txt(8) = "0,00"
                        End If
                    End With
                End If
                Call Inf_Inmueble
                'If Opcion = XEntregar Then Call listar
                Call listar
                dtc(2).SetFocus
                
            Case 2, 3 'banco y cuenta
                If Index = 2 Then dtc(3) = dtc(2).BoundText
                If Index = 3 Then dtc(2) = dtc(3).BoundText
                MousePointer = vbHourglass
'                For I = 0 To 30000
'                Next
                Call rtnLimpiar_Grid(Grid)
                Call listar
                MousePointer = vbDefault
        End Select
        
    End If
    
    End Sub

    Private Sub dtc_KeyPress(Index As Integer, KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If Index = 0 Then
        Call Validacion(KeyAscii, "0123456789")
        'If Not Len(dtc(0)) > 4 And KeyAscii = 13 Then Call dtc_Click(Index, 2)
    End If
    If KeyAscii = 13 Then Call dtc_Click(Index, 2)
    End Sub


    Private Sub dtp_Change(Index As Integer)
    Call Filtrar
    End Sub

'    Private Sub Form_Activate(): Call Listar
'    '
'    End Sub

    Private Sub Form_Load()
    'variables locales
    '
    img(0).Picture = LoadResPicture("Unchecked", vbResBitmap)
    img(1).Picture = LoadResPicture("Checked", vbResBitmap)
    For i = 0 To 1
        dtp(i) = Date
        dtp(i) = Null
    Next
    'dtp(0).Value = Date
    'dtp(1).Value = Date
    '
    If Opcion = Transito Then
        Me.Caption = "Cheques en Tránsito"
        Grid.FormatString = "^Cheque|^Fecha Cheque|Beneficiario|>Monto|^Fecha Pago|^Pagar|Inm"
        CHK.Visible = False
    ElseIf Opcion = XEntregar Then
        Me.Caption = "Cheques por entregar"
        Grid.FormatString = "^Cheque|^Fecha Cheque|Beneficiario|>Monto|^Fecha Pago|^Pagado|Inm"
        Call listar
    ElseIf Opcion = Entregado Then
        Me.Caption = "Cheques Pagados"
        Grid.Tag = "1000|1200|5300|1500|1200|0|0"
        CHK.Visible = False
        fra(3).Visible = False
    End If
    rstCrono(1).Open "SELECT * FROM Inmueble ORDER BY CodInm", cnnConexion, adOpenKeyset, _
    adLockReadOnly, adCmdText
    rstCrono(2).Open "SELECT * FROM Inmueble ORDER BY Nombre", cnnConexion, adOpenStatic, _
    adLockReadOnly, adCmdText
    
    Set dtc(0).RowSource = rstCrono(1)
    Set dtc(1).RowSource = rstCrono(2)
    '
    Set Grid.FontFixed = LetraTitulo(LoadResString(527), 7.5, , True)
    Set Grid.Font = LetraTitulo(LoadResString(528), 8)
    Call listar
    Call centra_titulo(Grid, True)
    If gcNivel > nuSUPERVISOR Then
        fra(3).Visible = False
        txt(7).Locked = True
    End If
    '
    End Sub

    Private Sub Form_Unload(Cancel As Integer)
    '
    On Error Resume Next
    For i = 0 To 3
        rstCrono(i).Close
        Set rstCrono(i) = Nothing
    Next
    '
    End Sub

    Private Sub CUENTA_INMUEBLE()
    'variables locales
    Dim strSQL As String
    '
    strSQL = "SELECT Bancos.*, Cuentas.* FROM Bancos INNER JOIN Cuentas ON Bancos.IDBanco = Cue" _
    & "ntas.IDBanco WHERE Cuentas.Inactiva=False;"

    If Dir(strOrigenDatos) <> "" Then
        If rstCrono(3).State = 1 Then rstCrono(3).Close
       rstCrono(3).Open strSQL, cnnOLEDB + strOrigenDatos, adOpenStatic, adLockReadOnly, adCmdText
        '
        Set dtc(2).RowSource = rstCrono(3)
        Set dtc(3).RowSource = rstCrono(3)
    Else
        Set dtc(2).RowSource = Nothing
        Set dtc(3).RowSource = Nothing
    End If
    '
    End Sub

    Private Sub Inf_Inmueble()
    
    With rstCrono(1)
        
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
        Grid.Rows = 2
        Call rtnLimpiar_Grid(Grid)
        If !Caja = sysCodCaja Then
            strOrigenDatos = gcPath & "\" + sysCodInm + "\inm.mdb"
            'blnPote = True
        Else
            strOrigenDatos = gcPath & !Ubica & "inm.mdb"
            'blnPote = False
        End If
        Call CUENTA_INMUEBLE
    End With
    '
    End Sub

    '---------------------------------------------------------------------------------------------
    '   Rutina:     Listar
    '
    '   De acuerdo a los parámtetros muestra en el grid el contenido
    '   del ADODB.Recordset
    '---------------------------------------------------------------------------------------------
    Private Sub listar(Optional Pagado As Boolean)
    'variables locales
    Dim strSQL As String
    '
    'Set Grid.DataSource = Nothing
    
    If Opcion = Transito Then   'cheques en tránsito
    
        strSQL = "SELECT Format(Ch.IDCheque,'000000'), Format(Ch.FechaCheque,'dd/mm/yyyy'), Ch." _
        & "Beneficiario, Format(Sum(CD.Monto),'#,##0.00 '), Format(Ch.Fecha,'dd/mm/yyyy'),CH.ID" _
        & "Estado,CD.Codinm FROM Cheque as CH INNER JOIN ChequeDetalle as CD ON Ch.Clave = CD.Clave GR" _
        & "OUP BY Ch.IDCheque,Ch.FechaCheque, Ch.Beneficiario, Ch.Fecha,ch.Cuenta,CH.IDEstado,CD" _
        & ".CodInm,CH.IDestado,Cd.CodInm HAVING " & IIf(dtc(3) = "", "", "Ch.Cuenta ='" & dtc(3) & "' AN" _
        & "D ") & "(CH.IDEstado=" & Opcion & " or CH.IDEstado=1)" & IIf(dtc(0) <> "", " AND CD.C" _
        & "odinm='" & dtc(0) & "';", ";")
    
        'strSql = "SELECT Format(Ch.IDCheque,'000000'), Format(Ch.FechaCheque,'dd/mm/yyyy'), Ch." _
        & "Beneficiario, Format(Sum(CD.Monto),'#,##0.00 '), Format(Ch.Fecha,'dd/mm/yyyy'),CH.ID" _
        & "Estado FROM Cheque as CH INNER JOIN ChequeDetalle as CD ON Ch.Clave = CD.Clave GR" _
        & "OUP BY Ch.IDCheque,Ch.FechaCheque, Ch.Beneficiario, Ch.Fecha,ch.Cuenta,CH.IDEstado,CD" _
        & ".CodInm,CH.IDestado HAVING " & IIf(dtc(3) = "", "", "Ch.Cuenta ='" & dtc(3) & "' AN" _
        & "D ") & "(CH.IDEstado=" & Opcion & " or CH.IDEstado=1)" & IIf(dtc(0) <> "", " AND CD.C" _
        & "odinm='" & dtc(0) & "';", ";")
        
    ElseIf Opcion = XEntregar Then 'cheques por pagar
    
        If Not Pagado Then
        
            strSQL = "SELECT Format(Ch.IDCheque,'000000'), Format(Ch.FechaCheque,'dd/mm/yyyy')," _
            & "Ch.Beneficiario,Format(Sum(CD.Monto),'#,##0.00 '), Format(Ch.Fecha,'dd/mm/yyyy'),'',CD.CodInm" _
            & " FROM Cheque as CH INNER JOIN ChequeDetalle as CD ON Ch.Clave = CD.Clave GROUP B" _
            & "Y Ch.IDCheque,Ch.FechaCheque, Ch.Beneficiario, Ch.Fecha,ch.Cuenta,CH.IDEstado,CD" _
            & ".CodInm HAVING CH.IDEstado=" & Opcion & IIf(dtc(0) <> "", " AND CD.CodInm='" _
            & dtc(0) & "'", "") & IIf(dtc(3) = "", "", "AND Ch.Cuenta ='" _
            & dtc(3) & "'")
        
            'strSql = "SELECT Format(Ch.IDCheque,'000000'), Format(Ch.FechaCheque,'dd/mm/yyyy')," _
            & "Ch.Beneficiario,Format(Sum(CD.Monto),'#,##0.00 '), Format(Ch.Fecha,'dd/mm/yyyy')" _
            & " FROM Cheque as CH INNER JOIN ChequeDetalle as CD ON Ch.Clave = CD.Clave GROUP B" _
            & "Y Ch.IDCheque,Ch.FechaCheque, Ch.Beneficiario, Ch.Fecha,ch.Cuenta,CH.IDEstado,CD" _
            & ".CodInm HAVING CH.IDEstado=" & Opcion & IIf(dtc(0) <> "", " AND CD.CodInm='" _
            & dtc(0) & "'", "") & IIf(dtc(3) = "", "", "AND Ch.Cuenta ='" _
            & dtc(3) & "'")
        Else
            strSQL = "SELECT Format(Ch.IDCheque,'000000'), Format(Ch.FechaCheque,'dd/mm/yyyy')," _
            & "Ch.Beneficiario, Format(Sum(CD.Monto),'#,##0.00 '), Format(Ch.Fecha,'dd/mm/yyyy')" _
            & " FROM Cheque as CH INNER JOIN ChequeDetalle as CD ON Ch.Clave = CD.Clave " _
            & "GROUP BY Ch.IDCheque,Ch.FechaCheque, Ch.Beneficiario, Ch.Fecha,ch.Cuenta,CH.IDEs" _
            & "tado,CD.CodInm HAVING CH.IDEstado=" & Entregado & " AND Fecha=Date()" & IIf(dtc(0) <> "", _
            " AND CD.CodInm='" & dtc(0) & "'", "") & IIf(dtc(3) = "", "", "AND Ch.Cuenta ='" _
            & dtc(3) & "'")
        
            'strSql = "SELECT Format(Ch.IDCheque,'000000'), Format(Ch.FechaCheque,'dd/mm/yyyy')," _
            & "Ch.Beneficiario, Format(Sum(CD.Monto),'#,##0.00 '), Format(Ch.Fecha,'dd/mm/yyyy')" _
            & " FROM Cheque as CH INNER JOIN ChequeDetalle as CD ON Ch.Clave = CD.Clave " _
            & "GROUP BY Ch.IDCheque,Ch.FechaCheque, Ch.Beneficiario, Ch.Fecha,ch.Cuenta,CH.IDEs" _
            & "tado,CD.CodInm HAVING CH.IDEstado=" & Entregado & " AND Fecha=Date()" & IIf(dtc(0) <> "", _
            " AND CD.CodInm='" & dtc(0) & "'", "") & IIf(dtc(3) = "", "", "AND Ch.Cuenta ='" _
            & dtc(3) & "'")
        End If
        
    ElseIf Opcion = Entregado Then  'cheques pagados
        '
        strSQL = "SELECT Format(Ch.IDCheque,'000000'), Format(Ch.FechaCheque,'dd/mm/yyyy'), Ch." _
        & "Beneficiario, Format(Sum(CD.Monto),'#,##0.00 '), Format(Ch.Fecha,'dd/mm/yyyy') FROM " _
        & "Cheque as CH INNER JOIN ChequeDetalle as CD ON Ch.Clave = CD.Clave GROUP BY Ch" _
        & ".IDCheque,Ch.FechaCheque, Ch.Beneficiario, Ch.Fecha,ch.Cuenta,CH.IDEstado,CD.CodInm " _
        & "HAVING CH.IDEstado=" & Opcion & IIf(dtc(0) <> "", " AND CD.Codinm='" & dtc(0) & "'", "") & _
        IIf(dtc(3) = "", "", " AND CH.Cuenta='" & dtc(3) & "'")
        '
        
        '
        'strSql = "SELECT Format(Ch.IDCheque,'000000'), Format(Ch.FechaCheque,'dd/mm/yyyy'), Ch." _
        & "Beneficiario, Format(Sum(CD.Monto),'#,##0.00 '), Format(Ch.Fecha,'dd/mm/yyyy') FROM " _
        & "Cheque as CH INNER JOIN ChequeDetalle as CD ON Ch.Clave = CD.Clave GROUP BY Ch" _
        & ".IDCheque,Ch.FechaCheque, Ch.Beneficiario, Ch.Fecha,ch.Cuenta,CH.IDEstado,CD.CodInm " _
        & "HAVING CH.IDEstado=" & Opcion & IIf(dtc(0) <> "", " AND CD.Codinm='" & dtc(0) & "'", "") & _
        IIf(dtc(3) = "", "", " AND CH.Cuenta='" & dtc(3) & "'")
        '
    End If
    'abre el ADODB.Recordset
    If rstCrono(0).State = 1 Then rstCrono(0).Close
    rstCrono(0).Open strSQL, cnnConexion, adOpenKeyset, adLockOptimistic, adCmdText
    
    If Opcion = Entregado Then
        Set Grid.DataSource = rstCrono(0)
        Call centra_titulo(Grid, True)
    Else
        Call Filtrar
    End If
    '
    End Sub

    Private Sub Grid_Click()
    'variables locales
    Dim Tipo As Estado
    '
    If Opcion = Entregado Or (Opcion = Transito And gcNivel > nuSUPERVISOR) Then Exit Sub
    '
    With Grid
    '
        If .Rows > 1 And IsNumeric(.TextMatrix(.RowSel, 0)) Then
        '
        .Col = 5
        
        Set .CellPicture = IIf(.CellPicture = img(0), img(1), img(0))
        .CellPictureAlignment = flexAlignCenterCenter
        'si la ventana es cheques en tránsito
        If Opcion = Transito Then
            'Tipo = IIf(Opcion = XEntregar, IIf(Tipo = Transito, XEntregar, Entregado), Tipo)
            Tipo = IIf(.CellPicture = img(1), XEntregar, Transito)
            .Col = 0
            cnnConexion.Execute "UPDATE Cheque SET IDEstado=" & Tipo & " WHERE IDCheque=" & .Text
            Call rtnBitacora("Cheque " & .Text & IIf(Tipo = XEntregar, " por entregar", _
            " colocado en tránsito"))
            If Tipo = Transito Then
                txt(3) = Format(CCur(txt(3)) + CCur(.TextMatrix(.RowSel, 3)), "#,##0.00")
                txt(4) = Format(CCur(txt(4)) - CCur(.TextMatrix(.RowSel, 3)), "#,##0.00")
            ElseIf Tipo = XEntregar Then
                txt(3) = Format(CCur(txt(3)) - CCur(.TextMatrix(.RowSel, 3)), "#,##0.00")
                txt(4) = Format(CCur(txt(4)) + CCur(.TextMatrix(.RowSel, 3)), "#,##0.00")
            End If
            .Refresh
        ElseIf Opcion = XEntregar Then  'cheques a entregar
            Tipo = IIf(.CellPicture = img(1), Entregado, XEntregar)
            .Col = 0
            cnnConexion.Execute "UPDATE Cheque SET IDEstado=" & Tipo & ",Fecha=Date(),Hora=Time" _
            & "(),Usario='" & gcUsuario & "' WHERE IDCheque=" & .Text
            Call rtnBitacora("Cheque " & .Text & IIf(Tipo = Entregado, " entregado", _
            " colocado por entregar"))
        End If
        '
        End If
        '
    End With
    '
    End Sub



    Private Sub txt_KeyPress(Index As Integer, KeyAscii As Integer)
    
    Select Case Index
        Case 1
            If KeyAscii = 46 Then KeyAscii = 44
            Call Validacion(KeyAscii, "0123456789,")
            If KeyAscii = 13 Then   'presionó [enter]
               If IsNumeric(txt(1)) Then txt(1) = Format(CCur(txt(1)), "#,##0.00")
               Call Filtrar
            End If
            
        Case 2
            'convierte todo en mayúsculas
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
            'If KeyAscii = 13 Then   'presionó [enter]
                'Call Filtrar
            'End If
        Case 7
            Call Validacion(KeyAscii, "0123456789")
            If KeyAscii = 13 Then txt(7) = Format(txt(7), "#,##0.00")
    End Select
    End Sub

    Private Sub Filtrar()
    'variables locales
    Dim strFecha As String
    Dim curMonto As Currency
    Dim strBenef As String
    Dim strFiltro As String
    Dim i As Long
    '
    If rstCrono(0).State = 0 Then rstCrono(0).Open
    'If rstCrono(0).State = 1 Then   'si está abierto el ADODB.Recordset
        If IsDate(dtp(0)) And IsDate(dtp(1)) Then  'filtra por fecha
            strFiltro = "Expr1001 >= " & dtp(0) & " AND Expr1001 <= " & dtp(1)
        ElseIf IsDate(dtp(0)) Then
            strFiltro = "Expr1001 >=" & dtp(0)
        ElseIf IsDate(dtp(1)) Then
            strFiltro = "Expr1001 <=" & dtp(1)
        End If
        If IsNumeric(txt(1)) Then   'filtra por monto
            strFiltro = strFiltro & IIf(strFiltro = "", "", " AND ") & "Expr1003 = " & txt(1)
        End If
        If txt(2) <> "" Then    'filtra por beneficiario
            strFiltro = strFiltro & IIf(strFiltro = "", "", " AND ") & "Beneficiario LIKE '" & txt(2) & "*'"
        End If
'    Else
'        rstCrono(0).Open
'    End If
    
    'Set Grid.DataSource = Nothing
    rstCrono(0).Filter = strFiltro
    
    'Set Grid.DataSource = rstCrono(0)
    If Not rstCrono(0).EOF And Not rstCrono(0).BOF Then
        rstCrono(0).MoveFirst
        Grid.Rows = rstCrono(0).RecordCount + 1
        i = 1
        Do
            With Grid
                
                For X = 0 To rstCrono(0).Fields.count - 1
                    .TextMatrix(i, X) = rstCrono(0).Fields(X)
                Next
                
            End With
            i = i + 1
            rstCrono(0).MoveNext
        Loop Until rstCrono(0).EOF
        Call Inicializar_Grid
    End If
    rstCrono(0).Close
    'Set rstCrono(0) = Nothing
    Call centra_titulo(Grid, True)
    
    '
    End Sub


    Private Sub Inicializar_Grid()
    '
    Grid.Visible = False
    txt(3) = "0,00"
    txt(4) = "0,00"
    With Grid
        '
        .Col = 5
        For i = 1 To .Rows - 1
            .Row = i
            If Opcion = Transito Then
                Set .CellPicture = img(.Text)
                txt(3) = Format(CCur(txt(3)) + CCur(.TextMatrix(i, 3)), "#,##0.00")
                If .Text = "1" Then txt(4) = Format(CCur(txt(4)) + _
                CCur(.TextMatrix(i, 3)), "#,##0.00")
                .Text = ""
            ElseIf Opcion = XEntregar Then
                Set .CellPicture = img(0)
            End If
            .CellPictureAlignment = flexAlignCenterCenter
        Next i
        .Visible = True
        If .Rows > 19 Then .TopRow = .Rows - 19
    End With
    '
    End Sub

    '---------------------------------------------------------------------------------------------
    '   Rutina:     Print_Report
    '
    '   Imprime el reporte según la opción que este utilizando el
    '   usuario
    '---------------------------------------------------------------------------------------------
    Private Sub Print_report()
    'variables locales
    Dim cnnLocal As New ADODB.Connection
    Dim BD As Database
    Dim TDF As TableDef
    Dim strTitulo As String
    Dim blnExiste As Boolean
    Dim rpReporte As ctlReport
    '
    On Error Resume Next
    If Opcion = Transito Then
        strTitulo = "Cheques en Tránsito"
    ElseIf Opcion = Entregado Then
        strTitulo = "Cheques Pagados"
    ElseIf Opcion = XEntregar Then
        strTitulo = "Cheques por Entregar"
        If CHK.Value = 1 Then strTitulo = "Cheques Pagados"
    End If
    'si no existe la bd local la crea
    If Dir(App.Path & "\Temp.mdb") = "" Then
        DBEngine.CreateDatabase "Temp.mdb", dbLangSpanish
    End If
    
    Set BD = DBEngine.OpenDatabase(App.Path & "\Temp.mdb")
    For Each TDF In BD.TableDefs
        If UCase(TDF.Name) = UCase("chq_XEntregar") Then blnExiste = True: Exit For
    Next
    cnnLocal.Open cnnOLEDB + App.Path + "\temp.mdb"
    'si no existe la tabla la crea
    If Not blnExiste Then
        cnnLocal.Execute "CREATE TABLE chq_XEntregar (Cheque LONG,FechaCheque DATETIME,Benef" _
        & "iciario TEXT(150),Monto Currency,FechaPago DATETIME,CodInm TEXT(6))"
    Else
        cnnLocal.Execute "DELETE * FROM chq_XEntregar"
    End If
    'introduce los datos del grid
    With Grid
        For i = 1 To .Rows - 1
            If .TextMatrix(i, 0) <> "" Then
            cnnLocal.Execute "INSERT INTO chq_XEntregar (Cheque,FechaCheque,Beneficiario,Monto,FechaPago,CodInm) " _
            & "VALUES ('" & .TextMatrix(i, 0) & "','" & .TextMatrix(i, 1) & "','" & _
            .TextMatrix(i, 2) & "','" & .TextMatrix(i, 3) & "','" & .TextMatrix(i, 4) & "','" & .TextMatrix(i, 6) & "')"
            End If
        Next
        cnnLocal.Close
    End With
    Set cnnLocal = Nothing
    BD.Close
    Set BD = Nothing
'
    Set rpReporte = New ctlReport
    With rpReporte
        .Reporte = gcReport & "\cxp_cheques.rpt"
        .OrigenDatos(0) = App.Path & "\Temp.mdb"
        .Formulas(0) = "Titulo='" & strTitulo & "'"
        If opt(0).Value = True Then
            .Salida = crPantalla
        Else
            .Salida = crImpresora
        End If
        .TituloVentana = strTitulo
        .Imprimir
        Call rtnBitacora("Impresión " & strTitulo)
        
    End With
'    '
    If Err.Number <> 0 Then
        MsgBox "Ha ocurrido errores durante el proceso..." & vbCrLf & "Consulte al administrado" _
        & "r del sistema.", vbInformation, App.ProductName
        Call rtnBitacora(Err.Number & " - " & Err.Description)
    End If
    '
    End Sub

Private Sub TXT_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
If Index = 2 Then Call Filtrar
End Sub
