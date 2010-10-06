VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmAutCierre 
   AutoRedraw      =   -1  'True
   Caption         =   "Autorizar Cierre Caja"
   ClientHeight    =   4440
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   5205
   Icon            =   "frmAutCierre.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4440
   ScaleWidth      =   5205
   Begin VB.Timer Timer1 
      Left            =   -105
      Top             =   4125
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Salir"
      Height          =   495
      Index           =   1
      Left            =   3525
      TabIndex        =   0
      Top             =   3465
      Width           =   1215
   End
   Begin TabDlg.SSTab stab 
      Height          =   3900
      Left            =   180
      TabIndex        =   1
      Top             =   225
      Width           =   4770
      _ExtentX        =   8414
      _ExtentY        =   6879
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Taquillas"
      TabPicture(0)   =   "frmAutCierre.frx":27A2
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "dGrid"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Command1(2)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "chk"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "Totales"
      TabPicture(1)   =   "frmAutCierre.frx":27BE
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "TXT"
      Tab(1).Control(1)=   "mGrid"
      Tab(1).Control(2)=   "Command1(0)"
      Tab(1).Control(3)=   "LBL(1)"
      Tab(1).Control(4)=   "LBL(0)"
      Tab(1).ControlCount=   5
      Begin VB.CheckBox chk 
         Caption         =   "Todas la Taquillas"
         Height          =   300
         Left            =   180
         TabIndex        =   9
         Top             =   3390
         Visible         =   0   'False
         Width           =   1725
      End
      Begin VB.TextBox TXT 
         Alignment       =   1  'Right Justify
         Height          =   300
         Left            =   -74730
         TabIndex        =   7
         Text            =   "0,00"
         Top             =   3405
         Width           =   1695
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Desautorizar"
         Height          =   495
         Index           =   2
         Left            =   2010
         TabIndex        =   6
         Top             =   3240
         Width           =   1215
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid mGrid 
         Height          =   2265
         Left            =   -74715
         TabIndex        =   5
         Tag             =   "1600|600|1600"
         Top             =   795
         Width           =   4230
         _ExtentX        =   7461
         _ExtentY        =   3995
         _Version        =   393216
         Cols            =   3
         FixedCols       =   0
         BackColorFixed  =   -2147483646
         ForeColorFixed  =   -2147483639
         BackColorBkg    =   -2147483636
         GridColor       =   -2147483636
         GridColorFixed  =   -2147483636
         FocusRect       =   0
         HighLight       =   0
         GridLines       =   0
         GridLinesFixed  =   1
         SelectionMode   =   1
         MergeCells      =   1
         AllowUserResizing=   1
         BorderStyle     =   0
         FormatString    =   "Forma Pago|N.Op|Total"
         _NumberOfBands  =   1
         _Band(0).Cols   =   3
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Autorizar"
         Height          =   495
         Index           =   0
         Left            =   -72990
         TabIndex        =   2
         Top             =   3240
         Width           =   1215
      End
      Begin MSDataGridLib.DataGrid dGrid 
         Bindings        =   "frmAutCierre.frx":27DA
         Height          =   2535
         Left            =   285
         TabIndex        =   3
         Top             =   585
         Width           =   4230
         _ExtentX        =   7461
         _ExtentY        =   4471
         _Version        =   393216
         AllowUpdate     =   0   'False
         HeadLines       =   1
         RowHeight       =   15
         RowDividerStyle =   3
         FormatLocked    =   -1  'True
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
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
         Caption         =   "Taquillas Disponibles"
         ColumnCount     =   2
         BeginProperty Column00 
            DataField       =   "IDTaquilla"
            Caption         =   "Taquilla"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   "00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   8202
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   "Usuario"
            Caption         =   "Cajero"
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
         SplitCount      =   1
         BeginProperty Split0 
            BeginProperty Column00 
               Alignment       =   2
               ColumnWidth     =   840,189
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   2654,929
            EndProperty
         EndProperty
      End
      Begin VB.Label LBL 
         Caption         =   "Saldo de Apertura:"
         Height          =   225
         Index           =   1
         Left            =   -74715
         TabIndex        =   8
         Top             =   3165
         Width           =   1620
      End
      Begin VB.Label LBL 
         Caption         =   "Label1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   330
         Index           =   0
         Left            =   -74745
         TabIndex        =   4
         Top             =   480
         Width           =   4350
      End
   End
   Begin MSAdodcLib.Adodc ado 
      Height          =   360
      Left            =   255
      Top             =   3555
      Visible         =   0   'False
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   635
      ConnectMode     =   1
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Taquillas"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
End
Attribute VB_Name = "frmAutCierre"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
    Public booSel As Boolean

    
    Private Sub Command1_Click(index As Integer)
    'variables locales
    Dim strT As String
    '
    strT = Format(Ado.Recordset("IDTaquilla"), "00")
    Select Case index
        Case 0  'autorizar cuadre y cierre
            If Respuesta("Cierre caja " & strT & vbCrLf & "Saldo Apertura Bs." & Txt & vbCrLf & _
            "Desea continuar?") Then
                
                cnnConexion.Execute "UPDATE Taquillas SET Cuadre=True,OpenSaldo='" _
                & CCur(Txt) & "' WHERE IDTaquilla =" & Ado.Recordset("IDTaquilla")
                
                MsgBox "Reportes Autorizados. Caja N° " & strT, vbInformation, App.ProductName
                    
                Call rtnBitacora("Reportes Autoriz. Caja Nº " & strT)
                    
            End If
            
        Case 1  'cerrar ventana
            Unload Me
            Set frmAutCierre = Nothing
        
        Case 2  'desautorizar
            If booSel Then
                IntTaquilla = IIf(CHK.Value = vbChecked, 99, Ado.Recordset("IDTaquilla"))
                Unload Me
            Else
                If Respuesta("Seguro de desautorizar el cierre de la caja " & strT & "?") Then
                    cnnConexion.Execute "UPDATE Taquillas SET Cuadre=False WHERE IDTaquilla =" & _
                    Ado.Recordset("IDTaquilla")
                    MsgBox "Reportes Des-Autorizados. Caja N° " & Ado.Recordset("IDTaquilla"), vbInformation, _
                    App.ProductName
                    Call rtnBitacora("Reportes Desautoriz. Caja Nº " & Ado.Recordset("IDTaquilla"))
                    '
                End If
            End If
    End Select
    '
    End Sub

    Private Sub dGrid_DblClick()
    If Ado.Recordset("IDTaquilla") <> "" Then
        If booSel Then
            IntTaquilla = IIf(CHK.Value = vbChecked, 99, Ado.Recordset("IDTaquilla"))
            Unload Me
        Else
            Lbl(0) = "Taquilla: " & Format(Ado.Recordset("IDTaquilla"), "00") & " Cajero: " & _
            Ado.Recordset("Usuario")
            Call rtnAutReport(Ado.Recordset("IDTaquilla"))
        End If
    End If
    End Sub

    Private Sub Form_Load()
    'variables locales
    Dim strSql$
    '
    If booSel Then
    
        stab.TabVisible(1) = False
        strSql = "SELECT * FROM Taquillas"
        CHK.Visible = True
        Command1(2).Caption = "Aceptar"
        
    Else    'es cierre de caja
    
        strSql = "SELECT * FROM taquillas WHERE IDTaquilla IN (SELECT IDTaquilla FROM Taquillas" _
        & " IN '" & gcPath & "\sac.mdb' WHERE estado=true);"
        CHK.Visible = False
        
    End If
    
    Ado.ConnectionString = cnnOLEDB + gcPath + "\tablas.mdb"
    Ado.CommandType = adCmdText
    Ado.RecordSource = strSql
    Ado.Refresh
    'Set dGrid.DataSource = ado
    dGrid.Refresh
    Set mGrid.FontFixed = LetraTitulo(LoadResString(527), 7.5, , True)
    Set mGrid.Font = LetraTitulo(LoadResString(528), 8)
    Call centra_titulo(mGrid, True)
    mGrid.ColAlignment(1) = flexAlignCenterCenter
    mGrid.ColAlignment(2) = flexAlignRightCenter
    CenterForm Me
    
    End Sub


    '---------------------------------------------------------------------------------------------
    Private Sub rtnAutReport(intCaja%) 'AUTORIZA EMISION DE REPORTES
    '---------------------------------------------------------------------------------------------
    'variables locales
    Dim rstCaja As ADODB.Recordset
    Dim rstTotales As ADODB.Recordset
    Dim curSaldoO@, curTotal@, Monto$
    Dim strSql$, strMensaje$, i%, J%
    Dim Linea
    '
    Timer1.Interval = 0
    '
    Set rstCaja = New ADODB.Recordset
    Call rtnLimpiar_Grid(mGrid)
    rstCaja.Open "SELECT * FROM Taquillas WHERE IDTaquilla=" & intCaja, cnnConexion, _
    adOpenKeyset, adLockReadOnly, adCmdText
    If rstCaja!Estado Then
        curSaldoO = rstCaja!OpenSaldo
        Set rstTotales = New ADODB.Recordset
        strSql = "SELECT Fpago, Format(Count(Fpago),'00'), Sum(TDFCheques.Monto) From TDFCheque" _
        & "s Where FechaMov = Date() And IDTaquilla=" & intCaja & " and Monto <> 0 GROUP BY Fpa" _
        & "go, FechaMov ORDER BY Fpago"
        rstTotales.Open strSql, cnnConexion, adOpenKeyset, adLockOptimistic, adCmdText
        '
        With rstTotales
            If Not .EOF And Not .BOF Then
                .MoveFirst: J = 1
                mGrid.Rows = .RecordCount + 4
                Do
                    For i = 0 To 2
                        mGrid.TextMatrix(J, i) = IIf(i = 2, Format(.Fields(i), "#,##0.00 "), _
                        .Fields(i))
                    Next i
                    curTotal = curTotal + .Fields(2)
                    .MoveNext: J = J + 1
                Loop Until .EOF
                mGrid.Row = J
                mGrid.Col = 0
                mGrid.RowSel = mGrid.Rows - 1
                mGrid.ColSel = mGrid.Cols - 1
                mGrid.FillStyle = flexFillRepeat
                mGrid.CellFontBold = True
                J = J + 1: mGrid.MergeRow(J) = True
                'aqui muestra el saldo de apertura y total de la caja
                mGrid.TextMatrix(J, 0) = "SALDO DE APERTURA"
                mGrid.TextMatrix(J, 1) = "SALDO DE APERTURA"
                mGrid.TextMatrix(J, 2) = Format(curSaldoO, "#,##0.00 ")
                J = J + 1: mGrid.MergeRow(J) = True
                mGrid.TextMatrix(J, 0) = "TOTAL EN CAJA:"
                mGrid.TextMatrix(J, 1) = "TOTAL EN CAJA:"
                mGrid.TextMatrix(J, 2) = Format(curTotal, "#,##0.00 ")
                stab.Tab = 1
            Else
                MsgBox "Totales en Cero...", vbInformation, App.ProductName
            End If
            .Close
        End With
        Set rstTotales = Nothing
        '
    Else
        MsgBox "Caja N° '" & intCaja & "' Cerrada", vbInformation, _
            "Imposible Autorizar.....", vbInformation, App.ProductName
    End If
    rstCaja.Close
    Set rstCaja = Nothing
    Timer1.Interval = 10000
    '
    End Sub

    Private Sub Timer1_Timer()
    If Ado.Recordset("IDTaquilla") <> "" Then Call rtnAutReport(Ado.Recordset("IDTaquilla"))
    End Sub

    Private Sub txt_KeyPress(KeyAscii As Integer)
    If KeyAscii = 46 Then KeyAscii = 44
    Call Validacion(KeyAscii, "0123456789,")
    If KeyAscii = 13 Then Command1_Click (0)
    
    End Sub

    Private Sub txt_LostFocus()
    If Txt = "" Then Txt = 0
    Txt = Format(Txt, "#,##0.00")
    End Sub
