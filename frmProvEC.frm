VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmProvEC 
   Caption         =   "Estado de Cuenta por Proveedor"
   ClientHeight    =   15
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   2385
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   15
   ScaleWidth      =   2385
   WindowState     =   2  'Maximized
   Begin ComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   480
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Width           =   2385
      _ExtentX        =   4207
      _ExtentY        =   847
      ButtonWidth     =   714
      ButtonHeight    =   688
      Appearance      =   1
      ImageList       =   "ImageList1"
      _Version        =   327682
      BeginProperty Buttons {0713E452-850A-101B-AFC0-4210102A8DA7} 
         NumButtons      =   6
         BeginProperty Button1 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "First"
            Object.Tag             =   ""
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Previous"
            Object.Tag             =   ""
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Next"
            Object.Tag             =   ""
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Last"
            Object.Tag             =   ""
            ImageIndex      =   4
         EndProperty
         BeginProperty Button5 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Print"
            Object.Tag             =   ""
            ImageIndex      =   11
         EndProperty
         BeginProperty Button6 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Close"
            Object.Tag             =   ""
            ImageIndex      =   12
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   6630
      Left            =   195
      TabIndex        =   3
      Top             =   1605
      Width           =   11445
      _ExtentX        =   20188
      _ExtentY        =   11695
      _Version        =   393216
      TabHeight       =   520
      ShowFocusRect   =   0   'False
      TabCaption(0)   =   "Cuentas por Pagar"
      TabPicture(0)   =   "frmProvEC.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lbl(0)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lbl(3)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "grid(0)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "txt(0)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "txt(1)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).ControlCount=   5
      TabCaption(1)   =   "Facturas Pagadas"
      TabPicture(1)   =   "frmProvEC.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "txt(3)"
      Tab(1).Control(1)=   "txt(2)"
      Tab(1).Control(2)=   "grid(1)"
      Tab(1).Control(3)=   "lbl(4)"
      Tab(1).Control(4)=   "lbl(1)"
      Tab(1).ControlCount=   5
      TabCaption(2)   =   "Estado de Cuenta"
      TabPicture(2)   =   "frmProvEC.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "txt(6)"
      Tab(2).Control(1)=   "txt(5)"
      Tab(2).Control(2)=   "txt(4)"
      Tab(2).Control(3)=   "grid(2)"
      Tab(2).Control(4)=   "lbl(2)"
      Tab(2).ControlCount=   5
      Begin VB.TextBox txt 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000002&
         BorderStyle     =   0  'None
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
         Height          =   240
         Index           =   6
         Left            =   -71235
         Locked          =   -1  'True
         TabIndex        =   19
         Text            =   "0,00"
         Top             =   6180
         Width           =   1275
      End
      Begin VB.TextBox txt 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000002&
         BorderStyle     =   0  'None
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
         Height          =   240
         Index           =   5
         Left            =   -72570
         Locked          =   -1  'True
         TabIndex        =   18
         Text            =   "0,00"
         Top             =   6180
         Width           =   1275
      End
      Begin VB.TextBox txt 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000002&
         BorderStyle     =   0  'None
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
         Height          =   240
         Index           =   4
         Left            =   -73905
         Locked          =   -1  'True
         TabIndex        =   17
         Text            =   "0,00"
         Top             =   6180
         Width           =   1275
      End
      Begin VB.TextBox txt 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000002&
         BorderStyle     =   0  'None
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
         Height          =   255
         Index           =   3
         Left            =   -71880
         Locked          =   -1  'True
         TabIndex        =   16
         Text            =   "0,00"
         Top             =   6165
         Width           =   1275
      End
      Begin VB.TextBox txt 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000002&
         BorderStyle     =   0  'None
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
         Height          =   255
         Index           =   2
         Left            =   -73215
         Locked          =   -1  'True
         TabIndex        =   15
         Text            =   "0,00"
         Top             =   6165
         Width           =   1275
      End
      Begin VB.TextBox txt 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000002&
         BorderStyle     =   0  'None
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
         Height          =   255
         Index           =   1
         Left            =   3120
         Locked          =   -1  'True
         TabIndex        =   14
         Text            =   "0,00"
         Top             =   6165
         Width           =   1275
      End
      Begin VB.TextBox txt 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000002&
         BorderStyle     =   0  'None
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
         Height          =   255
         Index           =   0
         Left            =   1785
         Locked          =   -1  'True
         TabIndex        =   13
         Text            =   "0,00"
         Top             =   6165
         Width           =   1275
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid grid 
         Height          =   5370
         Index           =   0
         Left            =   300
         TabIndex        =   4
         Tag             =   "0|1000|500|1000|1200|1200|5500"
         Top             =   600
         Width           =   10770
         _ExtentX        =   18997
         _ExtentY        =   9472
         _Version        =   393216
         Cols            =   7
         FixedCols       =   0
         BackColorFixed  =   -2147483646
         ForeColorFixed  =   -2147483639
         BackColorBkg    =   -2147483636
         GridColor       =   -2147483636
         GridColorFixed  =   -2147483636
         FormatString    =   "Cod|^Fecha|^Tipo|^Doc.|>Debe|>Haber|Detalle"
         _NumberOfBands  =   1
         _Band(0).Cols   =   7
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid grid 
         Height          =   5370
         Index           =   1
         Left            =   -74700
         TabIndex        =   5
         Tag             =   "0|1000|500|1000|1200|1200|5500"
         Top             =   600
         Width           =   10770
         _ExtentX        =   18997
         _ExtentY        =   9472
         _Version        =   393216
         Cols            =   7
         FixedCols       =   0
         BackColorFixed  =   -2147483646
         ForeColorFixed  =   -2147483639
         BackColorBkg    =   -2147483636
         GridColor       =   -2147483636
         GridColorFixed  =   -2147483636
         FormatString    =   "Cod|^Fecha|^Tipo|^Número|>Debe|>Haber|Detalle"
         _NumberOfBands  =   1
         _Band(0).Cols   =   7
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid grid 
         Height          =   5370
         Index           =   2
         Left            =   -74700
         TabIndex        =   6
         Tag             =   "0|1000|500|1000|1200|1200|5500|0"
         Top             =   600
         Width           =   10770
         _ExtentX        =   18997
         _ExtentY        =   9472
         _Version        =   393216
         Cols            =   8
         FixedCols       =   0
         BackColorFixed  =   -2147483646
         ForeColorFixed  =   -2147483639
         BackColorBkg    =   -2147483636
         GridColor       =   -2147483636
         GridColorFixed  =   -2147483636
         FormatString    =   "Cod|^Fecha|^Tipo|^Doc.|>Debe|>Haber|Detalle|"
         _NumberOfBands  =   1
         _Band(0).Cols   =   8
      End
      Begin VB.Label lbl 
         Alignment       =   2  'Center
         Caption         =   "(000)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   4
         Left            =   -73905
         TabIndex        =   12
         Top             =   6195
         Width           =   600
      End
      Begin VB.Label lbl 
         Alignment       =   2  'Center
         Caption         =   "(000)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   3
         Left            =   1100
         TabIndex        =   11
         Top             =   6200
         Width           =   600
      End
      Begin VB.Label lbl 
         Caption         =   "Totales:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   2
         Left            =   -74700
         TabIndex        =   10
         Top             =   6200
         Width           =   1020
      End
      Begin VB.Label lbl 
         Caption         =   "Totales:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   1
         Left            =   -74700
         TabIndex        =   9
         Top             =   6200
         Width           =   1020
      End
      Begin VB.Label lbl 
         Caption         =   "Totales:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   0
         Left            =   300
         TabIndex        =   8
         Top             =   6200
         Width           =   1020
      End
   End
   Begin VB.Frame fraProv 
      Caption         =   "Proveedor:"
      Height          =   885
      Left            =   180
      TabIndex        =   0
      Top             =   600
      Width           =   11415
      Begin MSDataListLib.DataCombo datP 
         Height          =   315
         Index           =   0
         Left            =   420
         TabIndex        =   1
         Top             =   345
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   556
         _Version        =   393216
         ListField       =   "Codigo"
         BoundColumn     =   "NombProv"
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo datP 
         Height          =   315
         Index           =   1
         Left            =   1560
         TabIndex        =   2
         Top             =   345
         Width           =   9690
         _ExtentX        =   17092
         _ExtentY        =   556
         _Version        =   393216
         ListField       =   "NombProv"
         BoundColumn     =   "Codigo"
         Text            =   ""
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
            Picture         =   "frmProvEC.frx":0054
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmProvEC.frx":01D6
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmProvEC.frx":0358
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmProvEC.frx":04DA
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmProvEC.frx":065C
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmProvEC.frx":07DE
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmProvEC.frx":0960
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmProvEC.frx":0AE2
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmProvEC.frx":0C64
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmProvEC.frx":0DE6
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmProvEC.frx":0F68
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmProvEC.frx":10EA
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmProvEC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
    'Formulario estado de cuenta por proveedor
    
    Dim rstP(1) As New ADODB.Recordset
    Dim Pro(2) As String
    
    Private Sub datP_Click(Index As Integer, Area As Integer)
    If Area = 2 Then
        If Index = 0 Then
            datP(1) = datP(0).BoundText
        Else
            datP(0) = datP(1).BoundText
        End If
        'llama la rutina que busca la información estado de cuenta
        Pro(SSTab1.Tab) = datP(0)
        If Not rstP(0).EOF Or Not rstP(0).BOF Then
            rstP(0).MoveFirst
            rstP(0).Find "Codigo='" & Pro(SSTab1.Tab) & "'"
        End If
        Call Cuenta_Proveedor
    End If
    End Sub
    
    Private Sub datP_KeyPress(Index As Integer, KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If Index = 0 Then Call Validacion(KeyAscii, "0123456789")
    If KeyAscii = 13 Then Call datP_Click(Index, 2)
    End Sub

    Private Sub Form_Load()
    'variables locales
    Dim I As Integer
    '
    rstP(0).Open "SELECT * FROM Proveedores ORDER BY Codigo", cnnConexion, adOpenKeyset, _
    adLockOptimistic, adCmdText
    Set datP(0).RowSource = rstP(0)
    '
    rstP(1).Open "SELECT * FROM Proveedores ORDER BY NombProv", cnnConexion, adOpenKeyset, _
    adLockOptimistic, adCmdText
    Set datP(1).RowSource = rstP(1)
    For I = 0 To 2
        Set Grid(I).FontFixed = LetraTitulo(LoadResString(527), 9, , True)
        Set Grid(I).Font = LetraTitulo(LoadResString(528), 8)
        Call centra_titulo(Grid(I), True)
    Next I
    '
    End Sub
    
    Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    '
    rstP(0).Close
    rstP(1).Close
    Set rstP(0) = Nothing
    Set rstP(1) = Nothing
    '
    End Sub

    '---------------------------------------------------------------------------------------------
    '   Rutina:     Cuenta_Proveedor
    '
    '   Busca la situación de un proveedor en específico, facturas pendientes,
    '   cheques pagados, saldo
    '---------------------------------------------------------------------------------------------
    Private Sub Cuenta_Proveedor()
    'variables locales
    Dim Total(1) As Currency
    Dim K%, j%, I%
    Dim rstCP As New ADODB.Recordset
    Dim cmdCP As New ADODB.Command
    '
    cmdCP.ActiveConnection = cnnConexion
    If SSTab1.Tab = 0 Then
        cmdCP.CommandText = "Proveedor_Pendiente"
        K = 0: j = 1
    ElseIf SSTab1.Tab = 1 Then
        cmdCP.CommandText = "Proveedor_Pagado"
        K = 2: j = 3
    Else
        cmdCP.CommandText = "Prov_EdoCta"
        K = 4: j = 5
    End If
    cmdCP.CommandType = adCmdTable
    cmdCP.Parameters("Proveedor") = datP(0)
    Set rstCP = cmdCP.Execute
    Set Grid(SSTab1.Tab).DataSource = rstCP
    Call centra_titulo(Grid(SSTab1.Tab), True)
    For I = 1 To Grid(SSTab1.Tab).Rows - 1
        Total(0) = Total(0) + CCur(Grid(SSTab1.Tab).TextMatrix(I, 4))
        Total(1) = Total(1) + CCur(Grid(SSTab1.Tab).TextMatrix(I, 5))
    Next I
    txt(K) = Format(Total(0), "#,##0.00")
    txt(j) = Format(Total(1), "#,##0.00")
    If SSTab1.Tab = 2 Then txt(6) = Format(CCur(txt(j)) - CCur(txt(K)), "#,##0.00")
    
    End Sub

    Private Sub SSTab1_Click(PreviousTab As Integer)
    If datP(0) <> "" Then
        If Pro(SSTab1.Tab) <> datP(0) Then Call Cuenta_Proveedor
        Pro(SSTab1.Tab) = datP(0)
    End If
    lbl(3) = "(" & Format(Grid(SSTab1.Tab).Rows - 1, "000") & ")"
    lbl(4) = "(" & Format(Grid(SSTab1.Tab).Rows - 1, "000") & ")"

    End Sub

    Private Sub Toolbar1_ButtonClick(ByVal Button As ComctlLib.Button)
    'variables locales
    Dim strSQL As String
    Dim Reporte As String
    Dim errLocal As Long
    Dim cnn As ADODB.Connection
    Dim rpReporte As ctlReport
    '
    With rstP(0)
    '
        Select Case Button.Key
        
            Case "First"    'ir al primer proveedor
                .MoveFirst
                datP(0) = .Fields("Codigo")
                datP(1) = .Fields("NombProv")
                Call rtnLimpiar_Grid(Grid(SSTab1.Tab))
                
            Case "Previous" 'ir al proveedor anterior
                .MovePrevious
                If .BOF Then .MoveLast
                datP(0) = .Fields("Codigo")
                datP(1) = .Fields("NombProv")
                Call rtnLimpiar_Grid(Grid(SSTab1.Tab))
                
            Case "Next" 'ir al proveedor siguiente
                .MoveNext
                If .EOF Then .MoveFirst
                datP(0) = .Fields("Codigo")
                datP(1) = .Fields("NombProv")
                Call rtnLimpiar_Grid(Grid(SSTab1.Tab))
                
            Case "Last" 'ir al último proveedor
                .MoveLast
                datP(0) = .Fields("Codigo")
                datP(1) = .Fields("NombProv")
                Call rtnLimpiar_Grid(Grid(SSTab1.Tab))
                
            Case "Print"    'imprimir pantalla
                Reporte = "cxp_proveedores"
                On Error Resume Next
                If Dir(App.Path & "\Temp.mdb") = "" Then
                    'si no existe la bd temporal la crea en el directoria raiz
                    DBEngine.CreateDatabase App.Path & "\Temp.mdb", dbLangSpanish
                End If
                Set cnn = New ADODB.Connection
                cnn.Open cnnOLEDB & App.Path & "\temp.mdb"
                cnn.Execute "DROP TABLE Proveedor_EDOCTA"
                cnn.Close
                Set cnn = Nothing
                
                Select Case SSTab1.Tab
                    '
                    Case 1, 2 'facturas recibidas y estado de cuenta
                        strSQL = "SELECT Proveedores.Codigo, Cheque.FechaCheque as Fecha, 'CH' AS" _
                        & " Tipo, Format(Cheque.IDCheque,'0000000') as DOC, Cheque_Total.Total_" _
                        & "Cheque as Debe, 0 as Haber, UCASE(Cheque.Concepto) as Descipcion INT" _
                        & "O Proveedor_EdoCta IN '" & App.Path & "\temp.mdb'FRO" _
                        & "M Proveedores INNER JOIN (((Cpp INNER JOIN ChequeFactura ON Cpp.Ndoc" _
                        & "= ChequeFactura.Ndoc) INNER JOIN Cheque_Total ON ChequeFactura.IDChe" _
                        & "que = Cheque_Total.IDCheque) INNER JOIN Cheque ON (Cheque.IDCheque =" _
                        & "ChequeFactura.IDCheque) AND (Cheque_Total.IDCheque = Cheque.IDCheque" _
                        & ")) ON Proveedores.Codigo = Cpp.CodProv Where (((Proveedores.Codigo" _
                        & ") = '" & datP(0) & "'))"
                        cnnConexion.Execute strSQL
                        If SSTab1.Tab = 2 Then  'agregar los otros registros
                            strSQL = "INSERT INTO Proveedor_EDOCTA (Codigo,Fecha,Tipo,DOC,Debe,Ha" _
                            & "ber,Descipcion) SELECT Cpp.CodProv, Cpp.FRecep, Cpp.Tipo, Cpp.Nd" _
                            & "oc, 0 AS Debe, Cpp.Total AS Haber, Cpp.Detalle INTO Proveedor_Ed" _
                            & "oCta IN '" & App.Path & "\temp.mdb' FROM Proveedores INNER JOIN " _
                            & "Cpp ON Proveedores.Codigo = Cpp.CodProv WHERE (((Cpp.CodProv)='" _
                            & datP(0) & "'));"
                            cnnConexion.Execute strSQL
                        End If
                     '
                    Case 0  'Pagos efectuados
                        strSQL = "SELECT Cpp.CodProv, Cpp.FRecep, Cpp.Tipo, Cpp.Ndoc, 0.00 AS Debe" _
                        & ", Cpp.Total AS Haber, Cpp.Detalle INTO Proveedor_EdoCta IN '" & _
                        App.Path & "\temp.mdb' FROM Proveedores INNER JOIN Cpp ON" _
                        & " Proveedores.Codigo = Cpp.CodProv WHERE Cpp.CodProv='" & datP(0) & "' AND (Cpp.Estatus='PENDIENTE' or Cpp.Estatus='ASIGNADO');"
                        cnnConexion.Execute strSQL
                        '
                End Select
                '
                Reporte = gcReport & Reporte & SSTab1.Tab & ".rpt"
                Set rpReporte = New ctlReport
                
                With rpReporte
                    .Salida = crPantalla
                    .Reporte = Reporte
                    .OrigenDatos(0) = App.Path & "\temp.mdb"
                    .OrigenDatos(1) = gcPath & "\sac.mdb"
                    .TituloVentana = SSTab1.TabCaption(SSTab1.Tab)
                    .Imprimir
                    Call rtnBitacora("Impresión " & SSTab1.TabCaption(SSTab1.Tab))
                    '
                End With
                Set rpReporte = Nothing
                
            Case "Close"    'cerrar el formulario
                Me.Hide
                '
        End Select
        
    End With
    '
    End Sub
