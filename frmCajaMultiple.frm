VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmCajaMultiple 
   Caption         =   "Pagos Múltiples"
   ClientHeight    =   8985
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   13050
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8985
   ScaleWidth      =   13050
   Begin VB.CommandButton cmd 
      Caption         =   "&Salir"
      Height          =   495
      Index           =   0
      Left            =   10560
      TabIndex        =   1
      Top             =   3165
      Width           =   1215
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   8295
      Left            =   345
      TabIndex        =   0
      Top             =   330
      Width           =   12630
      _ExtentX        =   22278
      _ExtentY        =   14631
      _Version        =   393216
      Tabs            =   1
      TabHeight       =   520
      TabCaption(0)   =   "Datos Generales"
      TabPicture(0)   =   "frmCajaMultiple.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "FRA"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "cmd(1)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      Begin VB.CommandButton cmd 
         Caption         =   "Command1"
         Height          =   495
         Index           =   1
         Left            =   6315
         TabIndex        =   14
         Top             =   4035
         Width           =   1215
      End
      Begin VB.Frame FRA 
         Caption         =   "Datos del Pago:"
         Height          =   2835
         Left            =   315
         TabIndex        =   2
         Top             =   705
         Width           =   10755
         Begin VB.ComboBox Cmb 
            DataField       =   "FPago"
            DataSource      =   "ADOcontrol(0)"
            Height          =   315
            Index           =   5
            ItemData        =   "frmCajaMultiple.frx":001C
            Left            =   405
            List            =   "frmCajaMultiple.frx":0029
            Sorted          =   -1  'True
            TabIndex        =   7
            ToolTipText     =   "Forma de Pago 1"
            Top             =   1710
            Width           =   1305
         End
         Begin VB.TextBox Txt 
            Alignment       =   1  'Right Justify
            DataField       =   "MontoCheque"
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "#,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   8202
               SubFormatType   =   0
            EndProperty
            DataSource      =   "ADOcontrol(0)"
            Height          =   315
            Index           =   1
            Left            =   6765
            TabIndex        =   6
            ToolTipText     =   "Doc. Monto 1"
            Top             =   1710
            Width           =   1515
         End
         Begin VB.ComboBox Cmb 
            DataField       =   "BancoDocumentoMovimientoCaja"
            DataSource      =   "ADOcontrol(0)"
            Height          =   315
            Index           =   2
            ItemData        =   "frmCajaMultiple.frx":004E
            Left            =   3630
            List            =   "frmCajaMultiple.frx":0050
            Sorted          =   -1  'True
            TabIndex        =   5
            ToolTipText     =   "Documento Banco 1"
            Top             =   1710
            Width           =   1785
         End
         Begin VB.CommandButton Command 
            Height          =   255
            Index           =   1
            Left            =   6450
            Picture         =   "frmCajaMultiple.frx":0052
            Style           =   1  'Graphical
            TabIndex        =   4
            Top             =   1740
            Width           =   255
         End
         Begin VB.TextBox Txt 
            Alignment       =   1  'Right Justify
            DataField       =   "NumDocumentoMovimientoCaja"
            DataSource      =   "ADOcontrol(0)"
            Height          =   315
            Index           =   6
            Left            =   1770
            MaxLength       =   11
            TabIndex        =   3
            ToolTipText     =   "Nro.Documento 1"
            Top             =   1710
            Width           =   1785
         End
         Begin MSMask.MaskEdBox MskFecha 
            Bindings        =   "frmCajaMultiple.frx":019C
            DataField       =   "FechaChequeMovimientoCaja"
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "dd/MM/yyyy"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   8202
               SubFormatType   =   3
            EndProperty
            DataSource      =   "ADOcontrol(0)"
            Height          =   315
            Index           =   1
            Left            =   5445
            TabIndex        =   8
            TabStop         =   0   'False
            ToolTipText     =   "Documento Fecha 1"
            Top             =   1710
            Width           =   1290
            _ExtentX        =   2275
            _ExtentY        =   556
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   10
            Format          =   "dd/mm/yyyy"
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin MSDataListLib.DataCombo Dat 
            DataField       =   "AptoMovimientoCaja"
            DataSource      =   "ADOcontrol(0)"
            Height          =   315
            Index           =   2
            Left            =   1920
            TabIndex        =   15
            ToolTipText     =   "Codigo del Propietario"
            Top             =   915
            Width           =   1050
            _ExtentX        =   1852
            _ExtentY        =   556
            _Version        =   393216
            ListField       =   "Codigo"
            BoundColumn     =   "Nombre"
            Text            =   " "
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
         Begin MSDataListLib.DataCombo Dat 
            DataField       =   "InmuebleMovimientoCaja"
            DataSource      =   "ADOcontrol(0)"
            Height          =   315
            Index           =   0
            Left            =   1920
            TabIndex        =   16
            ToolTipText     =   "Codigo del Inmueble"
            Top             =   480
            Width           =   1050
            _ExtentX        =   1852
            _ExtentY        =   556
            _Version        =   393216
            BackColor       =   -2147483643
            ListField       =   "CodInm"
            BoundColumn     =   "CodInm"
            Text            =   " "
            Object.DataMember      =   ""
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
         Begin MSDataListLib.DataCombo Dat 
            Height          =   315
            Index           =   1
            Left            =   2985
            TabIndex        =   17
            ToolTipText     =   "Nombre del Inmueble"
            Top             =   495
            Width           =   3915
            _ExtentX        =   6906
            _ExtentY        =   556
            _Version        =   393216
            ListField       =   "Nombre"
            BoundColumn     =   "Nombre"
            Text            =   " "
            Object.DataMember      =   ""
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
         Begin MSDataListLib.DataCombo Dat 
            Bindings        =   "frmCajaMultiple.frx":01BE
            Height          =   315
            Index           =   3
            Left            =   2985
            TabIndex        =   18
            ToolTipText     =   "Nombre del Propietario"
            Top             =   915
            Width           =   3915
            _ExtentX        =   6906
            _ExtentY        =   556
            _Version        =   393216
            ListField       =   "Nombre"
            BoundColumn     =   "Codigo"
            Text            =   " "
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
         Begin VB.Label Lbl 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   3
            Left            =   435
            TabIndex        =   20
            Top             =   975
            Width           =   1470
         End
         Begin VB.Label Lbl 
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   0
            Left            =   435
            TabIndex        =   19
            Top             =   525
            Width           =   1470
         End
         Begin VB.Label Lbl 
            Alignment       =   2  'Center
            BackColor       =   &H80000002&
            Caption         =   "Forma de Pago"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000009&
            Height          =   285
            Index           =   5
            Left            =   405
            TabIndex        =   13
            Top             =   1410
            Width           =   1365
         End
         Begin VB.Label Lbl 
            Alignment       =   2  'Center
            BackColor       =   &H80000002&
            Caption         =   "Monto"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000009&
            Height          =   285
            Index           =   4
            Left            =   6735
            TabIndex        =   12
            Top             =   1410
            Width           =   1545
         End
         Begin VB.Label Lbl 
            Alignment       =   2  'Center
            BackColor       =   &H80000002&
            Caption         =   "Fecha"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000009&
            Height          =   285
            Index           =   13
            Left            =   5400
            TabIndex        =   11
            Top             =   1410
            Width           =   1335
         End
         Begin VB.Label Lbl 
            Alignment       =   2  'Center
            BackColor       =   &H80000002&
            Caption         =   "Banco"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000009&
            Height          =   285
            Index           =   12
            Left            =   3585
            TabIndex        =   10
            Top             =   1410
            Width           =   1815
         End
         Begin VB.Label Lbl 
            Alignment       =   2  'Center
            BackColor       =   &H80000002&
            Caption         =   "Documento"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000009&
            Height          =   285
            Index           =   11
            Left            =   1755
            TabIndex        =   9
            Top             =   1410
            Width           =   1830
         End
      End
   End
End
Attribute VB_Name = "frmCajaMultiple"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Const INMUEBLE_CODIGO& = 0
Private Const INMUEBLE_NOMBRE& = 1
Private Const APARTAMENTO_CODIGO& = 2
Private Const APARTAMENTO_NOMBRE& = 3

Dim rst(3) As ADODB.Recordset
Dim StrRutaInmueble As String

Private Sub cmd_Click(Index As Integer)
Select Case Index
    Case 0  'salir
        Unload Me
        Set frmCajaMultiple = Nothing
End Select
End Sub

Private Sub Dat_Click(Index As Integer, Area As Integer)
If Area = 2 Then
    Select Case Index
        Case 0
            Call RtnBuscaInmueble("Inmueble.CodInm", Dat(0))
            
            
        Case 1
            Call RtnBuscaInmueble("Inmueble.Nombre", Dat(1))
        
        Case 2, 3
            Dat(IIf(Index = 2, 3, 2)).Text = Dat(Index).BoundText
    End Select
End If
End Sub

Private Sub Form_Load()
Dim sql As String


sql = "SELECT * FROM Inmueble WHERE Inactivo=False ORDER BY CodInm"
Set rst(INMUEBLE_CODIGO) = cnnConexion.Execute(sql)

sql = "SELECT * FROM Inmueble WHERE Inactivo=False ORDER BY Nombre"
Set rst(INMUEBLE_NOMBRE) = cnnConexion.Execute(sql)
Set rst(APARTAMENTO_CODIGO) = New ADODB.Recordset
Set rst(APARTAMENTO_NOMBRE) = New ADODB.Recordset

Set Dat(0).RowSource = rst(INMUEBLE_CODIGO)
Set Dat(1).RowSource = rst(INMUEBLE_NOMBRE)



Lbl(0) = LoadResString(101)
Lbl(3) = LoadResString(102)

End Sub

Private Sub Form_Resize()
On Error Resume Next
If (Me.WindowState = vbMaximized) Then
    SSTab1.Width = Me.ScaleWidth - (SSTab1.Left * 2)
    SSTab1.Height = Me.ScaleHeight - (SSTab1.Top * 2)
End If
End Sub

Sub RtnBuscaInmueble(Qdf$, DC As DataCombo)
Dim sql As String

Call BuscaInmueble(Qdf, DC)

Set objRst = ObjCmd.Execute
    'Call rtnLimpiar_Grid(FlexFacturas)
    With objRst
        If objRst.EOF Then
            For i = 1 To 3: Dat(i) = ""
            Next
            MsgBox "Inmueble No Registrado..", vbInformation, App.ProductName
            Dat(0).SetFocus
            Exit Sub
        End If
        Dat(0) = .Fields("CodInm")
        Dat(1) = .Fields("Nombre")
        StrRutaInmueble = gcPath & .Fields("Ubica") & "inm.mdb"
        strCodPC = .Fields("CodPagocondominio")
        sql = "SELECT * FROM Propietarios in '" & gcPath & .Fields("Ubica") & "inm.mdb" & "' ORDER BY Codigo"
        
        Set rst(APARTAMENTO_CODIGO) = Nothing
        Set rst(APARTAMENTO_NOMBRE) = Nothing
        
        Set rst(APARTAMENTO_CODIGO) = cnnConexion.Execute(sql)
        sql = "SELECT * FROM Propietarios in '" & gcPath & .Fields("Ubica") & "inm.mdb" & "' ORDER BY Nombre"
        Set rst(APARTAMENTO_NOMBRE) = cnnConexion.Execute(sql)
        
        Set Dat(2).RowSource = rst(APARTAMENTO_CODIGO)
        Set Dat(3).RowSource = rst(APARTAMENTO_NOMBRE)
        '*********************************************
        Dat(2) = ""
        Dat(3) = ""
        Dat(2).Enabled = True
        Dat(3).Enabled = True
        Dat(2).SetFocus
        '
    End With
    '
    Set objRst = Nothing
    End Sub

    
