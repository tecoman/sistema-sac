VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form FrmAnulacheques 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Anulación de Cheques"
   ClientHeight    =   30
   ClientLeft      =   45
   ClientTop       =   405
   ClientWidth     =   8685
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   30
   ScaleWidth      =   8685
   ShowInTaskbar   =   0   'False
   Begin ComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   480
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8685
      _ExtentX        =   15319
      _ExtentY        =   847
      ButtonWidth     =   714
      ButtonHeight    =   688
      Appearance      =   1
      ImageList       =   "ImageList1"
      _Version        =   327682
      BeginProperty Buttons {0713E452-850A-101B-AFC0-4210102A8DA7} 
         NumButtons      =   12
         BeginProperty Button1 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "First"
            Object.ToolTipText     =   "Primer Registro"
            Object.Tag             =   ""
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Previous"
            Object.ToolTipText     =   "Registro Anterior"
            Object.Tag             =   ""
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Next"
            Object.ToolTipText     =   "Siguiente Registro"
            Object.Tag             =   ""
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "End"
            Object.ToolTipText     =   "Último Registro"
            Object.Tag             =   ""
            ImageIndex      =   4
         EndProperty
         BeginProperty Button5 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "New"
            Object.ToolTipText     =   "Nuevo Registro"
            Object.Tag             =   ""
            ImageIndex      =   5
         EndProperty
         BeginProperty Button6 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Save"
            Object.ToolTipText     =   "Guardar Registro"
            Object.Tag             =   ""
            ImageIndex      =   6
         EndProperty
         BeginProperty Button7 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Find"
            Object.ToolTipText     =   "Buscar Registro"
            Object.Tag             =   ""
            ImageIndex      =   7
         EndProperty
         BeginProperty Button8 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Undo"
            Object.ToolTipText     =   "Cancelar Registro"
            Object.Tag             =   ""
            ImageIndex      =   8
         EndProperty
         BeginProperty Button9 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Delete"
            Object.ToolTipText     =   "Eliminar Registro"
            Object.Tag             =   ""
            ImageIndex      =   9
         EndProperty
         BeginProperty Button10 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Edit1"
            Object.ToolTipText     =   "Editar Registro"
            Object.Tag             =   ""
            ImageIndex      =   10
         EndProperty
         BeginProperty Button11 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Print"
            Object.ToolTipText     =   "Imprimir"
            Object.Tag             =   ""
            ImageIndex      =   11
         EndProperty
         BeginProperty Button12 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Close"
            Object.ToolTipText     =   "Salir"
            Object.Tag             =   ""
            ImageIndex      =   12
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc AdoAnulado 
      Height          =   330
      Left            =   60
      Top             =   4680
      Visible         =   0   'False
      Width           =   2325
      _ExtentX        =   4101
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   1
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
      DataSourceName  =   "sac"
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "AdoAnulado"
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
   Begin TabDlg.SSTab SSTab1 
      Height          =   4245
      Left            =   60
      TabIndex        =   1
      Top             =   420
      Width           =   9615
      _ExtentX        =   16960
      _ExtentY        =   7488
      _Version        =   393216
      Tabs            =   2
      Tab             =   1
      TabsPerRow      =   4
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Datos Generales"
      TabPicture(0)   =   "FrmAnulacheques.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "FraAnulacion"
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Lista"
      TabPicture(1)   =   "FrmAnulacheques.frx":001C
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "DataGrid1"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      Begin VB.Frame FraAnulacion 
         Enabled         =   0   'False
         Height          =   3225
         Left            =   -74910
         TabIndex        =   2
         Top             =   330
         Width           =   9435
         Begin VB.TextBox TxtCodChequera 
            DataField       =   "CodAnulacion"
            DataSource      =   "AdoAnulado"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1770
            MaxLength       =   5
            TabIndex        =   7
            Top             =   480
            Width           =   1305
         End
         Begin VB.TextBox TxtNumCuenta 
            BackColor       =   &H00C0C0C0&
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   4590
            Locked          =   -1  'True
            MaxLength       =   15
            TabIndex        =   6
            Top             =   960
            Width           =   4725
         End
         Begin VB.TextBox TxtDesde 
            DataField       =   "DesdeCheque"
            DataSource      =   "AdoAnulado"
            Height          =   315
            Left            =   1770
            MaxLength       =   8
            TabIndex        =   5
            Text            =   " "
            Top             =   1950
            Width           =   1305
         End
         Begin VB.TextBox TxtHasta 
            DataField       =   "HastaCheque"
            DataSource      =   "AdoAnulado"
            Height          =   315
            Left            =   4650
            MaxLength       =   8
            TabIndex        =   4
            Text            =   " "
            Top             =   1950
            Width           =   1125
         End
         Begin VB.CheckBox ChkAnulaCompleta 
            Alignment       =   1  'Right Justify
            Caption         =   "Anular Chequera Completa :"
            DataField       =   "AnulaCompleta"
            DataSource      =   "AdoAnulado"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   300
            TabIndex        =   3
            Top             =   2640
            Width           =   3015
         End
         Begin MSMask.MaskEdBox MskFecha 
            Bindings        =   "FrmAnulacheques.frx":0038
            DataField       =   "FechaAnulacion"
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "dd/MM/yyyy"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   8202
               SubFormatType   =   3
            EndProperty
            DataSource      =   "AdoAnulado"
            Height          =   315
            Left            =   1770
            TabIndex        =   8
            TabStop         =   0   'False
            Top             =   1440
            Width           =   1305
            _ExtentX        =   2302
            _ExtentY        =   556
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   10
            Format          =   "dd/MM/yyyy"
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin MSDataListLib.DataCombo DtcCodChequera 
            Bindings        =   "FrmAnulacheques.frx":005A
            DataField       =   "CodChequera"
            DataSource      =   "AdoAnulado"
            Height          =   330
            Left            =   1770
            TabIndex        =   16
            Top             =   960
            Width           =   1305
            _ExtentX        =   2302
            _ExtentY        =   582
            _Version        =   393216
            BackColor       =   -2147483643
            ListField       =   "CodChequera"
            BoundColumn     =   "Nombre"
            Text            =   ""
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "Fecha Registro:"
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
            Left            =   300
            TabIndex        =   14
            Top             =   1530
            Width           =   1260
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Cod. Anulación:"
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
            Left            =   300
            TabIndex        =   13
            Top             =   510
            Width           =   1275
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "Hasta Cheque :"
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
            Left            =   3240
            TabIndex        =   12
            Top             =   2040
            Width           =   1260
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            Caption         =   "Num. Cuenta :"
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
            Left            =   3300
            TabIndex        =   11
            Top             =   1020
            Width           =   1200
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            Caption         =   "Desde Cheque :"
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
            Left            =   300
            TabIndex        =   10
            Top             =   2040
            Width           =   1320
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Cod. Chequera :"
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
            Left            =   300
            TabIndex        =   9
            Top             =   1020
            Width           =   1335
         End
      End
      Begin MSDataGridLib.DataGrid DataGrid1 
         Bindings        =   "FrmAnulacheques.frx":0073
         Height          =   3390
         Left            =   150
         TabIndex        =   15
         Top             =   510
         Width           =   8790
         _ExtentX        =   15505
         _ExtentY        =   5980
         _Version        =   393216
         AllowUpdate     =   0   'False
         BackColor       =   16777215
         Enabled         =   -1  'True
         HeadLines       =   1
         RowHeight       =   16
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
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   6
         BeginProperty Column00 
            DataField       =   "CodChequera"
            Caption         =   "CodChequera"
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
         BeginProperty Column01 
            DataField       =   "DesdeCheque"
            Caption         =   "DesdeCheque"
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
         BeginProperty Column02 
            DataField       =   "HastaCheque"
            Caption         =   "HastaCheque"
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
         BeginProperty Column03 
            DataField       =   "CodAnulacion"
            Caption         =   "CodAnulacion"
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
         BeginProperty Column04 
            DataField       =   "FechaAnulacion"
            Caption         =   "FechaAnulacion"
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
         BeginProperty Column05 
            DataField       =   "AnulaCompleta"
            Caption         =   "AnulaCompleta"
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
               ColumnWidth     =   1094,74
            EndProperty
            BeginProperty Column01 
            EndProperty
            BeginProperty Column02 
            EndProperty
            BeginProperty Column03 
               ColumnWidth     =   1289,764
            EndProperty
            BeginProperty Column04 
               ColumnWidth     =   1530,142
            EndProperty
            BeginProperty Column05 
               ColumnWidth     =   1319,811
            EndProperty
         EndProperty
      End
   End
   Begin MSAdodcLib.Adodc AdoCheques 
      Height          =   360
      Left            =   2460
      Top             =   4680
      Visible         =   0   'False
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   635
      ConnectMode     =   3
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   1
      LockType        =   3
      CommandType     =   1
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
      Caption         =   "AdoCheques"
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
            Picture         =   "FrmAnulacheques.frx":008C
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmAnulacheques.frx":020E
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmAnulacheques.frx":0390
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmAnulacheques.frx":0512
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmAnulacheques.frx":0694
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmAnulacheques.frx":0816
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmAnulacheques.frx":0998
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmAnulacheques.frx":0B1A
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmAnulacheques.frx":0C9C
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmAnulacheques.frx":0E1E
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmAnulacheques.frx":0FA0
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmAnulacheques.frx":1122
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "FrmAnulacheques"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Private Sub DtcCodChequera_Change()
'    Dim StrBanco As String
'    Dim Cnx As ADODB.Connection
'    Dim Cnx2 As ADODB.Connection
'    Dim AdoBancos As ADODB.Recordset
'    Dim strSql As String
'
'    Set Cnx = New Connection
'    Set Cnx2 = New Connection
'    Set AdoBancos = New ADODB.Recordset
'
'    Cnx.Open "DSN=Sac"
'    Cnx2.Open "Driver=Microsoft Access Driver (*.MDB);DBQ=" + gcPath + gcUbica + "inm.mdb"
'    Set Rst = New ADODB.Recordset
'    StrBanco = "SELECT * " & _
'                "FROM CHEQUERA " & _
'                "WHERE CodChequera = '" & DtcCodChequera.Text & "' "
'    Rst.Source = StrBanco
'    Rst.ActiveConnection = Cnx
'    Rst.Open
'
'    If Not Rst.EOF And Not Rst.BOF Then
'        strSql = "SELECT * " & _
'                "FROM BANCOS " & _
'                "WHERE IDBanco = " & Rst("IDbanco")
'        AdoBancos.Open strSql, Cnx2, adOpenKeyset, adLockOptimistic
'
'        'TxtBanco = AdoBancos("nombrebanco")
'        TxtNumCuenta = Rst("codcuenta")
'    End If
'
'End Sub
'Private Sub Toolbar1_ButtonClick(ByVal Button As ComctlLib.Button)
'With AdoAnulado.Recordset
'    Select Case Button.Index
'
'        Case 1
'            'Primer Registro
'            Beep
'            .MoveFirst
'        Case 2
'            'Registro Previo
'            .MovePrevious
'            If .BOF Then
'                Beep
'                .MoveFirst
'            End If
'        Case 3
'            'Siguiente Registro
'            .MoveNext
'            If .EOF Then
'                Beep
'                .MoveLast
'            End If
'        Case 4
'            'Último Registro
'            Beep
'            .MoveLast
'        Case 5
'            On Error GoTo reingreso
'            'Nuevo Registro
'            If Not .EOF And Not .BOF Then
'                If .EditMode <> adEditAdd Then
'                    .AddNew
'                    'DataGrid1.AllowAddNew = True
'                    SSTab1.Tab = 0
'                    FraAnulacion.Enabled = True
'                   ' EditarCodigo
'                End If
'            Else
'                    .AddNew
'                    SSTab1.Tab = 0
'                    'DataGrid1.AllowAddNew = True
'                    FraAnulacion.Enabled = True
'            End If
'reingreso:
'            If Err.Number = -2147217842 Then
'
'            End If
'        Case 6
'            On Error GoTo validaentrada
'            'Actualizar Registro
'            MskFecha.PromptInclude = True
'            .Update
'            MsgBox "Registro actualizado..."
'            MskFecha.PromptInclude = False
'
'            'NoEditarCodigo
'
'validaentrada:
'            If Err.Number = -2147217900 Then
'                If .EditMode = adEditAdd Then
'                    SSTab1.Tab = 0
'                End If
'                Exit Sub
'            End If
'
'        Case 7
'             'Buscar Registro
'
'        Case 8
'            On Error GoTo cancelarerror
'            'Cancelar Registro
'                'DataGrid1.AllowAddNew = False
'                .CancelUpdate
'                FraAnulacion.Enabled = False
'
'cancelarerror:
'
'
'        Case 9
'            'Eliminar Registro
'            Dim Confirma As Integer
'            Confirma = MsgBox("Confirma eliminar el registro actual ?", vbOKCancel, "Eliminar Registro")
'            If Confirma = vbOK Then
'                .Delete
'                .MoveNext
'                If .EOF And Not .BOF Then
'                    .MoveLast
'                End If
'            End If
'        Case 12
'            'Descargar Formulario
'            Unload Me
'        Case 11
'            'Imprimir Registro
'
'    End Select
' End With
'
'End Sub
Private Sub Form_Load()

End Sub
