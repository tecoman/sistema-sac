VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form FrmFactura 
   AutoRedraw      =   -1  'True
   Caption         =   "Recepción de Facturas"
   ClientHeight    =   8145
   ClientLeft      =   60
   ClientTop       =   420
   ClientWidth     =   8880
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   8145
   ScaleWidth      =   8880
   WindowState     =   2  'Maximized
   Begin ComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   480
      Left            =   0
      TabIndex        =   20
      Top             =   0
      Width           =   8880
      _ExtentX        =   15663
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
      BorderStyle     =   1
   End
   Begin MSAdodcLib.Adodc Ado 
      Height          =   330
      Index           =   0
      Left            =   990
      Top             =   5385
      Visible         =   0   'False
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   582
      ConnectMode     =   16
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
      Caption         =   "Cpp"
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
   Begin VB.Frame FraFactura 
      Enabled         =   0   'False
      Height          =   2055
      Index           =   0
      Left            =   330
      TabIndex        =   65
      Top             =   1005
      Width           =   9045
      Begin VB.TextBox TxtFactura 
         DataField       =   "Detalle"
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
         Index           =   1
         Left            =   1425
         MaxLength       =   200
         TabIndex        =   13
         Text            =   " "
         Top             =   1635
         Width           =   7515
      End
      Begin VB.TextBox TxtNdoc 
         DataField       =   "Ndoc"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   315
         Index           =   0
         Left            =   1425
         Locked          =   -1  'True
         MaxLength       =   7
         TabIndex        =   7
         Top             =   945
         Width           =   1425
      End
      Begin VB.ComboBox CmbTipoPago 
         DataField       =   "Tipo"
         Height          =   315
         Index           =   0
         ItemData        =   "FrmFactura.frx":0000
         Left            =   6855
         List            =   "FrmFactura.frx":0013
         TabIndex        =   9
         Top             =   945
         Width           =   630
      End
      Begin VB.TextBox TxtFactura 
         DataField       =   "Fact"
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
         Index           =   0
         Left            =   1425
         MaxLength       =   7
         TabIndex        =   11
         Top             =   1290
         Width           =   1425
      End
      Begin VB.TextBox TxtDescripTipoDoc 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   315
         Index           =   0
         Left            =   7515
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   66
         Top             =   945
         Width           =   1425
      End
      Begin MSDataListLib.DataCombo DatInmueble 
         CausesValidation=   0   'False
         Height          =   330
         Index           =   1
         Left            =   2895
         TabIndex        =   2
         Top             =   210
         Width           =   6045
         _ExtentX        =   10663
         _ExtentY        =   582
         _Version        =   393216
         BackColor       =   -2147483643
         ListField       =   "Nombre"
         BoundColumn     =   "Codigo"
         Text            =   ""
         Object.DataMember      =   ""
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
      Begin MSDataListLib.DataCombo DatProveedor 
         Bindings        =   "FrmFactura.frx":002B
         Height          =   330
         Index           =   1
         Left            =   2895
         TabIndex        =   5
         Top             =   570
         Width           =   6045
         _ExtentX        =   10663
         _ExtentY        =   582
         _Version        =   393216
         BackColor       =   -2147483643
         ListField       =   "NombProv"
         BoundColumn     =   ""
         Text            =   ""
         Object.DataMember      =   ""
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
      Begin MSDataListLib.DataCombo DatInmueble 
         Bindings        =   "FrmFactura.frx":0036
         DataField       =   "CodInm"
         DataSource      =   "Ado(0)"
         Height          =   330
         Index           =   0
         Left            =   1425
         TabIndex        =   1
         Top             =   210
         Width           =   1425
         _ExtentX        =   2514
         _ExtentY        =   582
         _Version        =   393216
         BackColor       =   -2147483643
         ListField       =   "CodInm"
         BoundColumn     =   ""
         Text            =   ""
         Object.DataMember      =   ""
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
      Begin MSDataListLib.DataCombo DatProveedor 
         Bindings        =   "FrmFactura.frx":0041
         DataField       =   "CodProv"
         Height          =   330
         Index           =   0
         Left            =   1425
         TabIndex        =   4
         Top             =   570
         Width           =   1425
         _ExtentX        =   2514
         _ExtentY        =   582
         _Version        =   393216
         BackColor       =   -2147483643
         ListField       =   "Codigo"
         BoundColumn     =   ""
         Text            =   ""
         Object.DataMember      =   ""
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
      Begin VB.Label LblFactura 
         Alignment       =   1  'Right Justify
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
         Left            =   100
         TabIndex        =   0
         Top             =   300
         Width           =   1250
      End
      Begin VB.Label LblFactura 
         Alignment       =   1  'Right Justify
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
         Index           =   4
         Left            =   100
         TabIndex        =   12
         Top             =   1650
         Width           =   1250
      End
      Begin VB.Label LblFactura 
         Alignment       =   1  'Right Justify
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
         Index           =   1
         Left            =   100
         TabIndex        =   3
         Top             =   660
         Width           =   1250
      End
      Begin VB.Label LblFactura 
         Alignment       =   1  'Right Justify
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
         Index           =   2
         Left            =   100
         TabIndex        =   6
         Top             =   975
         Width           =   1250
      End
      Begin VB.Label LblFactura 
         AutoSize        =   -1  'True
         Caption         =   "Tipo de Documento :"
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
         Index           =   5
         Left            =   5040
         TabIndex        =   8
         Top             =   975
         Width           =   1770
      End
      Begin VB.Label LblFactura 
         Alignment       =   1  'Right Justify
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
         Left            =   100
         TabIndex        =   10
         Top             =   1320
         Width           =   1250
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   6330
      Left            =   45
      TabIndex        =   34
      Top             =   525
      Width           =   9615
      _ExtentX        =   16960
      _ExtentY        =   11165
      _Version        =   393216
      Tabs            =   4
      TabsPerRow      =   5
      TabHeight       =   520
      ShowFocusRect   =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Datos &Generales"
      TabPicture(0)   =   "FrmFactura.frx":0063
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "FraFactura(1)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Datos &Adicionales"
      TabPicture(1)   =   "FrmFactura.frx":007F
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "ImageList1"
      Tab(1).Control(1)=   "FraFactura(3)"
      Tab(1).ControlCount=   2
      TabCaption(2)   =   "&Lista"
      TabPicture(2)   =   "FrmFactura.frx":009B
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "DataGrid1"
      Tab(2).Control(1)=   "FrmBusca1(0)"
      Tab(2).Control(2)=   "FrmBusca"
      Tab(2).ControlCount=   3
      TabCaption(3)   =   "&Presupuestos"
      TabPicture(3)   =   "FrmFactura.frx":00B7
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "FrmBusca1(3)"
      Tab(3).Control(1)=   "FrmBusca1(2)"
      Tab(3).Control(2)=   "FrmBusca1(1)"
      Tab(3).Control(3)=   "LblFactura(19)"
      Tab(3).Control(4)=   "LblFactura(18)"
      Tab(3).Control(5)=   "LblFactura(17)"
      Tab(3).Control(6)=   "LblFactura(16)"
      Tab(3).ControlCount=   7
      Begin VB.Frame FrmBusca1 
         Caption         =   "Facturas-Presupuesto"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2445
         Index           =   3
         Left            =   -70365
         TabIndex        =   76
         Top             =   390
         Width           =   4695
         Begin MSDataGridLib.DataGrid DtgPresupuesto 
            Height          =   1755
            Index           =   2
            Left            =   315
            TabIndex        =   77
            Top             =   450
            Width           =   4080
            _ExtentX        =   7197
            _ExtentY        =   3096
            _Version        =   393216
            AllowUpdate     =   0   'False
            AllowArrows     =   -1  'True
            BorderStyle     =   0
            Enabled         =   -1  'True
            ColumnHeaders   =   -1  'True
            ForeColor       =   -2147483646
            HeadLines       =   1
            RowHeight       =   15
            WrapCellPointer =   -1  'True
            RowDividerStyle =   4
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
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ColumnCount     =   3
            BeginProperty Column00 
               DataField       =   "Fact"
               Caption         =   "Factura"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   1
                  Format          =   "0000000"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   8202
                  SubFormatType   =   0
               EndProperty
            EndProperty
            BeginProperty Column01 
               DataField       =   "Total"
               Caption         =   "Total"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   1
                  Format          =   """Bs"" #,##0.00"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   8202
                  SubFormatType   =   0
               EndProperty
            EndProperty
            BeginProperty Column02 
               DataField       =   "Freg"
               Caption         =   "Recepción"
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
               MarqueeStyle    =   4
               ScrollBars      =   2
               AllowRowSizing  =   0   'False
               AllowSizing     =   0   'False
               BeginProperty Column00 
                  Alignment       =   2
                  ColumnAllowSizing=   -1  'True
                  ColumnWidth     =   959,811
               EndProperty
               BeginProperty Column01 
                  ColumnWidth     =   1260,284
               EndProperty
               BeginProperty Column02 
                  ColumnWidth     =   1049,953
               EndProperty
            EndProperty
         End
      End
      Begin VB.Frame FrmBusca1 
         Caption         =   "Presupuestos Proveedor/Inmueble:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2430
         Index           =   2
         Left            =   -74715
         TabIndex        =   69
         Top             =   390
         Width           =   3945
         Begin MSDataGridLib.DataGrid DtgPresupuesto 
            Height          =   1695
            Index           =   1
            Left            =   300
            TabIndex        =   71
            Top             =   435
            Width           =   3360
            _ExtentX        =   5927
            _ExtentY        =   2990
            _Version        =   393216
            AllowUpdate     =   0   'False
            AllowArrows     =   -1  'True
            BorderStyle     =   0
            Enabled         =   -1  'True
            ColumnHeaders   =   -1  'True
            ForeColor       =   -2147483646
            HeadLines       =   1
            RowHeight       =   15
            WrapCellPointer =   -1  'True
            RowDividerStyle =   4
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
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ColumnCount     =   3
            BeginProperty Column00 
               DataField       =   "Fact"
               Caption         =   "Presupuesto"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   1
                  Format          =   "0000000"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   8202
                  SubFormatType   =   0
               EndProperty
            EndProperty
            BeginProperty Column01 
               DataField       =   "CodInm"
               Caption         =   "Inmueble"
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
               DataField       =   "CodProv"
               Caption         =   "Proveedor"
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
               MarqueeStyle    =   4
               ScrollBars      =   2
               AllowRowSizing  =   0   'False
               AllowSizing     =   0   'False
               BeginProperty Column00 
                  Alignment       =   2
                  ColumnAllowSizing=   -1  'True
                  ColumnWidth     =   959,811
               EndProperty
               BeginProperty Column01 
                  Alignment       =   2
                  ColumnWidth     =   810,142
               EndProperty
               BeginProperty Column02 
                  ColumnWidth     =   810,142
               EndProperty
            EndProperty
         End
      End
      Begin VB.Frame FrmBusca1 
         Caption         =   "Facturas Recibidas:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2055
         Index           =   1
         Left            =   -74750
         TabIndex        =   68
         Top             =   3810
         Width           =   9135
         Begin MSDataGridLib.DataGrid DtgPresupuesto 
            Height          =   1590
            Index           =   0
            Left            =   120
            TabIndex        =   70
            Top             =   330
            Width           =   8880
            _ExtentX        =   15663
            _ExtentY        =   2805
            _Version        =   393216
            AllowUpdate     =   0   'False
            AllowArrows     =   -1  'True
            BorderStyle     =   0
            Enabled         =   -1  'True
            ColumnHeaders   =   -1  'True
            ForeColor       =   -2147483646
            HeadLines       =   1
            RowHeight       =   17
            WrapCellPointer =   -1  'True
            RowDividerStyle =   6
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
            ColumnCount     =   5
            BeginProperty Column00 
               DataField       =   "Fact"
               Caption         =   "Factura"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   1
                  Format          =   "0000000"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   8202
                  SubFormatType   =   0
               EndProperty
            EndProperty
            BeginProperty Column01 
               DataField       =   "CodInm"
               Caption         =   "Inmueble"
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
               DataField       =   "CodProv"
               Caption         =   "Proveedor"
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
               DataField       =   "Detalle"
               Caption         =   "Descripción"
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
               DataField       =   "Total"
               Caption         =   "Total"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   1
                  Format          =   "#,##0.00 "
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   8202
                  SubFormatType   =   0
               EndProperty
            EndProperty
            SplitCount      =   1
            BeginProperty Split0 
               MarqueeStyle    =   4
               ScrollBars      =   2
               AllowRowSizing  =   0   'False
               AllowSizing     =   0   'False
               BeginProperty Column00 
                  Alignment       =   2
                  ColumnAllowSizing=   -1  'True
                  ColumnWidth     =   989,858
               EndProperty
               BeginProperty Column01 
                  Alignment       =   2
                  ColumnWidth     =   810,142
               EndProperty
               BeginProperty Column02 
                  ColumnWidth     =   824,882
               EndProperty
               BeginProperty Column03 
                  ColumnWidth     =   4155,024
               EndProperty
               BeginProperty Column04 
                  Alignment       =   1
                  ColumnWidth     =   1454,74
               EndProperty
            EndProperty
         End
      End
      Begin MSDataGridLib.DataGrid DataGrid1 
         Height          =   4140
         Left            =   -74873
         TabIndex        =   67
         Top             =   375
         Width           =   9390
         _ExtentX        =   16563
         _ExtentY        =   7303
         _Version        =   393216
         AllowUpdate     =   0   'False
         AllowArrows     =   -1  'True
         BorderStyle     =   0
         Enabled         =   -1  'True
         ColumnHeaders   =   -1  'True
         ForeColor       =   -2147483646
         HeadLines       =   1
         RowHeight       =   17
         WrapCellPointer =   -1  'True
         RowDividerStyle =   4
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
         Caption         =   "LISTA DE FACTURAS PENDIENTES"
         ColumnCount     =   5
         BeginProperty Column00 
            DataField       =   "Fact"
            Caption         =   "Factura"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   "0000000"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   8202
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   "CodInm"
            Caption         =   "Inmueble"
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
            DataField       =   "FRecep"
            Caption         =   "Recepción"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   "dd/MM/yyyy"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   8202
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column03 
            DataField       =   "Name"
            Caption         =   "Proveedor"
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
            DataField       =   "Total"
            Caption         =   "Total"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   "#,##0.00 "
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   8202
               SubFormatType   =   0
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            MarqueeStyle    =   4
            ScrollBars      =   2
            AllowRowSizing  =   0   'False
            AllowSizing     =   0   'False
            BeginProperty Column00 
               Alignment       =   2
               ColumnAllowSizing=   -1  'True
               ColumnWidth     =   1110,047
            EndProperty
            BeginProperty Column01 
               Alignment       =   2
               ColumnWidth     =   810,142
            EndProperty
            BeginProperty Column02 
               Alignment       =   2
               ColumnWidth     =   1065,26
            EndProperty
            BeginProperty Column03 
               ColumnWidth     =   3990,047
            EndProperty
            BeginProperty Column04 
               Alignment       =   1
               ColumnWidth     =   1739,906
            EndProperty
         EndProperty
      End
      Begin VB.Frame FraFactura 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   3405
         Index           =   1
         Left            =   285
         TabIndex        =   58
         Top             =   2610
         Width           =   9075
         Begin VB.Frame Frame1 
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   930
            Left            =   3225
            TabIndex        =   78
            Top             =   1905
            Width           =   2430
            Begin MSMask.MaskEdBox MaskEdBox4 
               Bindings        =   "FrmFactura.frx":00D3
               BeginProperty DataFormat 
                  Type            =   1
                  Format          =   "#,##0.00"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   8202
                  SubFormatType   =   1
               EndProperty
               Height          =   330
               Left            =   1215
               TabIndex        =   79
               Top             =   195
               Width           =   795
               _ExtentX        =   1402
               _ExtentY        =   582
               _Version        =   393216
               PromptInclude   =   0   'False
               MaxLength       =   5
               Format          =   "#,##0.00"
               PromptChar      =   "_"
            End
            Begin VB.Label LblFactura 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Tasa IVA :"
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
               Index           =   11
               Left            =   300
               TabIndex        =   80
               Top             =   270
               Width           =   855
            End
         End
         Begin VB.TextBox TxtMonto 
            Alignment       =   1  'Right Justify
            DataField       =   "Total"
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "#,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   8202
               SubFormatType   =   0
            EndProperty
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   1
            Left            =   7335
            TabIndex        =   33
            Top             =   2760
            Width           =   1590
         End
         Begin VB.TextBox TxtMonto 
            Alignment       =   1  'Right Justify
            DataField       =   "Monto"
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "#,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   8202
               SubFormatType   =   0
            EndProperty
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Index           =   0
            Left            =   7335
            TabIndex        =   30
            Top             =   1515
            Width           =   1590
         End
         Begin VB.Frame FraFactura 
            Caption         =   "Fechas"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1635
            Index           =   2
            Left            =   240
            TabIndex        =   59
            Top             =   150
            Width           =   2880
            Begin VB.CommandButton CmdFecha 
               Height          =   270
               Index           =   1
               Left            =   2270
               Picture         =   "FrmFactura.frx":00DE
               Style           =   1  'Graphical
               TabIndex        =   19
               Top             =   795
               Width           =   235
            End
            Begin VB.CommandButton CmdFecha 
               Height          =   270
               Index           =   2
               Left            =   2270
               Picture         =   "FrmFactura.frx":0228
               Style           =   1  'Graphical
               TabIndex        =   22
               Top             =   1200
               Width           =   235
            End
            Begin VB.CommandButton CmdFecha 
               Height          =   270
               Index           =   0
               Left            =   2270
               Picture         =   "FrmFactura.frx":0372
               Style           =   1  'Graphical
               TabIndex        =   16
               Top             =   390
               Width           =   235
            End
            Begin MSMask.MaskEdBox MskFecha 
               Bindings        =   "FrmFactura.frx":04BC
               DataField       =   "Fven"
               Height          =   315
               Index           =   2
               Left            =   1125
               TabIndex        =   21
               TabStop         =   0   'False
               Top             =   1170
               Width           =   1410
               _ExtentX        =   2487
               _ExtentY        =   556
               _Version        =   393216
               PromptInclude   =   0   'False
               AllowPrompt     =   -1  'True
               MaxLength       =   10
               Format          =   "dd/MM/yyyy"
               Mask            =   "##/##/####"
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox MskFecha 
               Bindings        =   "FrmFactura.frx":04DE
               DataField       =   "FRecep"
               Height          =   315
               Index           =   0
               Left            =   1125
               TabIndex        =   15
               TabStop         =   0   'False
               Top             =   360
               Width           =   1410
               _ExtentX        =   2487
               _ExtentY        =   556
               _Version        =   393216
               PromptInclude   =   0   'False
               MaxLength       =   12
               Format          =   "dd/MM/yyyy"
               Mask            =   "##/##/####"
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox MskFecha 
               Bindings        =   "FrmFactura.frx":0500
               DataField       =   "Fecr"
               Height          =   315
               Index           =   1
               Left            =   1125
               TabIndex        =   18
               TabStop         =   0   'False
               Top             =   765
               Width           =   1410
               _ExtentX        =   2487
               _ExtentY        =   556
               _Version        =   393216
               PromptInclude   =   0   'False
               MaxLength       =   10
               Format          =   "dd/mm/yyyy"
               Mask            =   "##/##/####"
               PromptChar      =   "_"
            End
            Begin VB.Label LblFactura 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "&Emision :"
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
               Index           =   7
               Left            =   50
               TabIndex        =   17
               Top             =   795
               Width           =   1000
            End
            Begin VB.Label LblFactura 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "&Vence :"
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
               Index           =   8
               Left            =   50
               TabIndex        =   60
               Top             =   1200
               Width           =   1000
            End
            Begin VB.Label LblFactura 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "&Recepción :"
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
               Index           =   6
               Left            =   50
               TabIndex        =   14
               Top             =   390
               Width           =   1000
            End
         End
         Begin VB.CheckBox Check1 
            Caption         =   "&Aplicar IVA"
            CausesValidation=   0   'False
            Height          =   240
            Left            =   3570
            TabIndex        =   24
            Top             =   1575
            Value           =   1  'Checked
            Width           =   1125
         End
         Begin VB.TextBox TxtFactura 
            DataField       =   "Benef"
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
            Index           =   4
            Left            =   4650
            TabIndex        =   23
            Text            =   " "
            Top             =   930
            Width           =   4305
         End
         Begin MSDataListLib.DataCombo CmbCargo 
            Bindings        =   "FrmFactura.frx":0522
            DataField       =   "Moneda"
            Height          =   330
            Left            =   5955
            TabIndex        =   26
            Top             =   450
            Width           =   825
            _ExtentX        =   1455
            _ExtentY        =   582
            _Version        =   393216
            BackColor       =   -2147483643
            ListField       =   "IdMoneda"
            BoundColumn     =   ""
            Text            =   "Bs."
            Object.DataMember      =   "CmdMoneda"
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
         Begin MSDataListLib.DataCombo DataCombo5 
            Bindings        =   "FrmFactura.frx":053E
            Height          =   330
            Left            =   6795
            TabIndex        =   61
            Top             =   450
            Width           =   2175
            _ExtentX        =   3836
            _ExtentY        =   582
            _Version        =   393216
            IntegralHeight  =   0   'False
            MatchEntry      =   -1  'True
            BackColor       =   -2147483643
            ListField       =   "Moneda"
            BoundColumn     =   ""
            Text            =   ""
            Object.DataMember      =   "CmdMoneda"
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
         Begin MSAdodcLib.Adodc Ado 
            Height          =   330
            Index           =   1
            Left            =   660
            Top             =   2610
            Visible         =   0   'False
            Width           =   2175
            _ExtentX        =   3836
            _ExtentY        =   582
            ConnectMode     =   0
            CursorLocation  =   3
            IsolationLevel  =   -1
            ConnectionTimeout=   15
            CommandTimeout  =   30
            CursorType      =   2
            LockType        =   3
            CommandType     =   2
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
            Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\sac\sac.mdb;Persist Security Info=False"
            OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\sac\sac.mdb;Persist Security Info=False"
            OLEDBFile       =   ""
            DataSourceName  =   ""
            OtherAttributes =   ""
            UserName        =   ""
            Password        =   ""
            RecordSource    =   "Cpp"
            Caption         =   "Inmuebles"
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
         Begin MSAdodcLib.Adodc Ado 
            Height          =   330
            Index           =   2
            Left            =   660
            Top             =   2940
            Visible         =   0   'False
            Width           =   2175
            _ExtentX        =   3836
            _ExtentY        =   582
            ConnectMode     =   0
            CursorLocation  =   3
            IsolationLevel  =   -1
            ConnectionTimeout=   15
            CommandTimeout  =   30
            CursorType      =   3
            LockType        =   3
            CommandType     =   2
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
            Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\sac\sac.mdb;Persist Security Info=False"
            OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\sac\sac.mdb;Persist Security Info=False"
            OLEDBFile       =   ""
            DataSourceName  =   ""
            OtherAttributes =   ""
            UserName        =   ""
            Password        =   ""
            RecordSource    =   "Cpp"
            Caption         =   "Inmuebles"
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
         Begin VB.Label LblFactura 
            Alignment       =   1  'Right Justify
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            DataField       =   "Ivm"
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "#,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   8202
               SubFormatType   =   0
            EndProperty
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Index           =   29
            Left            =   7335
            TabIndex        =   31
            Top             =   2040
            Width           =   1605
         End
         Begin VB.Label LblFactura 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Monto &Factura :"
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
            Index           =   12
            Left            =   5970
            TabIndex        =   28
            Top             =   1575
            Width           =   1305
         End
         Begin VB.Label LblFactura 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "&Moneda:"
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
            Index           =   9
            Left            =   5160
            TabIndex        =   25
            Top             =   510
            Width           =   705
         End
         Begin VB.Label LblFactura 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Monto Total :"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   14
            Left            =   5385
            TabIndex        =   32
            Top             =   2850
            Width           =   1650
         End
         Begin VB.Label LblFactura 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Monto Iva :"
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
            Index           =   13
            Left            =   6315
            TabIndex        =   29
            Top             =   2085
            Width           =   945
         End
         Begin VB.Label LblFactura 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "&Beneficiario:"
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
            Index           =   10
            Left            =   3555
            TabIndex        =   27
            Top             =   975
            Width           =   975
         End
      End
      Begin VB.Frame FraFactura 
         Enabled         =   0   'False
         Height          =   2385
         Index           =   3
         Left            =   -74760
         TabIndex        =   41
         Top             =   2640
         Width           =   9045
         Begin VB.TextBox TxtCampo1 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   1395
            TabIndex        =   49
            Top             =   285
            Width           =   2565
         End
         Begin VB.TextBox TxtCampo2 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   1425
            TabIndex        =   48
            Top             =   780
            Width           =   2565
         End
         Begin VB.TextBox TxtCampo3 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   1410
            TabIndex        =   47
            Top             =   1275
            Width           =   2565
         End
         Begin VB.TextBox TxtCampo4 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   1395
            TabIndex        =   46
            Top             =   1815
            Width           =   2565
         End
         Begin VB.TextBox TxtCampo5 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   6210
            TabIndex        =   45
            Top             =   285
            Width           =   2565
         End
         Begin VB.TextBox TxtCampo6 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   6195
            TabIndex        =   44
            Top             =   795
            Width           =   2565
         End
         Begin VB.TextBox TxtCampo7 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   6210
            TabIndex        =   43
            Top             =   1290
            Width           =   2565
         End
         Begin VB.TextBox TxtCampo8 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   6210
            TabIndex        =   42
            Top             =   1800
            Width           =   2565
         End
         Begin VB.Label LblFactura 
            Alignment       =   1  'Right Justify
            Caption         =   "Campo 6"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   26
            Left            =   5000
            TabIndex        =   57
            Top             =   855
            Width           =   900
         End
         Begin VB.Label LblFactura 
            Alignment       =   1  'Right Justify
            Caption         =   "Campo 1"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   21
            Left            =   200
            TabIndex        =   56
            Top             =   315
            Width           =   900
         End
         Begin VB.Label LblFactura 
            Alignment       =   1  'Right Justify
            Caption         =   "Campo 2"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   22
            Left            =   200
            TabIndex        =   55
            Top             =   810
            Width           =   900
         End
         Begin VB.Label LblFactura 
            Alignment       =   1  'Right Justify
            Caption         =   "Campo 3"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   23
            Left            =   200
            TabIndex        =   54
            Top             =   1290
            Width           =   900
         End
         Begin VB.Label LblFactura 
            Alignment       =   1  'Right Justify
            Caption         =   "Campo 4"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   24
            Left            =   200
            TabIndex        =   53
            Top             =   1830
            Width           =   900
         End
         Begin VB.Label LblFactura 
            Alignment       =   1  'Right Justify
            Caption         =   "Campo 5"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   25
            Left            =   5000
            TabIndex        =   52
            Top             =   360
            Width           =   900
         End
         Begin VB.Label LblFactura 
            Alignment       =   1  'Right Justify
            Caption         =   "Campo 7"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   27
            Left            =   5000
            TabIndex        =   51
            Top             =   1395
            Width           =   900
         End
         Begin VB.Label LblFactura 
            Alignment       =   1  'Right Justify
            Caption         =   "Campo 8"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   28
            Left            =   5000
            TabIndex        =   50
            Top             =   1875
            Width           =   900
         End
      End
      Begin VB.Frame FrmBusca1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   945
         Index           =   0
         Left            =   -70320
         TabIndex        =   37
         Top             =   4695
         Width           =   4560
         Begin VB.CommandButton BotBusca 
            Enabled         =   0   'False
            Height          =   330
            Index           =   1
            Left            =   4035
            MouseIcon       =   "FrmFactura.frx":0559
            MousePointer    =   99  'Custom
            Picture         =   "FrmFactura.frx":06AB
            Style           =   1  'Graphical
            TabIndex        =   81
            ToolTipText     =   "Imprimir"
            Top             =   555
            Width           =   435
         End
         Begin VB.TextBox Text3 
            BackColor       =   &H8000000F&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   90
            TabIndex        =   63
            Text            =   "Total Registros"
            Top             =   630
            Width           =   1305
         End
         Begin VB.TextBox TxtTot 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   1380
            Locked          =   -1  'True
            TabIndex        =   62
            Top             =   585
            Width           =   540
         End
         Begin VB.CommandButton BotBusca 
            Height          =   330
            Index           =   0
            Left            =   3570
            MouseIcon       =   "FrmFactura.frx":07AD
            MousePointer    =   99  'Custom
            Style           =   1  'Graphical
            TabIndex        =   39
            ToolTipText     =   "Buscar"
            Top             =   555
            Width           =   435
         End
         Begin VB.TextBox TxtBus 
            BackColor       =   &H00FFFFFF&
            Height          =   285
            Left            =   765
            TabIndex        =   38
            Top             =   210
            Width           =   3720
         End
         Begin VB.Label Label2 
            Caption         =   "&Buscar:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   105
            TabIndex        =   40
            Top             =   225
            Width           =   630
         End
      End
      Begin VB.Frame FrmBusca 
         Caption         =   "Buscar Por:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   1470
         Left            =   -74685
         TabIndex        =   35
         Top             =   4695
         Width           =   3840
         Begin VB.OptionButton OptBusca 
            Caption         =   "N° &Factura"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   240
            Index           =   1
            Left            =   2145
            TabIndex        =   84
            Top             =   420
            Width           =   1590
         End
         Begin VB.OptionButton OptBusca 
            Caption         =   "&Código Inmueble"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   225
            Index           =   0
            Left            =   225
            TabIndex        =   83
            Top             =   405
            Width           =   1830
         End
         Begin VB.OptionButton OptBusca 
            Caption         =   "Nº Documento"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   240
            Index           =   3
            Left            =   2145
            TabIndex        =   82
            Top             =   900
            Width           =   1590
         End
         Begin VB.OptionButton OptBusca 
            Caption         =   "Pr&oveedor"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   240
            Index           =   2
            Left            =   225
            TabIndex        =   36
            Top             =   870
            Width           =   1830
         End
      End
      Begin ComctlLib.ImageList ImageList1 
         Left            =   -75000
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
               Picture         =   "FrmFactura.frx":08FF
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "FrmFactura.frx":0A81
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "FrmFactura.frx":0C03
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "FrmFactura.frx":0D85
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "FrmFactura.frx":0F07
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "FrmFactura.frx":1089
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "FrmFactura.frx":120B
               Key             =   ""
            EndProperty
            BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "FrmFactura.frx":138D
               Key             =   ""
            EndProperty
            BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "FrmFactura.frx":150F
               Key             =   ""
            EndProperty
            BeginProperty ListImage10 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "FrmFactura.frx":1691
               Key             =   ""
            EndProperty
            BeginProperty ListImage11 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "FrmFactura.frx":1813
               Key             =   ""
            EndProperty
            BeginProperty ListImage12 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "FrmFactura.frx":1995
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin VB.Label LblFactura 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000005&
         DataField       =   "Total"
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "#,##0.00;(#,##0.00)"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   8202
            SubFormatType   =   1
         EndProperty
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   19
         Left            =   -73725
         TabIndex        =   75
         Top             =   3390
         Width           =   1605
      End
      Begin VB.Label LblFactura 
         BackColor       =   &H80000005&
         DataField       =   "DeTalle"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   18
         Left            =   -73725
         TabIndex        =   74
         Top             =   3060
         Width           =   8085
      End
      Begin VB.Label LblFactura 
         BackStyle       =   0  'Transparent
         Caption         =   "Total:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   17
         Left            =   -74745
         TabIndex        =   73
         Top             =   3420
         Width           =   420
      End
      Begin VB.Label LblFactura 
         BackStyle       =   0  'Transparent
         Caption         =   "Descripcion:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   16
         Left            =   -74745
         TabIndex        =   72
         Top             =   3090
         Width           =   900
      End
   End
   Begin MSComCtl2.MonthView MntCalendar 
      Height          =   2310
      Left            =   270
      TabIndex        =   64
      Top             =   3645
      Visible         =   0   'False
      Width           =   2565
      _ExtentX        =   4524
      _ExtentY        =   4075
      _Version        =   393216
      ForeColor       =   -2147483630
      BackColor       =   16777215
      BorderStyle     =   1
      Appearance      =   0
      ShowToday       =   0   'False
      StartOfWeek     =   54984705
      TitleBackColor  =   -2147483646
      TitleForeColor  =   16744576
      TrailingForeColor=   12632256
      CurrentDate     =   36942
   End
End
Attribute VB_Name = "FrmFactura"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'-----------------------------------------------'----------------------------------'--------------
Dim rst As ADODB.Recordset
Attribute rst.VB_VarHelpID = -1
Dim ObjRstF As ADODB.Recordset, ObjRstP As ADODB.Recordset, ObjRstFP As ADODB.Recordset
Dim I%, B%, IntDc%, Ramo$
Dim Mtto As Boolean
Public Descrip$
Dim rstLisPro(3) As New ADODB.Recordset
'-----------------------------------'----------------------------------'--------------------------

Private Sub Ado_MoveComplete(Index As Integer, ByVal adReason As ADODB.EventReasonEnum, _
ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As _
ADODB.Recordset)
'
If Index = 0 Then
    '
    If Not Ado(0).Recordset.EOF And Not Ado(0).Recordset.BOF Then
        '
        If Ado(0).Recordset.EditMode = adEditNone Then
            'busca la informacion del inmueble y propoetario
            Call RtnBuscaInmueble("CodInm", Ado(0).Recordset("CodInm"))
            Call RtnBuscaProveedor("Codigo", Ado(0).Recordset("CodProv"))
            '
        End If
        '
    End If
'
End If
'
End Sub

'-------------------------------------------------------------------------------------------------
Private Sub BotBusca_Click(Index As Integer)    '-
'-------------------------------------------------------------------------------------------------
'
Dim rpReporte As ctlReport
Select Case Index
    '
    Case 0  'Botón Buscar
'   -------------------------
        If TxtBus <> "" Then Call rtnBusca
        
    Case 1  'Boton Imprimir
'   -------------------------
        BotBusca(1).Enabled = False
        Set rpReporte = New ctlReport
        With rpReporte
        
            .Reporte = gcReport + "recepfc.rpt"
            If OptBusca(0).Value = True Then .FormuladeSeleccion = "{Cpp.CodInm}='" & TxtBus & "'"
            If OptBusca(1).Value = True Then .FormuladeSeleccion = "{Cpp.Fact}='" & TxtBus & "'"
            If OptBusca(2).Value = True Then .FormuladeSeleccion = "{Proveedores.NombProv}='" _
                & TxtBus & "'"
            Call rtnBitacora("Recepción de Facturas: Printer Busqueda " & .FormuladeSeleccion)
            .Imprimir
            
        End With
        Set rpReporte = Nothing
'
End Select
'
End Sub

'----------------------------
Private Sub Check1_Click() '-
'----------------------------

If Check1.Value = 1 Then

    MaskEdBox4.Text = (CCur(gnIva))
    lblFactura(29) = Format(CCur(TxtMonto(0) * MaskEdBox4 / 100), "#,##0.00")

Else

    MaskEdBox4.Text = Format(CCur(0), "#,##0.00")
    lblFactura(29) = Format(CCur(0), "#,##0.00")

End If
MaskEdBox4.Refresh
TxtMonto(0) = Format(CCur(TxtMonto(0)), "#,##0.00")
TxtMonto(1) = Format(CSng(TxtMonto(Index)) + CSng(lblFactura(29)), "#,##0.00")

End Sub

Private Sub Check1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then Check1.Value = IIf(Check1.Value = 1, 0, 1)
End Sub

'---------------------------------------------
Private Sub CmbCargo_Click(Area As Integer) '-
'---------------------------------------------
If Area = 2 Then    'Si hace Click en un elemento de la lista
    'Se ejecuta la Rutina (Moneda)
    Call RtnMoneda(CmbCargo, DataCombo5, "IDMoneda", "Moneda")
End If
End Sub

'----------------------------------------------------
Private Sub CmbCargo_KeyPress(KeyAscii As Integer) '-
'----------------------------------------------------
'Si presiona {enter} se ejecuta la rutina (RtnMoneda)
If KeyAscii = 13 Then Call RtnMoneda(CmbCargo, DataCombo5, "IDMoneda", "Moneda")
End Sub

    '-------------------------------------------------
    Private Sub CmbTipoPago_Click(Index As Integer) '-
    '-------------------------------------------------
    '
    With TxtDescripTipoDoc(0)
    '
        Select Case CmbTipoPago(0).ListIndex
            '
            Case 0: .Text = "Presupuesto"
            Case 1: .Text = "Factura"
            Case 2: .Text = "Caja Chica"
            Case 3: .Text = "Relac.Gastos"
            Case 4: .Text = "Caja Hono."
            '
        End Select
    '
    End With
    '
    End Sub

    '-------------------------------------------------------------------------
    Private Sub CmbTipoPago_KeyPress(Index As Integer, KeyAscii As Integer) '-
    '-------------------------------------------------------------------------
    '
    If Index = 0 Then Call Validacion(KeyAscii, "")

    If KeyAscii = 13 Then   'Si el usuario presiona {enter}
        Select Case Index
        
            Case 0
                If CmbTipoPago(Index) = "" Then _
                MsgBox "Debe Introducir algun valor...", vbExclamation: Exit Sub
                TxtFactura(0).SetFocus
            
        End Select
    '
    End If
    '
    End Sub

    '----------------------------------------------
    Private Sub CmdFecha_Click(Index As Integer) '-
    '----------------------------------------------
    If Index = 3 Then   'Command Ver Facturas
    '
    Set ObjRstF = New ADODB.Recordset
    'Selecciona todas las facturas que posiblemente tengan Presupuesto
    ObjRstF.Open "SELECT * From Cpp WHERE (((Tipo)='FC') AND ((FRecep)=#" _
        & Format(DtFechaFact, "MM/DD/YYYY") & "#) AND (([CodInm] & [CodProv]) In (SELECT CodInm" _
        & "& CodProv FROM Cpp WHERE  Tipo ='PR' AND Estatus<>'FACTURADO')) AND ((Estatus)='PEND" _
        & "IENTE')) ORDER BY CodInm, CodProv;", cnnConexion, adOpenKeyset, adLockOptimistic
    
    'Visualiza en el Grid el Resultado de la selección
    Set DtgPresupuesto(0).DataSource = ObjRstF
    Else    'Command de Fechas
        Call RtnMuestraCalendar(Index)
        B = Index
    End If
    '
    End Sub

    '-----------------------------------------------
    Private Sub DataCombo5_Click(Area As Integer) '-
    '-----------------------------------------------
    '
    If Area = 2 Then    'Si hace clik en un elemento de la lista
        Call RtnMoneda(DataCombo5, CmbCargo, "Moneda", "IDMoneda")
    End If
    End Sub
    '------------------------------------------------------
    Private Sub DataCombo5_KeyPress(KeyAscii As Integer) '-
    '------------------------------------------------------
    If KeyAscii = 13 Then Call RtnMoneda(DataCombo5, CmbCargo, "Moneda", "IDMoneda")
    End Sub

    '-------------------------------
    Private Sub DataGrid1_Click() '-
    '-------------------------------
    With Ado(0).Recordset
        .MoveFirst
        .Find "Ndoc = '" & rst.Fields("Ndoc") & "'"
        'If Not .EOF Then Call RtnAvanzaReg
    End With
    End Sub

'------------------------------------------------------------------
Private Sub DatInmueble_Click(Index As Integer, Area As Integer) '-
'------------------------------------------------------------------
'
If Area = 2 Then    'Si hace Click Sobre un elemento de la lista
    Select Case Index
        'Llama la rutina que busc inf. del inmueble seleccionado
        Case 0, 2: Call RtnBuscaInmueble("CodInm", DatInmueble(Index))
        'Llama la rutina que busc inf. del inmueble seleccionado
        Case 1, 3: Call RtnBuscaInmueble("Nombre", DatInmueble(Index))
    End Select
    '
End If
'
End Sub

'-------------------------------------------------------------------------
Private Sub DatInmueble_KeyPress(Index As Integer, KeyAscii As Integer) '-
'-------------------------------------------------------------------------

KeyAscii = Asc(UCase(Chr(KeyAscii)))    'Convierte todo en mayúsculas

Select Case Index
    'Llama la rutina que valida la entrada de datos
    Case 0, 2: Call Validacion(KeyAscii, "1234567890")
    'Llama la rutina que valida la entrada de datos
    Case 1, 3: Call Validacion(KeyAscii, " ABCDEFGHIJKLMNOPQRSTUVWQYZ.,-/")
    
End Select
'LLAMA LA RUTINA QUE BUSCA LA INFORMACION DEL INMUEBLE SELECCIONADO*
If KeyAscii = 13 Then   'Presionó {Enter}
    Select Case Index
        
        Case 0, 2
            If DatInmueble(Index) = "" Then DatInmueble(Index + 1).SetFocus: Exit Sub
            Call RtnBuscaInmueble("CodInm", DatInmueble(Index))
        
        Case 1, 3
            If DatInmueble(Index) = "" Then DatInmueble(Index + 1).SetFocus: Exit Sub
            Call RtnBuscaInmueble("Nombre", DatInmueble(Index))
        
    End Select

End If

End Sub
'-------------------------------------------------------------------
Private Sub DatProveedor_Click(Index As Integer, Area As Integer) '*
'-------------------------------------------------------------------
'
If Area = 2 Then 'SI HACE CLICK SOBRE ALGUN ELEMENTO DE LA LISTA
    IntDc = Index
    Call RtnBuscaProveedor(DatProveedor(Index).ListField, DatProveedor(Index).Text, True)
    
    If FraFactura(0).Enabled = True Then
        If IntDc = 0 Then
            CmbTipoPago(0).SetFocus
        Else
            CmbTipoPago(0).SetFocus
        End If
        CmbTipoPago(0).ListIndex = 1
    End If
    

End If
'
End Sub

'--------------------------------------------------------------------------
Private Sub DatProveedor_KeyPress(Index As Integer, KeyAscii As Integer) '-
'--------------------------------------------------------------------------

KeyAscii = Asc(UCase(Chr(KeyAscii))) 'CONVIERTE TODO A MAYUSCULAS
Select Case Index
    Case 0: Call Validacion(KeyAscii, "1234567890")
End Select

'LLAMA LA RUTINA QUE BUSCA LA INFORMACION REFERENTE AL PROVEEDOR*

If KeyAscii = 13 Then 'SI PRESIONA ENTER
    IntDc = Index
    
    Select Case Index
                
        Case 0
            
            If DatProveedor(0) = "" Then DatProveedor(1).SetFocus: Exit Sub
            Call RtnBuscaProveedor("Codigo", DatProveedor(Index).Text, True)
                        
        Case 1
            
            If DatProveedor(1) = "" Then DatProveedor(0).SetFocus: Exit Sub
            Call RtnBuscaProveedor("Nombprov", DatProveedor(Index).Text, True)
        
    End Select
    If FraFactura(0).Enabled = True Then
        If IntDc = 0 Then
            CmbTipoPago(0).SetFocus
        Else
            CmbTipoPago(0).SetFocus
        End If
        CmbTipoPago(0).ListIndex = 1
    End If
    
    If DatProveedor(0) <> "" And DatProveedor(1) <> "" Then
        If Index = 0 Then
            CmbTipoPago(0).SetFocus
        Else
            CmbTipoPago(0).SetFocus
        End If
        CmbTipoPago(0).ListIndex = 1
    Else
    DatProveedor(0).SetFocus
    End If

End If

End Sub



'----------------------------------------------------
Private Sub DtgPresupuesto_Click(Index As Integer) '-
'----------------------------------------------------
'
Select Case Index
'
    Case 1  'Case de Lista de Presupuestos
    '
        If Not ObjRstP.EOF Then
            Call RtnFaP
            FrmBusca1(3) = "Facturas-Presupuesto: " & Format(FntTotalFP, "#,##0.00")
            Call Listar_Facturas(DtgPresupuesto(1).Columns(1) & DtgPresupuesto(1).Columns(2))
        End If
        '
End Select
'
End Sub

'-------------------------------------------------------
Private Sub DtgPresupuesto_DblClick(Index As Integer) '-
'-------------------------------------------------------
'variables locales
Dim strMensaje As String
Dim StrNdoc As String
Dim curTotal As Currency
'
MousePointer = vbhorglass
Select Case Index
    'Lista de Facturas
    Case 0  'Valida que la factura pueda ser aplicada al presupuesto seleccionado
        '
        If ObjRstF.RecordCount <= 0 Then Exit Sub
        If ObjRstF!CodInm = ObjRstP!CodInm And _
            ObjRstF!codProv = ObjRstP!codProv Then
            
            'Si la Sum(Total) de las facturas del presupuesto excede
            'el monto total del mismo, consulta si se hace el ajuste
            'Call RtnFaP: CurTotal = FntTotalFP + ObjRstF!Total
            curTotal = FntTotalFP + ObjRstF!Total
            If curTotal > ObjRstP!Total Then
                curTotal = curTotal - ObjRstP!Total
                
                strMensaje = "Si agrega esta factura excederá el monto total del presupuesto." _
                & vbCrLf & "¿Desea continuar y generar el respectivo ajuste?"
                If Not Respuesta(strMensaje) Then
                    Exit Sub
                Else
                    'Genera el registro de ajuste en Cpp
                    StrNdoc = FntStrDoc
                    cnnConexion.Execute "INSERT INTO Cpp (Tipo,Ndoc,Fact,CodProv,Benef,Detalle,Monto,Ivm," _
                    & "Total,Frecep,Fecr,Fven,CodInm,Moneda,Estatus,Usuario,Freg) " _
                    & "SELECT Tipo,'" & StrNdoc & "',Fact, CodProv, Benef,Detalle, " & curTotal & ", 0, " _
                    & curTotal & ",Frecep,Fecr,Fven,CodInm,Moneda,Estatus,'" & gcUsuario & "','" _
                    & Date & "' FROM Cpp WHERE Ndoc ='" & ObjRstF!NDoc & "'"
                    '
                End If
                '
            End If
            'El campo Estatus Relaciona la Factura con el presupuesto asignado
            ObjRstF.Update "Estatus", ObjRstP!NDoc
            ObjRstF.Requery
            ObjRstFP.Requery
            rst.Requery
            curTotal = FntTotalFP
            FrmBusca1(3) = "Facturas-Presupuesto: " & Format(curTotal, "#,##0.00")
            If curTotal = ObjRstP!Total Then
                ObjRstP.Update "Campo8", ObjRstP!Estatus
                ObjRstP.Update "Estatus", "FACTURADO"
                ObjRstP.Requery
            End If
        Else
            
            MsgBox "Estimado(a): " & gcUsuario & " la factura que seleccionó " & vbCrLf & "no c" _
            & "orresponde con el presupuesto pre-seleccionado " & vbCrLf & "Verifique esto, y v" _
            & "uelca a intertarlo...", vbInformation, App.ProductName
            
        End If
        
    Case 2  'Case de Lista Factura--->Presupuesto
        On Error Resume Next
        ObjRstFP.Update "Estatus", "PENDIENTE"
        ObjRstF.Requery
        ObjRstFP.Requery
        rst.Requery
        curTotal = FntTotalFP
        FrmBusca1(3) = "Facturas-Presupuesto: " & Format(curTotal, "#,##0.00")
        If curTotal <> ObjRstF!Total Then
            ObjRstP.Update "Estatus", ObjRstP!Campo8
            ObjRstP.Update "Campo8", ""
            ObjRstP.Requery
        End If
        '
End Select
MousePointer = vbDefault
'
End Sub

    '---------------------------------------------------------------------------------------------
    Private Sub Form_Load() '-
    '---------------------------------------------------------------------------------------------
    'variables locales
    
    Dim StrInmueble$, strSQL$
    '
    BotBusca(0).Picture = LoadResPicture("Buscar", vbResIcon)
    rstLisPro(0).Open "SELECT * FROM Proveedores ORDER BY Codigo", cnnConexion, adOpenKeyset, _
    adLockOptimistic, adCmdText
    
    rstLisPro(1).Open "SELECT * FROM Proveedores ORDER BY NombProv", cnnConexion, adOpenKeyset, _
    adLockOptimistic, adCmdText
    
    rstLisPro(2).Open "SELECT * FROM Inmueble ORDER BY Codinm", cnnConexion, adOpenKeyset, _
    adLockOptimistic, adCmdText
    
    rstLisPro(3).Open "SELECT * FROM Inmueble ORDER BY Nombre", cnnConexion, adOpenKeyset, _
    adLockOptimistic, adCmdText
'    With Ado(1)
'        .CursorLocation = adUseClient
'        .CursorType = adOpenDynamic
'        .LockType = adLockOptimistic
'        .ConnectionString = cnnConexion.ConnectionString
'        .RecordSource = "SELECT * FROM Inmueble ORDER BY Codinm"
'        .CommandType = adCmdText
'        .Refresh
'    End With
    
'    With Ado(2)
'        .CursorLocation = adUseClient
'        .CursorType = adOpenStatic
'        .LockType = adLockReadOnly
'        .RecordSource = "SELECT * FROM Inmueble ORDER BY Nombre"
'        .CommandType = adCmdText
'        .ConnectionString = cnnConexion.ConnectionString
'        .Refresh
'    End With
    
    With Ado(0)
        '.CursorLocation = adUseClient
        '.CursorType = adOpenStatic
        '.LockType = adLockOptimistic
        .RecordSource = "Cpp"
        .CommandType = adCmdTable
        .ConnectionString = cnnConexion.ConnectionString
        .Refresh
        .Recordset.Filter = "Estatus='PENDIENTE'"
        .Recordset.Sort = "FRecep DESC"
    End With
    '------------------------------------------------------------------------------------
    Set DatInmueble(0).DataSource = Ado(0)
    'Set DatProveedor(0).DataSource = Ado(0)
    Set TxtNdoc(0).DataSource = Ado(0)
    Set TxtFactura(0).DataSource = Ado(0)
    Set TxtFactura(1).DataSource = Ado(0)
    Set MskFecha(0).DataSource = Ado(0)
    Set MskFecha(1).DataSource = Ado(0)
    Set MskFecha(2).DataSource = Ado(0)
    Set CmbCargo.DataSource = Ado(0)
    Set TxtFactura(4).DataSource = Ado(0)
    Set TxtMonto(0).DataSource = Ado(0)
    Set lblFactura(29).DataSource = Ado(0)
    Set TxtMonto(1).DataSource = Ado(0)
    Set CmbTipoPago(0).DataSource = Ado(0)
    Set CmbCargo.DataSource = Ado(0)
    Set TxtCampo1.DataSource = Ado(0)
    Set TxtCampo2.DataSource = Ado(0)
    Set TxtCampo3.DataSource = Ado(0)
    Set TxtCampo4.DataSource = Ado(0)
    Set TxtCampo5.DataSource = Ado(0)
    Set TxtCampo6.DataSource = Ado(0)
    Set TxtCampo7.DataSource = Ado(0)
    Set TxtCampo8.DataSource = Ado(0)
    '------------------------------------------------------------------------
    lblFactura(0) = LoadResString(101)
    lblFactura(1) = LoadResString(124)
    lblFactura(2) = LoadResString(125)
    lblFactura(3) = LoadResString(126)
    lblFactura(4) = LoadResString(127)
    '
    
    Set DatProveedor(0).RowSource = rstLisPro(0)
    Set DatProveedor(1).RowSource = rstLisPro(1)
    Set DatInmueble(0).RowSource = rstLisPro(2)
    Set DatInmueble(1).RowSource = rstLisPro(3)
        
    Set DataGrid1.HeadFont = LetraTitulo(LoadResString(527), 7.5, True)
    Set DataGrid1.Font = LetraTitulo(LoadResString(528), 8)
    '
    'Configura la presentacion del sstab1
    
    Set rst = New ADODB.Recordset
    rst.CursorLocation = adUseClient
    rst.CursorType = adOpenKeyset
    rst.LockType = adLockOptimistic
    StrInmueble = "SELECT monedas.moneda FROM monedas WHERE monedas.idmoneda = '" & CmbCargo & "'"
    rst.Source = StrInmueble
    rst.ActiveConnection = cnnConexion
    rst.Open
    If Not rst.EOF And Not rst.BOF Then
        DataCombo5.Text = rst.Fields("moneda")
    Else
        DataCombo5.Text = ""
    End If
    rst.Close
    Set rst = Nothing
'    If Not Ado(0).Recordset.EOF Then
'        If Ado(0).Recordset.Fields("ivm") > 0 Then
'            Check1.Value = 1
'        Else
'            Check1.Value = 0
'        End If
'        Check1.Refresh
'        SSTab1.Tab = 0
'    End If
    
'    If DatInmueble(0) <> "" Then    'SI EXISTE UN COD. DE INMUEBLE ENTONCES
'        Dim RsCmd As ADODB.Recordset
'        For Each RsCmd In DEFrmFactura.Recordsets
'            If RsCmd.State = 0 Then
'            RsCmd.Open
'            End If
'        Next
'        If DatInmueble(0) = "" Then
'            MsgBox "Falta Código del Inmueble. Consulte al proveedor", vbInformation, _
'            App.ProductName
'        ElseIf DatProveedor(Index).Text = "" Then
'            MsgBox "Falta Código del Proveedor. Consulte al proveedor", vbInformation, _
'            App.ProductName
'        Else
'            Call RtnBuscaInmueble("CodInm", DatInmueble(0))
'            Call RtnBuscaProveedor("Codigo", DatProveedor(Index).Text)
'        End If
'    End If
    '
    strSQL = "SELECT Cpp.*, Proveedores.NombProv As Name FROM Proveedores INNER JOIN Cpp ON Pro" _
    & "veedores.Codigo = Cpp.CodProv WHERE Estatus='PENDIENTE' ORDER BY Cpp.Frecep DESC, " _
    & "Proveedores.NombProv;"
    '
    Set rst = New ADODB.Recordset
    rst.Open strSQL, cnnConexion, adOpenStatic, adLockReadOnly, adCmdText
    Set DataGrid1.DataSource = rst
    '
    TxtTot = rst.RecordCount
      '
    End Sub

    Private Sub Form_Resize()
    'configura la presentacion del formulario
    'en pantalla
    If Me.WindowState = vbMaximized Then
        '
        SSTab1.Top = (ScaleHeight / 2) - (SSTab1.Height / 2)
        SSTab1.Left = (ScaleWidth / 2) - (SSTab1.Width / 2)
        FraFactura(0).Top = SSTab1.Top + 405
        FraFactura(0).Left = SSTab1.Left + 255
    '
    End If
    '
    End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
rstLisPro(0).Close
rstLisPro(1).Close
rstLisPro(2).Close
rstLisPro(3).Close
End Sub

'    '--------------------------------------------
'    Private Sub Form_Unload(Cancel As Integer) '-
'    '--------------------------------------------
''    Dim RsCmd As ADODB.Recordset
''    'Cierrra Todos Los ADODB.Recordset que permanezcan abiertos
''    For Each RsCmd In DEFrmFactura.Recordsets
''        RsCmd.Close
''        Set RsCmd = Nothing
''    Next
'
'    '
'    End Sub


    '--------------------------------------------------------------
    Private Sub MntCalendar_DateClick(ByVal DateClicked As Date) '-
    '--------------------------------------------------------------
    '
    MskFecha(B) = DateClicked: MntCalendar.Visible = False
    If B = 0 Then
        MskFecha(0).PromptInclude = True
        'asigna fecha vencimiento
        MskFecha(2) = fecha_vcto
        MskFecha(1).SetFocus
    End If
    If B = 2 Or B = 1 Then TxtFactura(4).SetFocus
    If B = 0 Then MskFecha(B + 1).SetFocus
    '
    End Sub

'-------------------------------------------------------
Private Sub MntCalendar_KeyPress(KeyAscii As Integer) '-
'-------------------------------------------------------
If KeyAscii = 13 Then   'Presionó {enter}
    
    MskFecha(B) = MntCalendar.Value: MntCalendar.Visible = False
    If B = 1 Then MskFecha(B).PromptInclude = True
    If B = 0 Then
        MskFecha(B).PromptInclude = True
        MskFecha(2) = fecha_vcto
        MskFecha(1).SetFocus 'TxtFactura(4).SetFocus
    End If
    If B = 2 Or B = 1 Then TxtFactura(4).SetFocus
    'If B = 0 Then MskFecha(B + 1).SetFocus

End If
'
End Sub

'----------------------------------------------------------------------
Private Sub MskFecha_KeyPress(Index As Integer, KeyAscii As Integer) '-
'----------------------------------------------------------------------
'
B = Index
If KeyAscii = 13 Then   'Presionó {enter}
    
    If MskFecha(Index) = "" Then    'si es nulo el valor del campo
        Call RtnMuestraCalendar(Index)
    Else
        If Len(MskFecha(1)) < 8 Then
            MsgBox "Error en el formato de fecha" _
            & Chr(13) & "Formato: 'DD/MM/AAAA'" _
            & Chr(13) & "Ejm: " & Date
            With MskFecha(1)
                .SetFocus
                .SelStart = 0
                .SelLength = Len(.Text) + 2
            End With
            Exit Sub
        End If
    MskFecha(Index).PromptInclude = True
    If Not IsDate(MskFecha(Index).Text) Then
        MsgBox "Fecha nó valida", vbCritical, App.ProductName
        With MskFecha(Index)
            .SetFocus
            .SelStart = 0
            .SelLength = Len(.Text) + 2
            .PromptInclude = False
        End With
        Exit Sub
    End If
    MskFecha(Index).PromptInclude = False
    If B = 0 Then
        MskFecha(B).PromptInclude = True
        MskFecha(2) = fecha_vcto
        MskFecha(1).SetFocus 'TxtFactura(4).SetFocus
    End If
    If B = 2 Or B = 1 Then TxtFactura(4).SetFocus
    'If B = 0 Then MskFecha(B + 1).SetFocus
    
    End If
    
End If

End Sub

    '---------------------------------------------------------------------------------------------
    Private Sub OptBusca_Click(Index As Integer) '-
    '---------------------------------------------------------------------------------------------
    '
    Select Case Index   'Selecciona la opción de busqueda para ordenar las facturas
    
        Case 0: rst.Sort = "CodInm, Frecep DESC"  'Ordena Por Codigo de Inmueble"
    '   -------------------------------
    
        Case 1: rst.Sort = "Fact DESC" 'ordena por n. factura
    '   -------------------------------

        Case 2: rst.Sort = "Name, Frecep DESC" 'Por Nombre del Beneficiario
    '   -------------------------------
    
        Case 3: rst.Sort = "Ndoc DESC"  'Por número de documento
    '   -------------------------------
    '
    End Select
    '
    With TxtBus 'Enfoca y Selecciona el Valor de Campo TxtBus
        .SetFocus
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
    '
    End Sub

    'Rev.-22/08/2001------------------------------------------------------------------------------
    Private Sub SSTab1_Click(PreviousTab As Integer) '-
    '---------------------------------------------------------------------------------------------
    '
    Select Case SSTab1.tab  'Ocurre al clickear un ficha
    '
        Case 0  'Ficha Datos Generales
    '   -----------------------------------------
        FraFactura(0).Visible = True
    '
        Case 1  'Ficha Datos Administrativos
    '   -----------------------------------------
        MntCalendar.Visible = False
        FraFactura(0).Visible = True
        
    '
        Case 2  'Ficha Lista
    '   -----------------------------------------
            FraFactura(0).Visible = False
            MntCalendar.Visible = False
'            Set DataGrid1.DataSource = Nothing
            rst.Requery
'            Set DataGrid1.DataSource = rst
    '
        Case 3  'Ficha Presupuestos
    '   -----------------------------------------
        FraFactura(0).Visible = False
        MntCalendar.Visible = False
        Set ObjRstP = New ADODB.Recordset
    '   ------------------------------------------------------------------------------------------
        'Selecciona todos los presupuestos pendientes por asignar
        'ObjRstP.Open "SELECT * From Cpp WHERE (((Cpp.Tipo)='PR') AND Estatus<>'FACTURADO'" _
        & "AND (([CodInm] & [CodProv]) In (SELECT CodInm & CodProv FROM Cpp WHERE  Tipo ='FC' A" _
        & "ND Estatus<>'FACTURADO'))) ORDER BY Cpp.CodInm, Cpp.CodProv;", cnnConexion, _
        adOpenKeyset, adLockOptimistic
        ObjRstP.Open "SELECT * FROM Cpp WHERE Tipo='PR' AND EStatus <>'FACTURADO' ORDER BY Codin" _
        & "m,CodProv", cnnConexion, adOpenKeyset, adLockOptimistic, adCmdText
        '
        'Visualiza en el Grid el Resultado de la selección
        DtgPresupuesto(1).SetFocus
        Set DtgPresupuesto(1).DataSource = ObjRstP
        Set lblFactura(18).DataSource = ObjRstP
        Set lblFactura(19).DataSource = ObjRstP
    '   ------------------------------------------------------------------------------------------
        If Not ObjRstP.EOF Then Call RtnFaP
    '
    End Select
    '
    End Sub

    '--------------------------------------------------------------------
    Private Sub Toolbar1_ButtonClick(ByVal Button As ComctlLib.Button) '-
    '--------------------------------------------------------------------
    Dim strProveedor As String
    'Eventos Botones de la Barra de Herramientas y Navegación
    
    With Ado(0).Recordset
        
        Select Case Button.Index
            Case 1 'Primer Registro
                If .RecordCount <= 0 Then Exit Sub
                .MoveFirst
                'Call RtnAvanzaReg

            Case 2 ' Registro Anterior
        
            If .RecordCount <= 0 Then Exit Sub
            .MovePrevious
            If .BOF Then .MoveLast
            'Call RtnAvanzaReg
            
        Case 3 'Registro Siguiente
            If .RecordCount <= 0 Then Exit Sub
            .MoveNext
            If .EOF Then .MoveFirst
            'Call RtnAvanzaReg
        
        Case 4  ' Ultimo Registro
            If .RecordCount <= 0 Then Exit Sub
            .MoveLast
            'Call RtnAvanzaReg
            
        Case 5 'agregar registro
            'Nuevo Registro
            SSTab1.tab = 0
            For I = 0 To 3
                FraFactura(I).Enabled = True
            Next I
            .AddNew
            Call rtnBitacora("Recepción de facutra: Nuevo Registro")
            Call RtnEstado(5, Toolbar1)
            DatInmueble(1) = ""
            DatProveedor(1) = ""
            DatProveedor(0) = ""
            MaskEdBox4.Text = CCur(gnIva)
            CmbCargo.Text = IDMoneda
            'DataCombo5.Text = CmbCargo.BoundText
            TxtNdoc(0) = "0000000"
            SSTab1.TabEnabled(2) = False
            TxtMonto(0) = Format(0, "#,##0.00")
            TxtMonto(1) = Format(0, "#,##0.00")
            lblFactura(29) = Format(0, "#,##0.00")
            Check1.Value = 1
            DatInmueble(0).SetFocus
            
        Case 6  'Actualizar Registro
            
            If Not Campos_Requeridos Then
                If .EditMode = adEditAdd Then If Factura_Duplicada Then Exit Sub
                
                For I = 0 To 3: FraFactura(I).Enabled = False
                Next I
            
                If TxtMonto(0) <> "" Then
                    
                    If TxtMonto(0) <= 0 Then
                        
                        MsgBox "El monto debe ser mayor a 0 (cero)...", vbInformation, App.ProductName
                        .CancelUpdate
                        Call rtnBitacora("Recepción de Factura:Cancelar Ndoc: " & TxtNdoc(0))
                    Else
                        '.Fields("total") = CCur(TxtMonto(1))
                        .Fields("Estatus") = "PENDIENTE"
                        .Fields("Usuario") = gcUsuario
                        .Fields("Freg") = Date
                        .Fields("CodProv") = DatProveedor(0)
                        strProveedor = DatProveedor(0)
                        For I = 0 To 2
                            MskFecha(I).PromptInclude = True
                        Next
                        If .EditMode = adEditAdd Then TxtNdoc(0) = FntStrDoc
                        Ado(0).Recordset.Update
                        Call rtnBitacora("Recepción de Factura.Update Ndoc: " & TxtNdoc(0))
                        For I = 0 To 2
                            MskFecha(I).PromptInclude = False
                        Next
                        Call Actualiza_DeudaProv(strProveedor)
                        '
                        rst.Requery
                        DataGrid1.ReBind
                        Call RtnEstado(6, Toolbar1)
                        MsgBox "Registro actualizado...", vbInformation, App.ProductName
                    End If
                Else
                    MsgBox "No se pudo actualizar el registro, verifique el monto...", _
                    vbInformation, App.ProductName
                End If
                SSTab1.TabEnabled(2) = True
            End If
        Case 7  'Buscar Registro
            SSTab1.tab = 2
            TxtBus.SetFocus
        
        Case 8  'Cancelar Registro
            Call RtnEstado(6, Toolbar1)
            For I = 0 To 2
                MskFecha(I).PromptInclude = True
            Next
            .CancelUpdate
            For I = 0 To 2: MskFecha(I).PromptInclude = False
            Next
            For I = 0 To 3: FraFactura(I).Enabled = False
            Next I
            SSTab1.TabEnabled(2) = True
'            If DatInmueble(0) <> "" And Len(DatInmueble(0)) = _
'            4 Then Call RtnBuscaInmueble("CodInm", DatInmueble(0))
'            If DatProveedor(0) <> "" Then
'                Call RtnBuscaProveedor("Codigo", DatProveedor(Index).Text)
'            End If
'
            
        Case 9  'Eliminar Registro
            
            If Respuesta(LoadResString(526)) Then
                strProveedor = DatProveedor(0)
                Call rtnBitacora("Recepción de Facturas:Delete Ndoc:" & TxtNdoc(0))
                    .Delete
                    rst.Requery
                    rst.Requery
                MsgBox "Factura eliminada...", vbInformation, App.ProductName
                .MoveNext
                If .EOF Then .MoveLast
                
                'Call RtnAvanzaReg
                Call Actualiza_DeudaProv(strProveedor)
                
            End If
            rst.Requery
            
        Case 12: Unload Me
            
        Case 11
            FrmReport.Frame1.Visible = True
            FrmReport.MskDesde = Date
            FrmReport.MskHasta = Date
            mcTitulo = "Reporte de Recepción de Facturas"
            mcReport = "RecepFc.Rpt"
            mcOrdCod = ""
            mcOrdAlfa = ""
            'mcCrit = "{cpp.FRecep}>={@desde} and {cpp.FRecep}<={@hasta}"
            With FrmReport
                .Frame1.Visible = True
                .Show
            End With
            
        Case 10 'Editar Registro
            If .Fields("estatus") = "ASIGNADO" Then
                MsgBox "Este Documento ya fué Asignado como un Gasto, no puede Editarlo...", _
                vbInformation, App.ProductName
                Exit Sub
            End If
            '
            Call rtnBitacora("Recepción de Factura:Edit Ndoc:" & TxtNdoc(0))
            Call RtnEstado(5, Toolbar1)
            For I = 0 To 3: FraFactura(I).Enabled = True
            Next
            
    End Select
    '
End With
'
End Sub

'--------------------------------------------------
Private Sub TxtBus_KeyPress(KeyAscii As Integer) '-
'--------------------------------------------------
KeyAscii = Asc(UCase(Chr(KeyAscii)))    'Convierte todo en mayúsculas
'
If KeyAscii = 13 Then   'Si Presionó {enter}
'
    If TxtBus <> "" Then Call rtnBusca
'
End If
End Sub


'---------------------------------------------------
Private Sub TxtFactura_GotFocus(Index As Integer) '-
'---------------------------------------------------
'
Select Case Index
    
    Case 0  '
    '
    With TxtFactura(0)
        Randomize
        .Text = Right(DatInmueble(0), 2) & Int((99999 - 0 + 1) * Rnd + 0)
        .SelStart = 0
        .SelLength = Len(Trim(.Text))
    End With
    '
    Case 1  '
    '
        If Descrip <> "" Then
            TxtFactura(1) = UCase("MTTO. " & Ramo & " MES " & Descrip & " " & DatInmueble(1) _
            & "(" & DatInmueble(0) & ")")
            'SendKeys (Chr(13))
        End If
    '
End Select
'
End Sub

'------------------------------------------------------------------------
Private Sub TxtFactura_KeyPress(Index As Integer, KeyAscii As Integer) '-
'------------------------------------------------------------------------
    
KeyAscii = Asc(UCase(Chr(KeyAscii)))    'Convierte en Mayuscula


If KeyAscii = 13 Then   'Avanza al presionar {enter}
            
    If TxtFactura(Index).Text = "" Then 'El valor del campo es 'vacío'
            
        MsgBox "Debe Ingresar la Expresión", vbInformation, App.ProductName
        TxtFactura(Index).SetFocus
        Exit Sub    'Sale de la rutina
            
    End If
                    
    Select Case Index
        
        Case 0
            If Not Factura_Duplicada Then TxtFactura(Index + 1).SetFocus
                
        Case 1
            MskFecha(0).Text = Date
            MskFecha(2) = fecha_vcto
            MskFecha(Index).SetFocus

        Case 4
            With TxtMonto(0)
                .SetFocus
                .SelStart = 0
                .SelLength = Len(.Text)
            End With
        
        End Select
           
    End If
        
End Sub

'----------------------------------------------------
Sub RtnBuscaInmueble(StrCampo, StrValor As String) '-
'----------------------------------------------------

'RUTINA QUE BUSCA LA INFORMACION DEL INMUEBLE
'SELECCIONADO POR EL USUARIO*
With rstLisPro(2)
    .MoveFirst
    .Find StrCampo & " like '%" & StrValor & "%'"
    If Not .EOF Then
        If UCase(StrCampo) <> "CODINM" Then DatInmueble(0) = .Fields("CodInm")
        DatInmueble(1) = .Fields("Nombre")
        'DatInmueble(0) = .Fields("CodInm")
    Else
        MsgBox "Inmueble No Registrado..", vbExclamation, App.ProductName
        DatInmueble(1) = ""
        If FraFactura(0).Enabled Then DatInmueble(0).SetFocus
        Exit Sub
    End If
End With
If FraFactura(0).Enabled = True Then DatProveedor(0).SetFocus
End Sub

'-------------------------------------------------
Private Sub TxtMonto_GotFocus(Index As Integer) '-
'-------------------------------------------------
With TxtMonto(Index)
    .SelStart = 0
    .SelLength = Len(TxtMonto(Index))
End With
End Sub

'----------------------------------------------------------------------
Private Sub TxtMonto_KeyPress(Index As Integer, KeyAscii As Integer) '-
'----------------------------------------------------------------------

If KeyAscii = 46 Then KeyAscii = 44 'CONVIERTE EL PUNTO EN COMA
Call Validacion(KeyAscii, "1234567890.,")

If KeyAscii = 13 Then   'Presionó {enter}
    Select Case Index
        
        Case 0
            
            If Check1.Value = 1 Then
                If Not IsNumeric(MaskEdBox4) Then MaskEdBox4 = 0
                lblFactura(29) = Format(CCur(TxtMonto(0) * MaskEdBox4 / 100), "#,##0.00")
                
            Else
                lblFactura(29) = Format(0, "#,##0.00")
                
            End If
            TxtMonto(0) = Format(CCur(TxtMonto(0)), "#,##0.00")
            TxtMonto(1) = Format(CSng(TxtMonto(Index)) + CSng(lblFactura(29)), "#,##0.00")
        
        Case 1
            
        
    End Select
End If
End Sub

    '---------------------------------------------------------------------------------------------
    Private Sub RtnMuestraCalendar(Index As Integer) '-
    '---------------------------------------------------------------------------------------------
    '
    MntCalendar.Visible = Not MntCalendar.Visible
    If MntCalendar.Visible = True Then
    MntCalendar.Top = 3660 + (405 * Index)
    MntCalendar.SetFocus: MntCalendar.Value = Date
    '
    End If
    '
    End Sub

    '-----------------------------------------------------
    Sub RtnBuscaProveedor(StrCampo, StrValor As String, Optional Busca As Boolean) '-
    '-----------------------------------------------------
    'RUTINA QUE BUSCA CODIGO/NOMBRE DE PROVEEDOR*
    Descrip = ""
    With rstLisPro(0)
        '
        .MoveFirst
        .Find StrCampo & " like '%" & StrValor & "%'"
        Mtto = False
        If Not .EOF Then
            DatProveedor(0) = .Fields("Codigo")
            DatProveedor(1) = .Fields("NombProv")
            If Not Ado(0).Recordset.EditMode = adEditNone Then
                TxtFactura(4) = IIf(IsNull(.Fields("Beneficiario")), TxtFactura(4), .Fields("Beneficiario"))
            End If
            '
            If !Actividad Like "MANT*" And Busca Then
                MousePointer = vbHourglass
                If IsNull(!Ramo) Or !Ramo = "" Or IsNull(!Actividad) Or !Actividad = "" Then
                    MsgBox gcUsuario & " cancela esta operación y completa la siguiente informa" _
                    & "ción en la ficha de este proveedor: Ramo y/o Actividad", vbInformation, _
                    App.ProductName
                Else
                    Mtto = True
                    Ramo = !Ramo
                    frmList.Proveedor = !Codigo
                    frmList.Inm = DatInmueble(0)
                    frmList.Mantenimiento = True
                    frmList.Show vbModal
                End If
                MousePointer = vbDefault
            End If
            '
        Else
            MsgBox "Proveedor No Registrado", vbExclamation, App.ProductName
        End If
    End With
    
    End Sub

    '----------------------------
    Private Sub RtnAvanzaReg() '-
    '----------------------------
'    'OCURRE CADA VEZ QUE EL APUNTADOR
'    'CAMBIA DE POSICION EN EL ADODB.Recordset
'    Call RtnBuscaInmueble("CodInm", DatInmueble(0)) 'Busca El Inmueble
'    Call RtnBuscaProveedor("Codigo", DatProveedor(Index).Text)  'Busca El Proveedor
'    If Ado(0).Recordset.Fields("ivm") > 0 Then
'        Check1.Value = 1
'    Else
'        Check1.Value = 0
'    End If
    End Sub

    '----------------------------------------------------------------------------------
    Private Sub RtnMoneda(Data1, Data2 As DataCombo, StrValor1, StrValor2 As String) '-
    '----------------------------------------------------------------------------------
'    With DEFrmFactura.rsCmdMoneda
'        .MoveFirst
'        .Find StrValor1 & " ='" & Data1 & "'"
'        If Not .EOF Then Data2 = .Fields(StrValor2)
'    End With
    End Sub

    '----------------------
    '   Rutina: RtnFap
    '
    '   Selecciona y muestra en el grid las facturas correspondientes a determinado
    '   presupuesto
    Private Sub RtnFaP() '-
    '----------------------
    Set ObjRstFP = New ADODB.Recordset
    ObjRstFP.Open "SELECT * FROM Cpp WHERE Estatus='" & ObjRstP!NDoc & "'", cnnConexion, _
    adOpenKeyset, adLockOptimistic
    Set DtgPresupuesto(2).DataSource = ObjRstFP
    End Sub

    '-------------------------------
    Private Function FntTotalFP() '-
    '-------------------------------
    'Calcula la sumatoria de las facturas
    'asignadas a un determinado presupuesto
    With ObjRstFP
        If .EOF Then FntTotalFP = 0: Exit Function
        .MoveFirst
        Do Until .EOF
            FntTotalFP = FntTotalFP + ObjRstFP!Total
            .MoveNext
        Loop
    End With
    End Function

    '---------------------------------------------------------- '
    Public Function FntStrDoc() As String                       '-
    '---------------------------------------------------------- '
    'variables locales
    Dim adoCppTran As ADODB.Recordset
    '
    FntStrDoc = Trim(Right(Format(Date, "dd/mm/yyyy"), 2)) + Trim(Format(Month(Date), "00"))
    '
    Set adoCppTran = New ADODB.Recordset
    adoCppTran.Open "SELECT MAX(ndoc) AS Numero FROM Cpp WHERE Ndoc LIKE '" & FntStrDoc & "%'", _
    cnnConexion, adOpenKeyset, adLockOptimistic
    If IsNull(adoCppTran!Numero) Then
        FntStrDoc = FntStrDoc + "001"
    Else
        FntStrDoc = FntStrDoc + Format(CInt(Mid(adoCppTran!Numero, 5)) + 1, "000")
    End If
    
    adoCppTran.Close
    Set adoCppTran = Nothing
    End Function

    '---------------------------------------------------------------------------------------------
    Private Sub rtnBusca()  '-
    '---------------------------------------------------------------------------------------------
    Dim StrCampo$
    
    If OptBusca(0).Value = False And _
    OptBusca(1).Value = False And OptBusca(2).Value = False And OptBusca(3).Value = False Then
    '
            MsgBox "Debe Seleccionar por lo menos una opción de Busqueda " & Chr(13) & "Marque " _
                & "una casilla de opción en el recuadro 'Buscar Por:'" & Chr(13) _
                & "Intentelo Nuevamente", vbInformation, App.ProductName
            Exit Sub    'Sale de la Rutina
    '
    End If
    '
    If OptBusca(0).Value = True Then StrCampo = "CodInm"
    If OptBusca(1).Value = True Then StrCampo = "Fact"
    If OptBusca(2).Value = True Then StrCampo = "Benef"
    If OptBusca(3).Value = True Then StrCampo = "Ndoc"
    
    '
    With rst    'con el ADODB.Recordset
    '
        .MoveFirst
        .Find StrCampo & " Like '*" & Trim(TxtBus.Text) & "*'"
        If .EOF Then
            MsgBox "No encontré la expresión seleccionada....", vbInformation, App.ProductName
            TxtBus.SetFocus
        Else
            BotBusca(1).Enabled = True
            TxtBus = .Fields(StrCampo)
        End If
    '
    End With
    '
    End Sub

    
    '---------------------------------------------------------------------------------------------
    '   Funcion:    Factura_duplicada
    '
    '   Devuelve 'True' si para el mismo proveedor se tiene registrada una
    '   factura con igual número, de lo contrario retorna 'False'
    '---------------------------------------------------------------------------------------------
    Private Function Factura_Duplicada() As Boolean
    '
    Dim rstFD As New ADODB.Recordset     'Variables locales
    Dim strSQL As String
    Dim Res As Integer
    '
    strSQL = "SELECT * FROM Cpp WHERE CodProv='" & DatProveedor(0) & "' AND CodInm='" & _
    DatInmueble(0) & "';"
    '
    rstFD.Open strSQL, cnnConexion, adOpenStatic, adLockReadOnly, adCmdText
    rstFD.Filter = "Fact='" & TxtFactura(0) & "'"
    
    If Not rstFD.EOF Or Not rstFD.BOF Then
        '
        Factura_Duplicada = MsgBox("Este proveedor ya tiene esta factura registrada " _
        & vbCrLf & "Fecha de Recepción: " & rstFD!Frecep & vbCrLf & "Bolívar" _
        & "es: " & Format(rstFD!Total, "#,##0.00"), vbInformation, "Factura Duplicada")
        Call rtnBitacora("Intento Fact. Duplicada Prov:" & DatProveedor(0) & "/" & TxtFactura(0))
        '
    End If
    rstFD.Filter = 0
    rstFD.Filter = "Total=" & CCur(TxtMonto(1))
    If Not rstFD.EOF Or Not rstFD.BOF Then
        '
        Res = MsgBox("Este proveedor ya tiene registrada una factura por el mismo monto" & _
        vbCrLf & "¿Desea ver las facturas registradas?", vbQuestion + vbYesNo, App.ProductName)
        If Res = vbYes Then
            frmList.Proveedor = DatProveedor(0)
            frmList.Inm = DatInmueble(0)
            frmList.Mantenimiento = False
            frmList.MontoFactura = CCur(TxtMonto(1))
            frmList.Show vbModal
            If FrmFactura.Tag = 1 Then Factura_Duplicada = True
        Else
            
            Res = MsgBox("¿Desea guardarla de todas formas?", vbQuestion + vbYesNo, _
            App.ProductName)
            If Res = vbNo Then Factura_Duplicada = True
        End If
        '
    End If
    '
    rstFD.Close
    Set rstFD = Nothing
    '
    End Function

    Private Sub Listar_Facturas(Proveedor As String)
    '
    Set ObjRstF = New ADODB.Recordset
    'Selecciona todas las facturas que posiblemente tengan Presupuesto
    ObjRstF.Open "SELECT * FROM Cpp WHERE Tipo='FC' AND CodInm & CodProv='" & Proveedor & _
    "' AND Estatus='PENDIENTE'", cnnConexion, adOpenKeyset, adLockOptimistic, adCmdText
    'Visualiza en el Grid el Resultado de la selección
    Set DtgPresupuesto(0).DataSource = ObjRstF
    '
    End Sub

    '---------------------------------------------------------------------------------------------
    '   Rutina:     Actualiza_deudaProv
    '
    '   Entrada:    Proveedor, variable que contiene el código del prov.
    '
    '   Suma las facturas pendientes del proveedor determinado y actualiza el
    '   campo deuda de la tabla proveedores
    '---------------------------------------------------------------------------------------------
    Private Sub Actualiza_DeudaProv(Proveedor As String)
    'variables locales
    Dim rstDeuda As New ADODB.Recordset
    Dim HoraI As Date
    '
    rstDeuda.CursorLocation = adUseClient
    rstDeuda.Open "SELECT Sum(Total) as Resta FROM Cpp WHERE CodProv='" & Proveedor & "' AND Estatu" _
    & "s <>'PAGADO';", cnnConexion, adOpenDynamic, adLockReadOnly, adCmdText
    HoraI = Time
'    Do While (IsNull(rstDeuda("Resta"))) And (DateAdd("S", 5, HoraI) > Time()) 'hace hasta que no sea nulo
'        rstDeuda.Requery
'    Loop
    '
    'If IsNull(rstDeuda("Resta")) Then rstDeuda("REsta") = 0
    cnnConexion.Execute "UPDATE Proveedores SET Deuda='" & IIf(IsNull(rstDeuda("Resta")), 0, rstDeuda("Resta")) & "' WHERE Codigo" _
    & "='" & Proveedor & "';"
    '
    rstDeuda.Close
    Set rstDeuda = Nothing
    '
    End Sub
    

    Private Function fecha_vcto() As Date
    '
    Dim Dia As Integer
    If Mtto Then
        Dia = Weekday("01/" & Format(Date, "mm/yyyy"))
        If Dia = 7 Then
            Dia = 6
        ElseIf Dia = 6 Then
            Dia = 1
        Else: Dia = Dia + (5 - Dia)
        End If
        fecha_vcto = DateAdd("d", 14, Dia & Format(Date, "/mm/yyyy"))
    Else
        MskFecha(0).PromptInclude = True
        fecha_vcto = DateAdd("D", 30, MskFecha(0))
        MskFecha(0).PromptInclude = False
    End If
    '
    End Function

    'se actualiza la lista de proveedores
    Public Sub Act_Lis()
    For I = 0 To 1: rstLisPro(I).Requery
    Next
    End Sub
    
    Private Function Campos_Requeridos() As Boolean
    'valida la captura de los datos
    Dim Cadena As String
    '
    For I = 0 To 2
        MskFecha(I).PromptInclude = True
    Next
    '
    If DatInmueble(0) = "" Then Cadena = "- Código del Inmueble" & vbCrLf
    If DatProveedor(0) = "" Then Cadena = Cadena & "- Código del Proveedor" & vbCrLf
    If CmbTipoPago(0) = "" Then Cadena = Cadena & "- Tipo de Documento" & vbCrLf
    If TxtFactura(1) = "" Then Cadena = Cadena & "- Descripción de la factura" & vbCrLf
    If Not IsDate(MskFecha(0)) Then Cadena = Cadena & "- Fecha de Recepción" & vbCrLf
    If Not IsDate(MskFecha(1)) Then Cadena = Cadena & "- Fecha de Emisión" & vbCrLf
    If Not IsDate(MskFecha(2)) Then Cadena = Cadena & "- Fecha de Vencimiento" & vbCrLf
    If CmbCargo = "" Then Cadena = Cadena & "- Tipo de moneda" & vbCrLf
    If TxtFactura(4) = "" Then Cadena = Cadena & "- Beneficiario" & vbCrLf
    If Not IsNumeric(TxtMonto(0)) Then Cadena = Cadena & "- Monto de la Factura" & vbCrLf
    If Not IsNumeric(lblFactura(29)) Then Cadena = Cadena & "- Monto Impuesto" & vbCrLf
    If Not IsNumeric(TxtMonto(1)) Then Cadena = Cadena & "- Monto Total de la factura"
    For I = 0 To 2
        MskFecha(I).PromptInclude = False
    Next
    '
    If Cadena <> "" Then
    
        Campos_Requeridos = MsgBox("No se puede procesar este documento." & vbCrLf _
        & vbCrLf & "Revise el contenido del(os) campo(s) siguiente(s):" & vbCrLf & Cadena, _
        vbCritical, App.ProductName)
        
    End If
    
    
    End Function
