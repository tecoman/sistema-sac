VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Begin VB.Form FrmProveedor 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ficha de Proveedor"
   ClientHeight    =   6120
   ClientLeft      =   1245
   ClientTop       =   1545
   ClientWidth     =   9630
   ControlBox      =   0   'False
   FillStyle       =   0  'Solid
   Icon            =   "FrmProveedor.frx":0000
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6120
   ScaleWidth      =   9630
   WindowState     =   2  'Maximized
   Begin ComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   480
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9630
      _ExtentX        =   16986
      _ExtentY        =   847
      ButtonWidth     =   714
      ButtonHeight    =   688
      Appearance      =   1
      ImageList       =   "ImageList1"
      _Version        =   327682
      BeginProperty Buttons {0713E452-850A-101B-AFC0-4210102A8DA7} 
         NumButtons      =   12
         BeginProperty Button1 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Top"
            Object.ToolTipText     =   "Primer Registro"
            Object.Tag             =   ""
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Back"
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
   Begin VB.Frame FrmProveedor 
      Enabled         =   0   'False
      Height          =   1620
      Index           =   0
      Left            =   345
      TabIndex        =   1
      Top             =   960
      Width           =   9045
      Begin VB.CheckBox ChkProveedor 
         Caption         =   "Inactivo"
         DataField       =   "Inactivo"
         DataSource      =   "rstProv(8)"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   1
         Left            =   7965
         TabIndex        =   7
         Top             =   1200
         Width           =   975
      End
      Begin VB.CheckBox ChkProveedor 
         Caption         =   "Nacional"
         DataField       =   "Nacional"
         DataSource      =   "rstProv(8)"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   0
         Left            =   7950
         TabIndex        =   6
         Top             =   720
         Width           =   975
      End
      Begin VB.TextBox TxtProveedor 
         DataField       =   "Codigo"
         DataSource      =   "rstProv(8)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   315
         Index           =   1
         Left            =   1065
         Locked          =   -1  'True
         TabIndex        =   5
         Top             =   255
         Width           =   960
      End
      Begin VB.TextBox TxtProveedor 
         DataField       =   "Rif"
         DataSource      =   "rstProv(8)"
         Height          =   315
         Index           =   2
         Left            =   1065
         TabIndex        =   4
         Top             =   705
         Width           =   1635
      End
      Begin VB.TextBox TxtProveedor 
         DataField       =   "NombProv"
         DataSource      =   "rstProv(8)"
         Height          =   315
         Index           =   0
         Left            =   4440
         TabIndex        =   3
         Top             =   240
         Width           =   4440
      End
      Begin VB.TextBox TxtProveedor 
         DataField       =   "Nit"
         DataSource      =   "rstProv(8)"
         Height          =   315
         Index           =   3
         Left            =   6015
         TabIndex        =   2
         Top             =   712
         Width           =   1635
      End
      Begin MSDataListLib.DataCombo CmbProveedor 
         Bindings        =   "FrmProveedor.frx":000C
         DataField       =   "Ramo"
         DataSource      =   "rstProv(8)"
         Height          =   315
         Index           =   0
         Left            =   1065
         TabIndex        =   8
         Top             =   1155
         Width           =   3195
         _ExtentX        =   5636
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   -2147483643
         ListField       =   "Nombre"
         BoundColumn     =   "Nombre"
         Text            =   ""
      End
      Begin MSMask.MaskEdBox MskFecha 
         Bindings        =   "FrmProveedor.frx":0023
         DataField       =   "FecReg"
         DataSource      =   "rstProv(8)"
         Height          =   315
         Left            =   6045
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   1185
         Width           =   1140
         _ExtentX        =   2011
         _ExtentY        =   556
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   10
         Format          =   "dd/MM/yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label LblProveedor 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   0
         Left            =   75
         TabIndex        =   14
         Top             =   300
         Width           =   900
      End
      Begin VB.Label LblProveedor 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   1
         Left            =   75
         TabIndex        =   13
         Top             =   745
         Width           =   900
      End
      Begin VB.Label LblProveedor 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   2
         Left            =   75
         TabIndex        =   12
         Top             =   1190
         Width           =   900
      End
      Begin VB.Label LblProveedor 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   3
         Left            =   3390
         TabIndex        =   11
         Top             =   270
         Width           =   900
      End
      Begin VB.Label LblProveedor 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   4
         Left            =   4995
         TabIndex        =   10
         Top             =   750
         Width           =   900
      End
      Begin VB.Label LblProveedor 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   5
         Left            =   4350
         TabIndex        =   9
         Top             =   1185
         Width           =   1575
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   5655
      Left            =   60
      TabIndex        =   16
      Top             =   570
      Width           =   9615
      _ExtentX        =   16960
      _ExtentY        =   9975
      _Version        =   393216
      Tabs            =   4
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
      TabPicture(0)   =   "FrmProveedor.frx":0045
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "FrmProveedor(1)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Datos Administrativos"
      TabPicture(1)   =   "FrmProveedor.frx":0061
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "FrmProveedor(2)"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Datos Adicionales"
      TabPicture(2)   =   "FrmProveedor.frx":007D
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "FrmProveedor(3)"
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "Lista"
      TabPicture(3)   =   "FrmProveedor.frx":0099
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "DataGrid1"
      Tab(3).Control(1)=   "FrmProveedor(5)"
      Tab(3).Control(2)=   "FrmProveedor(4)"
      Tab(3).ControlCount=   3
      Begin VB.Frame FrmProveedor 
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
         Height          =   3420
         Index           =   1
         Left            =   285
         TabIndex        =   44
         Top             =   1950
         Width           =   9075
         Begin VB.TextBox TxtProveedor 
            DataField       =   "Postal"
            DataSource      =   "rstProv(8)"
            Height          =   315
            Index           =   6
            Left            =   6030
            TabIndex        =   48
            Top             =   300
            Width           =   1200
         End
         Begin VB.TextBox TxtProveedor 
            DataField       =   "Email"
            DataSource      =   "rstProv(8)"
            Height          =   315
            Index           =   7
            Left            =   6030
            TabIndex        =   47
            Top             =   2865
            Width           =   2805
         End
         Begin VB.TextBox TxtProveedor 
            DataField       =   "Contacto"
            DataSource      =   "rstProv(8)"
            Height          =   315
            Index           =   4
            Left            =   1020
            TabIndex        =   46
            Top             =   813
            Width           =   3240
         End
         Begin VB.TextBox TxtProveedor 
            DataField       =   "Direccion"
            DataSource      =   "rstProv(8)"
            ForeColor       =   &H00404040&
            Height          =   1230
            Index           =   5
            Left            =   210
            MultiLine       =   -1  'True
            TabIndex        =   45
            Top             =   1965
            Width           =   4125
         End
         Begin MSMask.MaskEdBox MskTelefono 
            Bindings        =   "FrmProveedor.frx":00B5
            DataField       =   "Fax"
            DataSource      =   "rstProv(8)"
            Height          =   315
            Index           =   0
            Left            =   6015
            TabIndex        =   49
            Top             =   1830
            Width           =   2850
            _ExtentX        =   5027
            _ExtentY        =   556
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   16
            Format          =   "(0###)-###-##-##"
            Mask            =   "(####)-###-##-##"
            PromptChar      =   "_"
         End
         Begin MSDataListLib.DataCombo CmbProveedor 
            Bindings        =   "FrmProveedor.frx":00C0
            DataField       =   "Cargo"
            DataSource      =   "rstProv(8)"
            Height          =   315
            Index           =   2
            Left            =   1020
            TabIndex        =   50
            Top             =   1326
            Width           =   3255
            _ExtentX        =   5741
            _ExtentY        =   556
            _Version        =   393216
            BackColor       =   -2147483643
            ListField       =   "Nombre"
            BoundColumn     =   "Nombre"
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo CmbProveedor 
            Bindings        =   "FrmProveedor.frx":00D8
            DataField       =   "Actividad"
            DataSource      =   "rstProv(8)"
            Height          =   315
            Index           =   1
            Left            =   1020
            TabIndex        =   51
            Top             =   300
            Width           =   3240
            _ExtentX        =   5715
            _ExtentY        =   556
            _Version        =   393216
            BackColor       =   -2147483643
            ListField       =   "Nombre"
            BoundColumn     =   "Nombre"
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo CmbProveedor 
            Bindings        =   "FrmProveedor.frx":00F4
            DataField       =   "Ciudad"
            DataSource      =   "rstProv(8)"
            Height          =   315
            Index           =   3
            Left            =   6030
            TabIndex        =   52
            Top             =   813
            Width           =   2850
            _ExtentX        =   5027
            _ExtentY        =   556
            _Version        =   393216
            BackColor       =   -2147483643
            ListField       =   "Nombre"
            BoundColumn     =   ""
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo CmbProveedor 
            Bindings        =   "FrmProveedor.frx":010D
            DataField       =   "Estado"
            DataSource      =   "rstProv(8)"
            Height          =   315
            Index           =   4
            Left            =   6030
            TabIndex        =   53
            Top             =   1326
            Width           =   2850
            _ExtentX        =   5027
            _ExtentY        =   556
            _Version        =   393216
            BackColor       =   -2147483643
            ListField       =   "Nombre"
            BoundColumn     =   ""
            Text            =   ""
         End
         Begin MSMask.MaskEdBox MskTelefono 
            Bindings        =   "FrmProveedor.frx":0126
            DataField       =   "Telefonos"
            DataSource      =   "rstProv(8)"
            Height          =   315
            Index           =   1
            Left            =   6015
            TabIndex        =   54
            Top             =   2355
            Width           =   2850
            _ExtentX        =   5027
            _ExtentY        =   556
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   16
            Format          =   "###########"
            Mask            =   "(####)-###-##-##"
            PromptChar      =   "_"
         End
         Begin VB.Label LblProveedor 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   15
            Left            =   5025
            TabIndex        =   64
            Top             =   2895
            Width           =   900
         End
         Begin VB.Label LblProveedor 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   14
            Left            =   4725
            TabIndex        =   63
            Top             =   2385
            Width           =   1200
         End
         Begin VB.Label LblProveedor 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   13
            Left            =   5025
            TabIndex        =   62
            Top             =   1869
            Width           =   900
         End
         Begin VB.Label LblProveedor 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   12
            Left            =   5025
            TabIndex        =   61
            Top             =   1356
            Width           =   900
         End
         Begin VB.Label LblProveedor 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   11
            Left            =   5025
            TabIndex        =   60
            Top             =   843
            Width           =   900
         End
         Begin VB.Label LblProveedor 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   10
            Left            =   5025
            TabIndex        =   59
            Top             =   330
            Width           =   900
         End
         Begin VB.Label LblProveedor 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   9
            Left            =   75
            TabIndex        =   58
            Top             =   1740
            Width           =   900
         End
         Begin VB.Label LblProveedor 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   8
            Left            =   75
            TabIndex        =   57
            Top             =   1356
            Width           =   900
         End
         Begin VB.Label LblProveedor 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   7
            Left            =   75
            TabIndex        =   56
            Top             =   843
            Width           =   900
         End
         Begin VB.Label LblProveedor 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   6
            Left            =   75
            TabIndex        =   55
            Top             =   330
            Width           =   900
         End
      End
      Begin MSDataGridLib.DataGrid DataGrid1 
         Bindings        =   "FrmProveedor.frx":0131
         Height          =   2940
         Left            =   -74715
         TabIndex        =   82
         Top             =   600
         Visible         =   0   'False
         Width           =   9045
         _ExtentX        =   15954
         _ExtentY        =   5186
         _Version        =   393216
         AllowUpdate     =   0   'False
         AllowArrows     =   -1  'True
         HeadLines       =   1
         RowHeight       =   15
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
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Listado de Proveedores"
         ColumnCount     =   5
         BeginProperty Column00 
            DataField       =   "Codigo"
            Caption         =   "Código"
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
            DataField       =   "NombProv"
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
         BeginProperty Column02 
            DataField       =   "Ramo"
            Caption         =   "Ramo"
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
            DataField       =   "Contacto"
            Caption         =   "Contacto"
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
            DataField       =   "Telefonos"
            Caption         =   "Teléfono"
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
               ColumnWidth     =   675,213
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   3044,977
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   1470,047
            EndProperty
            BeginProperty Column03 
               ColumnWidth     =   1725,165
            EndProperty
            BeginProperty Column04 
            EndProperty
         EndProperty
      End
      Begin VB.Frame FrmProveedor 
         Enabled         =   0   'False
         Height          =   3420
         Index           =   3
         Left            =   -74715
         TabIndex        =   65
         Top             =   1950
         Width           =   9075
         Begin VB.TextBox TxtProveedor 
            DataField       =   "Campo8"
            DataSource      =   "rstProv(8)"
            Height          =   315
            Index           =   22
            Left            =   6210
            TabIndex        =   73
            Top             =   2235
            Width           =   2565
         End
         Begin VB.TextBox TxtProveedor 
            DataField       =   "Campo7"
            DataSource      =   "rstProv(8)"
            Height          =   315
            Index           =   21
            Left            =   6210
            TabIndex        =   72
            Top             =   1605
            Width           =   2565
         End
         Begin VB.TextBox TxtProveedor 
            DataField       =   "Campo6"
            DataSource      =   "rstProv(8)"
            Height          =   315
            Index           =   20
            Left            =   6195
            TabIndex        =   71
            Top             =   975
            Width           =   2565
         End
         Begin VB.TextBox TxtProveedor 
            DataField       =   "Campo5"
            DataSource      =   "rstProv(8)"
            Height          =   315
            Index           =   19
            Left            =   6195
            TabIndex        =   70
            Top             =   360
            Width           =   2565
         End
         Begin VB.TextBox TxtProveedor 
            DataField       =   "Campo4"
            DataSource      =   "rstProv(8)"
            Height          =   315
            Index           =   18
            Left            =   1395
            TabIndex        =   69
            Top             =   2235
            Width           =   2565
         End
         Begin VB.TextBox TxtProveedor 
            DataField       =   "Campo3"
            DataSource      =   "rstProv(8)"
            Height          =   315
            Index           =   17
            Left            =   1395
            TabIndex        =   68
            Top             =   1605
            Width           =   2565
         End
         Begin VB.TextBox TxtProveedor 
            DataField       =   "Campo2"
            DataSource      =   "rstProv(8)"
            Height          =   315
            Index           =   16
            Left            =   1410
            TabIndex        =   67
            Top             =   975
            Width           =   2565
         End
         Begin VB.TextBox TxtProveedor 
            DataField       =   "Campo1"
            DataSource      =   "rstProv(8)"
            Height          =   315
            Index           =   15
            Left            =   1410
            TabIndex        =   66
            Top             =   360
            Width           =   2565
         End
         Begin VB.Label Label1 
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
            Index           =   0
            Left            =   5070
            TabIndex        =   81
            Top             =   1005
            Width           =   810
         End
         Begin VB.Label Label1 
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
            Index           =   1
            Left            =   345
            TabIndex        =   80
            Top             =   390
            Width           =   750
         End
         Begin VB.Label Label1 
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
            Index           =   2
            Left            =   315
            TabIndex        =   79
            Top             =   1005
            Width           =   780
         End
         Begin VB.Label Label1 
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
            Index           =   3
            Left            =   315
            TabIndex        =   78
            Top             =   1635
            Width           =   765
         End
         Begin VB.Label Label1 
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
            Index           =   4
            Left            =   315
            TabIndex        =   77
            Top             =   2265
            Width           =   810
         End
         Begin VB.Label Label1 
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
            Index           =   5
            Left            =   5070
            TabIndex        =   76
            Top             =   390
            Width           =   735
         End
         Begin VB.Label Label1 
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
            Index           =   6
            Left            =   5115
            TabIndex        =   75
            Top             =   1635
            Width           =   810
         End
         Begin VB.Label Label1 
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
            Index           =   7
            Left            =   5100
            TabIndex        =   74
            Top             =   2265
            Width           =   825
         End
      End
      Begin VB.Frame FrmProveedor 
         BackColor       =   &H80000004&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1335
         Index           =   5
         Left            =   -71865
         TabIndex        =   39
         Top             =   3810
         Width           =   4365
         Begin VB.CommandButton BotBusca 
            Height          =   330
            Left            =   3765
            Picture         =   "FrmProveedor.frx":014A
            Style           =   1  'Graphical
            TabIndex        =   41
            ToolTipText     =   "Buscar"
            Top             =   750
            Width           =   405
         End
         Begin VB.TextBox TxtProveedor 
            BackColor       =   &H00FFFFFF&
            Height          =   285
            Index           =   23
            Left            =   1005
            TabIndex        =   40
            Top             =   285
            Width           =   3165
         End
         Begin VB.Label Label2 
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   1
            Left            =   1005
            TabIndex        =   43
            Top             =   765
            Width           =   2745
         End
         Begin VB.Label Label2 
            Caption         =   "Buscar:"
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
            Index           =   0
            Left            =   165
            TabIndex        =   42
            Top             =   315
            Width           =   630
         End
      End
      Begin VB.Frame FrmProveedor 
         BackColor       =   &H80000004&
         Caption         =   "Buscar Por:"
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
         Height          =   1320
         Index           =   4
         Left            =   -74730
         TabIndex        =   34
         Top             =   3795
         Width           =   2775
         Begin VB.OptionButton OptBusca 
            BackColor       =   &H80000004&
            Caption         =   "Código"
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
            Left            =   120
            TabIndex        =   38
            Tag             =   "Codigo, NombProv"
            Top             =   390
            Value           =   -1  'True
            Width           =   1260
         End
         Begin VB.OptionButton OptBusca 
            BackColor       =   &H80000004&
            Caption         =   "Contacto"
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
            Index           =   2
            Left            =   105
            TabIndex        =   37
            Tag             =   "Contacto, Codigo, NombProv"
            Top             =   825
            Width           =   1260
         End
         Begin VB.OptionButton OptBusca 
            BackColor       =   &H80000004&
            Caption         =   "Nombre"
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
            Left            =   1380
            TabIndex        =   36
            Tag             =   "NombProv, Codigo"
            Top             =   375
            Width           =   1245
         End
         Begin VB.OptionButton OptBusca 
            BackColor       =   &H80000004&
            Caption         =   "Ramo"
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
            Index           =   3
            Left            =   1365
            TabIndex        =   35
            Tag             =   "Ramo, Codigo, NombProv"
            Top             =   825
            Width           =   1230
         End
      End
      Begin VB.Frame FrmProveedor 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3420
         Index           =   2
         Left            =   -74730
         TabIndex        =   17
         Top             =   1950
         Width           =   9075
         Begin VB.TextBox TxtProveedor 
            Alignment       =   1  'Right Justify
            DataField       =   "Limite"
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "#,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   8202
               SubFormatType   =   0
            EndProperty
            DataSource      =   "rstProv(8)"
            Height          =   315
            Index           =   8
            Left            =   2010
            Locked          =   -1  'True
            TabIndex        =   24
            Top             =   765
            Width           =   1605
         End
         Begin VB.TextBox TxtProveedor 
            Alignment       =   1  'Right Justify
            DataField       =   "Deuda"
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "#,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   8202
               SubFormatType   =   0
            EndProperty
            DataSource      =   "rstProv(8)"
            Height          =   315
            Index           =   11
            Left            =   6900
            Locked          =   -1  'True
            TabIndex        =   23
            Top             =   705
            Width           =   1620
         End
         Begin VB.TextBox TxtProveedor 
            DataField       =   "Beneficiario"
            DataSource      =   "rstProv(8)"
            Height          =   315
            Index           =   13
            Left            =   2010
            TabIndex        =   22
            Top             =   1965
            Width           =   4440
         End
         Begin VB.TextBox TxtProveedor 
            DataField       =   "Notas"
            DataSource      =   "rstProv(8)"
            ForeColor       =   &H00808080&
            Height          =   735
            Index           =   14
            Left            =   2010
            MultiLine       =   -1  'True
            TabIndex        =   21
            Top             =   2505
            Width           =   6525
         End
         Begin VB.TextBox TxtProveedor 
            Alignment       =   1  'Right Justify
            DataField       =   "FecUltPag"
            DataSource      =   "rstProv(8)"
            Height          =   315
            Index           =   12
            Left            =   6900
            Locked          =   -1  'True
            TabIndex        =   20
            Top             =   1110
            Width           =   1620
         End
         Begin VB.TextBox TxtProveedor 
            Alignment       =   1  'Right Justify
            DataField       =   "UltPago"
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "#,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   8202
               SubFormatType   =   0
            EndProperty
            DataSource      =   "rstProv(8)"
            Height          =   315
            Index           =   9
            Left            =   2010
            Locked          =   -1  'True
            TabIndex        =   19
            Top             =   1185
            Width           =   1605
         End
         Begin VB.TextBox TxtProveedor 
            Alignment       =   1  'Right Justify
            DataField       =   "DiasCred"
            DataSource      =   "rstProv(8)"
            Height          =   315
            Index           =   10
            Left            =   6900
            Locked          =   -1  'True
            TabIndex        =   18
            Top             =   315
            Width           =   750
         End
         Begin MSDataListLib.DataCombo CmbProveedor 
            Bindings        =   "FrmProveedor.frx":06D4
            DataField       =   "Condicion"
            DataSource      =   "rstProv(8)"
            Height          =   315
            Index           =   5
            Left            =   2010
            TabIndex        =   25
            Top             =   345
            Width           =   1605
            _ExtentX        =   2831
            _ExtentY        =   556
            _Version        =   393216
            BackColor       =   -2147483643
            ListField       =   "Nombre"
            BoundColumn     =   "Nombre"
            Text            =   ""
         End
         Begin VB.Line Line2 
            BorderColor     =   &H00E0E0E0&
            Index           =   1
            X1              =   60
            X2              =   9045
            Y1              =   1785
            Y2              =   1785
         End
         Begin VB.Label LblProveedor 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   23
            Left            =   5010
            TabIndex        =   33
            Top             =   1140
            Width           =   1600
         End
         Begin VB.Label LblProveedor 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   22
            Left            =   5010
            TabIndex        =   32
            Top             =   735
            Width           =   1600
         End
         Begin VB.Label LblProveedor 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   21
            Left            =   5010
            TabIndex        =   31
            Top             =   345
            Width           =   1600
         End
         Begin VB.Label LblProveedor 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   20
            Left            =   150
            TabIndex        =   30
            Top             =   2505
            Width           =   1605
         End
         Begin VB.Label LblProveedor 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   19
            Left            =   150
            TabIndex        =   29
            Top             =   1995
            Width           =   1605
         End
         Begin VB.Label LblProveedor 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   18
            Left            =   150
            TabIndex        =   28
            Top             =   1215
            Width           =   1605
         End
         Begin VB.Label LblProveedor 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   17
            Left            =   150
            TabIndex        =   27
            Top             =   795
            Width           =   1605
         End
         Begin VB.Label LblProveedor 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   16
            Left            =   150
            TabIndex        =   26
            Top             =   375
            Width           =   1605
         End
         Begin VB.Line Line2 
            BorderColor     =   &H00808080&
            BorderWidth     =   2
            Index           =   0
            X1              =   60
            X2              =   9060
            Y1              =   1785
            Y2              =   1785
         End
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
            Picture         =   "FrmProveedor.frx":06DF
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmProveedor.frx":0861
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmProveedor.frx":09E3
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmProveedor.frx":0B65
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmProveedor.frx":0CE7
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmProveedor.frx":0E69
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmProveedor.frx":0FEB
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmProveedor.frx":116D
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmProveedor.frx":12EF
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmProveedor.frx":1471
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmProveedor.frx":15F3
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmProveedor.frx":1775
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "FrmProveedor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
    '---------------------------------------------------------------------------------------------
    '   SINAI TECH, C.A 11/2002
    '   Registro y edición de Proveedores
    '---------------------------------------------------------------------------------------------
    Dim CnTabla As ADODB.Connection
    Dim mcNro As String
    Dim i As Integer
    Dim rstProv(8) As ADODB.Recordset
    Enum ado
        Ramo
        Actividad
        Cargo
        Ciudad
        Estado
        Condicion
        Pro
        CodPro
    End Enum
    '---------------------------------------------------------------------------------------------

    Private Sub BotBusca_Click()
    '
        With rstProv(8)
            .MoveFirst
            If OptBusca(0).Value = True Then
                .Find "Codigo = '" & Format(TxtProveedor(23), "0000") & "'"
            End If
            If OptBusca(1).Value = True Then
                .Find "NombProv LIKE '%" & TxtProveedor(23) & "%'"
            End If
            If OptBusca(2).Value = True Then
                .Find "Contacto LIKE '%" & TxtProveedor(23) & "%'"
            End If
            If OptBusca(3).Value = True Then
                .Find "Ramo likE '%" & TxtProveedor(23) & "%'"
            End If
                
        If .EOF Then
        '
            MsgBox "No encontré el proveedor " & TxtBus, vbExclamation, "Busqueda de Proveedor"
            .MoveFirst
            With TxtProveedor(23)
                .SetFocus
                .SelStart = 0
                .SelLength = Len(TxtProveedor(23))
            End With
        
        End If
        '
        End With
        '
    End Sub

    Private Sub CmbProveedor_KeyPress(Index%, KeyAscii%)
    'convierte todo a ma
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    '
    If KeyAscii = 13 Then
    '
        Select Case Index
            
            Case 0: CmbProveedor(inex + 1).SetFocus
            '
            Case 1: TxtProveedor(4).SetFocus
            '
            Case 2: TxtProveedor(5).SetFocus
            '
            Case 3: CmbProveedor(Index + 1).SetFocus
            '
            Case 4: mskTelefono(0).SetFocus
            '
        End Select
        '
    End If
    '
    End Sub

    '---------------------------------------------------------------------------------------------
    Private Sub Form_Resize()   '
    '---------------------------------------------------------------------------------------------
    'configura la presentacion de los controles en pantalla
    '
    With SSTab1
        .Width = Me.ScaleWidth - .Left
        .Height = Me.ScaleHeight - .Top
        FrmProveedor(3).Left = (.Width - FrmProveedor(3).Width) / 2
        FrmProveedor(2).Left = (.Width - FrmProveedor(2).Width) / 2
        FrmProveedor(1).Left = (.Width - FrmProveedor(1).Width) / 2
        FrmProveedor(0).Left = .Left + FrmProveedor(1).Left
        DataGrid1.Left = 285
        DataGrid1.Width = .Width - (DataGrid1.Left * 2)
        
        FrmProveedor(4).Top = .Height - (FrmProveedor(4).Height) - 200
        FrmProveedor(5).Top = .Height - (FrmProveedor(4).Height) - 200
        DataGrid1.Height = FrmProveedor(4).Top - DataGrid1.Top - 200
    End With
    With DataGrid1
        .Columns(0).Width = .Width * 5 / 100
        .Columns(1).Width = .Width * 20 / 100
        .Columns(2).Width = .Width * 20 / 100
        .Columns(3).Width = .Width * 20 / 100
        .Columns(4).Width = .Width * 30 / 100
    End With
    
    End Sub

    '---------------------------------------------------------------------------------------------
    Private Sub Form_Unload(Cancel As Integer)  '
    'Rutina Desctructor---------------------------------------------------------------------------
    For i = 0 To 8  'cierra y destruye conexiones y ADODB.Recordset's
        If rstProv(i).State = 1 Then rstProv(i).Close
        Set rstProv(i) = Nothing
    Next
    CnTabla.Close: Set CnTabla = Nothing
    '
    End Sub

    Private Sub MskTelefono_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 And Index = 0 Then mskTelefono(Index + 1).SetFocus
    If KeyAscii = 13 And Index = 1 Then TxtProveedor(7).SetFocus
    '    Select Case Index
    '        Case 0
    '
    '        Case 1
    '            TxtProveedor(7).SetFocus
    '    End Select
    'End If
    
    End Sub

    Private Sub OptBusca_Click(Index As Integer): rstProv(8).Sort = OptBusca(Index).Tag
    End Sub

    Private Sub SSTab1_Click(PreviousTab As Integer)
        
        Select Case SSTab1.Tab
        
            Case 0, 1, 2: FrmProveedor(0).Visible = True: DataGrid1.Visible = False
            
            Case 3
                FrmProveedor(0).Visible = False
                DataGrid1.Visible = True
                Label2(1) = "TOTAL REGISTROS: " & rstProv(8).RecordCount
            
        End Select
    
    End Sub
    
   '----------------------------------------------------------------------------------------------
   Private Sub Form_Load()  '
   '----------------------------------------------------------------------------------------------
   'variables locales
   Dim ctl As Control
   LblProveedor(0) = LoadResString(128): LblProveedor(1) = LoadResString(129)
   LblProveedor(2) = LoadResString(131): LblProveedor(3) = LoadResString(124)
   LblProveedor(4) = LoadResString(130): LblProveedor(5) = LoadResString(132)
   LblProveedor(6) = LoadResString(133): LblProveedor(7) = LoadResString(135)
   LblProveedor(8) = LoadResString(137): LblProveedor(9) = LoadResString(139)
   LblProveedor(10) = LoadResString(134): LblProveedor(11) = LoadResString(136)
   LblProveedor(12) = LoadResString(138): LblProveedor(13) = LoadResString(113)
   LblProveedor(14) = LoadResString(109): LblProveedor(15) = LoadResString(114)
   LblProveedor(16) = LoadResString(140): LblProveedor(17) = LoadResString(142)
   LblProveedor(18) = LoadResString(115): LblProveedor(19) = LoadResString(144)
   LblProveedor(20) = LoadResString(118): LblProveedor(21) = LoadResString(141)
   LblProveedor(22) = LoadResString(143): LblProveedor(23) = LoadResString(111)
   Set DataGrid1.HeadFont = LetraTitulo(LoadResString(527), 7.5, True)
   Set DataGrid1.Font = LetraTitulo(LoadResString(528), 8)
    
   '
   'CONFIGURA EL ORIGEN DE DATOS
    For i = 0 To 8: Set rstProv(i) = New ADODB.Recordset
    Next
    Set CnTabla = New ADODB.Connection
    'Configura Ado
    CnTabla.CursorLocation = adUseClient
    CnTabla.Open cnnOLEDB + gcPath + "\Tablas.mdb"
    '
    rstProv(ado.Actividad).Open "SELECT * FROM Actividad ORDER BY Nombre", CnTabla, _
    adOpenForwardOnly, adLockOptimistic, adCmdText
    '
    rstProv(ado.Ramo).Open "SELECT * FROM Ramo ORDER BY Nombre", CnTabla, _
    adOpenForwardOnly, adLockOptimistic, adCmdText
    '
    rstProv(ado.Ciudad).Open "SELECT * FROM Ciudades ORDER BY Nombre", CnTabla, _
    adOpenForwardOnly, adLockOptimistic, adCmdText
    '
    rstProv(ado.Estado).Open "SELECT * FROM Estados ORDER BY Nombre", CnTabla, _
    adOpenForwardOnly, adLockOptimistic, adCmdText
    '
    rstProv(ado.Cargo).Open "SELECT * FROM Cargos ORDER BY Nombre", CnTabla, _
    adOpenForwardOnly, adLockOptimistic, adCmdText
    '
    rstProv(ado.Condicion).Open "SELECT * FROM Condicion ORDER BY Nombre", CnTabla, _
    adOpenForwardOnly, adLockOptimistic, adCmdText
    'rstProv(Ado.Pro).Open "SELECT * FROM Proveedores ORDER BY Codigo", cnnConexion, _
    adOpenStatic, adLockReadOnly
   'DEFINE LOR ORIGENES DE LOS ELEMENTOS DE LAS LISTAS DE LOS OBJETOS COMBO
   
    For i = O To 5: Set CmbProveedor(i).RowSource = rstProv(i)
    Next
    '------
    SSTab1.Tab = 0
    rstProv(8).Open "Proveedores", cnnConexion, adOpenKeyset, adLockOptimistic, adCmdTable
    'asigna la propiedad datasource a los controles
    'On Error Resume Next
    For Each ctl In Controls
        If TypeName(ctl) = "TextBox" Or TypeName(ctl) = "DataCombo" Or TypeName(ctl) = _
        "DataGrid" Or TypeName(ctl) = "CheckBox" Or TypeName(ctl) = "MaskEdBox" Then
            Set ctl.DataSource = rstProv(8)
        End If
    Next
    
    End Sub

    Private Sub Toolbar1_ButtonClick(ByVal Button As Button)
    'variables locales
    Dim Formulario As Form                  '*
    Dim Registro As Long                    '*
    Dim codProA$, codProN, Nombre$          '*
    Dim rstPro As New ADODB.Recordset       '*
    Dim Intento As Integer                  '*
    '---------------------------------------'*
    With rstProv(8)
    '
        Select Case Button.Key
    '
            Case Is = "Top"     'IR AL REGISTRO INICIAL
    '
                .MoveFirst
    '
            Case Is = "Back"    'REGISTRO ANTERIOR
    '
                .MovePrevious
                If .BOF Then .MoveLast
    '
            Case Is = "Next"    'IR LA SIGUIENTE REGISTRO
    '
                .MoveNext
                If .EOF Then .MoveFirst
    '
            Case Is = "End" 'IR LA ULTIMO REGISTRO
    '
                .MoveLast
                
    Case Is = "New"           ' Nuevo Proveedor.
        SSTab1.Tab = 0
        For i = 0 To 3: FrmProveedor(i).Enabled = True
        Next i
        Set DataGrid1.DataSource = Nothing
        MskFecha.PromptInclude = True
        .AddNew
        MskFecha = Date
        
        rstProv(ado.CodPro).Open "SELECT Codigo FROM Proveedores ORDER BY Codigo", _
        cnnConexion, adOpenKeyset, adLockBatchOptimistic, adCmdText
        Registro = 1
        With rstProv(ado.CodPro)
            .MoveFirst
10        If CCur(!Codigo) = Registro Then
                Registro = Registro + 1
                .MoveNext
                If .EOF Then
                    TxtProveedor(1) = Format(Registro, "0000")
                Else
                    GoTo 10
                End If
            Else
                TxtProveedor(1) = Format(Registro, "0000")
            End If
            .Close
        End With
        TxtProveedor(0).SetFocus
        Call RtnEstado(Button.Index, Toolbar1)
        
    Case Is = "Save"    'ACTUALIZAR REGISTRO
            
        For i = 0 To 3: FrmProveedor(i).Enabled = False
        Next i
        If Not Valida_guardar Then
            Call RtnEstado(Button.Index, Toolbar1)
            !Usuario = gcUsuario
            MskFecha.PromptInclude = True
            .Update
            MskFecha.PromptInclude = False
            For Each Formulario In Forms
                If Formulario.Name = "FrmFactura" Then
                    'actualiza la lista de proveedores en el formulario de recepción
                    'de facturas
                    FrmFactura.Act_Lis
                    Exit For
                End If
            Next
            MsgBox " Registro Actualizado ... ", vbInformation, App.ProductName
            Set DataGrid1.DataSource = rstProv(8)
        End If
        '
    Case Is = "Find"              ' Buscar.
        SSTab1.Tab = 3

    Case Is = "Undo"    'CANCELAR REGISTRO
        Call RtnEstado(Button.Index, Toolbar1)
        
        For i = 0 To 3: FrmProveedor(i).Enabled = False
        Next i
        .CancelUpdate
        If .RecordCount <= 0 Then Exit Sub
        .MoveFirst
        MsgBox " Registro Cancelado ... ", vbInformation, App.ProductName
        Set DataGrid1.DataSource = rstProv(8)
        '
    Case Is = "Delete"
        
        On Error GoTo rtnDelete
        'Set DataGrid1.DataSource = Nothing
        codProA = !Codigo: Nombre = IIf(IsNull(!NombProv), "", !NombProv)
        If Respuesta(LoadResString(526) & vbCrLf & codProA & " - " & Nombre) Then
            Registro = .AbsolutePosition
            .Delete
            .Requery
            MsgBox " Registro Eliminado ... ", vbInformation, App.ProductName
            Call rtnBitacora("Eliminar Proveedor #" & codProA & "/" & Nombre)
        End If
        'Set DataGrid1.DataSource = rstProv(8)
        
rtnDelete:
    If Err.Number <> 0 Then 'si ocurre algún error
        '
        Call rtnBitacora("Error al Eliminar Proveedor Error #" & Err)
        If Err = 3200 Or Err = -2147467259 Then

20        codProN = InputBox$("Ingrese el Código del Proveedor que sustituye a: " & codProA & _
            " - " & Nombre, "Código Proveedor", codProA)
            If codProN = "" Then
                MsgBox "Registro NO Eliminado", vbInformation, App.ProductName
                Exit Sub
            Else
                '
                rstPro.Open "SELECT * FROM Proveedores WHERE Codigo='" & codProN & "'", _
                cnnConexion, adOpenStatic, adLockReadOnly, adCmdText
                If rstPro.EOF Or rstPro.BOF Then
                    MsgBox "El Código de proveedor [" & codProN & "] no corresponde a ningún pr" _
                    & "oveedor en nuestros registros. Por favor introduzcalo nuevamente...", _
                    vbInformation, App.ProductName
                    rstPro.Close
                    Intento = Intento + 1
                    If Intento = 3 Then
                        MsgBox "Lo ha intentado '" & Intento & "' veces, La transacción ha sido" _
                        & " cancelada...", vbInformation, App.ProductName
                        Exit Sub
                    End If
                    GoTo 20
                End If
                rstPro.Close
                Set rstPro = Nothing
            End If
            cnnConexion.Execute "UPDATE Cpp SET CodProv='" & codProN & "' WHERE CodProv='" _
            & codProA & "'"
            Call rtnBitacora("Actualizado Cpp " & codProA & "/" & codProN)
            For K = 0 To 100000  'retardo
            Next
            .Update
            
            Call rtnBitacora("Eliminar Proveedor #" & codProA & "/" & Nombre)
        Else
            MsgBox Err.Description, vbCritical, "Error " & App.ProductName
        End If
    End If
    
    Case Is = "Close"
        Unload Me                  ' Cerrar y Salir

    Case Is = "Edit1"
        For i = 0 To 3: FrmProveedor(i).Enabled = True
        Next i
        Call RtnEstado(Button.Index, Toolbar1)

    Case Is = "Print"              ' Imprimir
        mcTitulo = "Listado de Proveedores"
        mcReport = "LisProv.Rpt"
        mcOrdCod = "+{Proveedores.Codigo}"
        mcOrdAlfa = "+{Proveedores.NombProv}"
        mcCrit = ""
        FrmReport.Show
        '
    End Select
    
End With
'
End Sub

    Private Sub TxtProveedor_KeyPress(Index As Integer, KeyAscii As Integer)
    If Index = 7 Then
        KeyAscii = Asc(LCase(Chr(KeyAscii)))
    Else
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
    End If
    If KeyAscii = 13 Then
    Select Case Index
        
        Case 0  'TEX NOMBRE PROVEEDOR
            'BUSCA SI YA FUE PROCESADO ESTE PROVEEDOR
            If Valida_guardar Then
                    With TxtProveedor(0)
                        .SelStart = 0
                        .SelLength = Len(.Text)
                    End With
            End If
    '        With rstProv(Ado.Pro)
    '            .MoveFirst
    '            .Find "NombProv = '" & TxtProveedor(0) & "'"
    '            If .EOF Then
    '                TxtProveedor(Index + 2).SetFocus
    '            Else
    '                MsgBox "Proveedor YA CREADO con el Codigo: '" & .Fields("Codigo") & " '", _
    '                vbExclamation + vbCritical
    '                With TxtProveedor(0)
    ''                    .SetFocus
    '                    .SelStart = 0
    '                    .SelLength = Len(.Text)
    '                End With
    '
    '            End If
    '        End With
            
        Case 2: TxtProveedor(Index + 1).SetFocus
        
        Case 3: CmbProveedor(0).SetFocus
        
        Case 4: CmbProveedor(2).SetFocus
        
        Case 23: BotBusca.SetFocus
        
    End Select
    
    End If

End Sub

'-------------------------------------------------------------------------------------------------
'   Funcion:    Valida_Guardar
'
'   si se esta agregando un nuevo registro, verifica que que el proveedor no esta
'   registrado, si lo esta devuelve TRUE sino FALSE
'-------------------------------------------------------------------------------------------------
Private Function Valida_guardar() As Boolean
'
If rstProv(8).EditMode = adEditAdd Then
    With rstProv(ado.Pro)
        .Open "Proveedores", cnnConexion, adOpenStatic, adLockReadOnly, adCmdTable
        .Filter = "NombProv='" & TxtProveedor(0) & "'"
        If Not .EOF Then
            Valida_guardar = MsgBox("Proveedor ya registrado con el Código: " & !Codigo, _
            vbInformation, App.EXEName)
        End If
        .Close
    End With
End If
End Function
