VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmInmueble 
   Caption         =   "Ficha de Inmuebles"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   ControlBox      =   0   'False
   Icon            =   "FrmInmueble.frx":0000
   MDIChild        =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   WindowState     =   2  'Maximized
   Begin ComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   480
      Left            =   0
      TabIndex        =   11
      Top             =   0
      Width           =   4680
      _ExtentX        =   8255
      _ExtentY        =   847
      ButtonWidth     =   714
      ButtonHeight    =   688
      AllowCustomize  =   0   'False
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
      MousePointer    =   99
      MouseIcon       =   "FrmInmueble.frx":000C
   End
   Begin VB.Frame fraInm 
      Enabled         =   0   'False
      Height          =   1620
      Index           =   0
      Left            =   300
      TabIndex        =   25
      Top             =   1020
      Width           =   9165
      Begin VB.TextBox txtInm 
         DataField       =   "nSSO"
         ForeColor       =   &H00000000&
         Height          =   315
         Index           =   28
         Left            =   6285
         MaxLength       =   12
         TabIndex        =   115
         Top             =   285
         Width           =   1380
      End
      Begin VB.TextBox txtInm 
         DataField       =   "RIF"
         ForeColor       =   &H00000000&
         Height          =   315
         Index           =   27
         Left            =   2850
         MaxLength       =   12
         TabIndex        =   114
         Top             =   315
         Width           =   1485
      End
      Begin VB.CheckBox chkADM 
         Alignment       =   1  'Right Justify
         Caption         =   "Administradora SAC:"
         DataField       =   "SAC"
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
         Left            =   3465
         TabIndex        =   101
         Top             =   1170
         Width           =   2025
      End
      Begin VB.CommandButton cmdInm 
         Height          =   255
         Index           =   0
         Left            =   8595
         Picture         =   "FrmInmueble.frx":0326
         Style           =   1  'Graphical
         TabIndex        =   35
         Top             =   1140
         Visible         =   0   'False
         Width           =   250
      End
      Begin VB.TextBox txtInm 
         DataField       =   "CodInm"
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
         Index           =   0
         Left            =   1095
         TabIndex        =   28
         Top             =   315
         Width           =   960
      End
      Begin VB.TextBox txtInm 
         DataField       =   "Nombre"
         Height          =   315
         Index           =   1
         Left            =   1095
         TabIndex        =   27
         Top             =   705
         Width           =   7755
      End
      Begin VB.CheckBox cnkInm 
         Caption         =   "Inactivo"
         DataField       =   "Inactivo"
         DataSource      =   "AdoInmueble"
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
         Left            =   7935
         TabIndex        =   26
         Top             =   315
         Width           =   975
      End
      Begin MSDataListLib.DataCombo cmbInm 
         DataField       =   "TipoInm"
         Height          =   315
         Index           =   0
         Left            =   1095
         TabIndex        =   29
         Top             =   1110
         Width           =   2025
         _ExtentX        =   3572
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   -2147483643
         ListField       =   "Nombre"
         BoundColumn     =   "Nombre"
         Text            =   ""
      End
      Begin MSMask.MaskEdBox mskInm 
         Bindings        =   "FrmInmueble.frx":0470
         DataField       =   "FecIni"
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
         Index           =   0
         Left            =   7575
         TabIndex        =   34
         TabStop         =   0   'False
         ToolTipText     =   "Documento Fecha 1"
         Top             =   1110
         Width           =   1290
         _ExtentX        =   2275
         _ExtentY        =   556
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   10
         Format          =   "dd/MM/yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label lblInm 
         Caption         =   "Nº Empresa SSO:"
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
         Index           =   45
         Left            =   4800
         TabIndex        =   116
         Top             =   330
         Width           =   1530
      End
      Begin VB.Label lblInm 
         Caption         =   "RIF.:"
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
         Index           =   43
         Left            =   2310
         TabIndex        =   113
         Top             =   360
         Width           =   840
      End
      Begin VB.Label lblInm 
         Caption         =   "Codigo:"
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
         Left            =   180
         TabIndex        =   33
         Top             =   367
         Width           =   840
      End
      Begin VB.Label lblInm 
         Caption         =   "Inmueble"
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
         Left            =   150
         TabIndex        =   32
         Top             =   750
         Width           =   840
      End
      Begin VB.Label lblInm 
         Caption         =   "Tipo:"
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
         Left            =   180
         TabIndex        =   31
         Top             =   1155
         Width           =   765
      End
      Begin VB.Label lblInm 
         Caption         =   "Fecha de Registro"
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
         Left            =   5970
         TabIndex        =   30
         Top             =   1155
         Width           =   1530
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   7140
      Left            =   105
      TabIndex        =   5
      Top             =   600
      Width           =   9600
      _ExtentX        =   16933
      _ExtentY        =   12594
      _Version        =   393216
      Tabs            =   5
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
      TabCaption(0)   =   "Datos Generales"
      TabPicture(0)   =   "FrmInmueble.frx":0492
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "fraInm(3)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Administrativos"
      TabPicture(1)   =   "FrmInmueble.frx":04AE
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "fraInm(1)"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Servicios Básicos"
      TabPicture(2)   =   "FrmInmueble.frx":04CA
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "fraInm(6)"
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "Junta de Condominio"
      TabPicture(3)   =   "FrmInmueble.frx":04E6
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "FraJunta"
      Tab(3).Control(1)=   "fraInm(2)"
      Tab(3).ControlCount=   2
      TabCaption(4)   =   "Lista"
      TabPicture(4)   =   "FrmInmueble.frx":0502
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "dtgInm"
      Tab(4).Control(1)=   "fraInm(5)"
      Tab(4).Control(2)=   "fraInm(4)"
      Tab(4).Control(3)=   "fraInm(7)"
      Tab(4).Control(4)=   "AdoInm"
      Tab(4).ControlCount=   5
      Begin VB.Frame fraInm 
         Height          =   165
         Index           =   2
         Left            =   -74775
         TabIndex        =   104
         Top             =   2130
         Width           =   9150
      End
      Begin VB.Frame FraJunta 
         Caption         =   "Junta de Condominio"
         Height          =   4515
         Left            =   -74775
         TabIndex        =   102
         Top             =   2445
         Width           =   9165
         Begin MSAdodcLib.Adodc adoJC 
            Height          =   330
            Left            =   150
            Top             =   4095
            Visible         =   0   'False
            Width           =   6465
            _ExtentX        =   11404
            _ExtentY        =   582
            ConnectMode     =   0
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
            Caption         =   "Junta de Condominio"
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
         Begin MSDataGridLib.DataGrid dtgJC 
            Bindings        =   "FrmInmueble.frx":051E
            Height          =   3870
            Left            =   195
            TabIndex        =   103
            Top             =   375
            Width           =   8580
            _ExtentX        =   15134
            _ExtentY        =   6826
            _Version        =   393216
            AllowUpdate     =   0   'False
            AllowArrows     =   -1  'True
            Appearance      =   0
            BorderStyle     =   0
            Enabled         =   -1  'True
            HeadLines       =   1,5
            RowHeight       =   15
            TabAcrossSplits =   -1  'True
            TabAction       =   2
            WrapCellPointer =   -1  'True
            RowDividerStyle =   4
            FormatLocked    =   -1  'True
            BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
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
            Caption         =   "J U N T A  D E  C O N D O M I N I O"
            ColumnCount     =   10
            BeginProperty Column00 
               DataField       =   "Codigo"
               Caption         =   "Codigo"
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
               DataField       =   "Nombre"
               Caption         =   "Nombre"
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
               DataField       =   "Cedula"
               Caption         =   "Cedula"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   1
                  Format          =   "#,##0  "
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   8202
                  SubFormatType   =   0
               EndProperty
            EndProperty
            BeginProperty Column03 
               DataField       =   "CarJunta"
               Caption         =   "CarJunta"
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
               Caption         =   "Telefonos"
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
               DataField       =   "ExtOfc"
               Caption         =   "ExtOfc"
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
            BeginProperty Column06 
               DataField       =   "TelfHab"
               Caption         =   "TelfHab"
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
            BeginProperty Column07 
               DataField       =   "Celular"
               Caption         =   "Celular"
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
            BeginProperty Column08 
               DataField       =   "Fax"
               Caption         =   "Fax"
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
            BeginProperty Column09 
               DataField       =   "Email"
               Caption         =   "Email"
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
               ScrollGroup     =   3
               Size            =   3
               BeginProperty Column00 
                  Alignment       =   2
                  DividerStyle    =   6
                  WrapText        =   -1  'True
                  ColumnWidth     =   840,189
               EndProperty
               BeginProperty Column01 
                  ColumnWidth     =   2520
               EndProperty
               BeginProperty Column02 
                  Alignment       =   1
                  ColumnWidth     =   1305,071
               EndProperty
               BeginProperty Column03 
                  Alignment       =   2
                  ColumnWidth     =   1739,906
               EndProperty
               BeginProperty Column04 
                  ColumnWidth     =   1739,906
               EndProperty
               BeginProperty Column05 
                  ColumnWidth     =   1739,906
               EndProperty
               BeginProperty Column06 
                  ColumnWidth     =   1739,906
               EndProperty
               BeginProperty Column07 
                  ColumnWidth     =   1739,906
               EndProperty
               BeginProperty Column08 
                  ColumnWidth     =   1739,906
               EndProperty
               BeginProperty Column09 
                  ColumnWidth     =   1739,906
               EndProperty
            EndProperty
         End
      End
      Begin MSAdodcLib.Adodc AdoInm 
         Height          =   330
         Left            =   -71325
         Top             =   5040
         Visible         =   0   'False
         Width           =   2115
         _ExtentX        =   3731
         _ExtentY        =   582
         ConnectMode     =   0
         CursorLocation  =   3
         IsolationLevel  =   -1
         ConnectionTimeout=   15
         CommandTimeout  =   30
         CursorType      =   2
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
         Caption         =   "AdoInm"
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
      Begin VB.Frame fraInm 
         Caption         =   "Aplciar Filtro:"
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
         Index           =   7
         Left            =   -72825
         TabIndex        =   79
         Top             =   5550
         Width           =   1485
         Begin VB.OptionButton OptBusca 
            Caption         =   "Inactivos"
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
            Index           =   4
            Left            =   300
            TabIndex        =   82
            Tag             =   "Inactivo=True"
            Top             =   975
            Width           =   1080
         End
         Begin VB.OptionButton OptBusca 
            Caption         =   "Ningúno"
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
            Index           =   3
            Left            =   150
            TabIndex        =   81
            Tag             =   "0"
            Top             =   270
            Value           =   -1  'True
            Width           =   1080
         End
         Begin VB.OptionButton OptBusca 
            Caption         =   "Activos"
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
            TabIndex        =   80
            Tag             =   "Inactivo=False"
            Top             =   615
            Width           =   1080
         End
      End
      Begin VB.Frame fraInm 
         Height          =   4515
         Index           =   6
         Left            =   -74820
         TabIndex        =   63
         Top             =   2100
         Width           =   9165
         Begin MSComctlLib.ImageList imglist 
            Left            =   1560
            Top             =   3240
            _ExtentX        =   1005
            _ExtentY        =   1005
            BackColor       =   -2147483643
            ImageWidth      =   16
            ImageHeight     =   16
            MaskColor       =   12632256
            _Version        =   393216
            BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
               NumListImages   =   2
               BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "FrmInmueble.frx":0532
                  Key             =   ""
               EndProperty
               BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "FrmInmueble.frx":0692
                  Key             =   ""
               EndProperty
            EndProperty
         End
         Begin MSComctlLib.Toolbar Toolbar2 
            Height          =   360
            Left            =   5535
            TabIndex        =   78
            Top             =   3960
            Width           =   3405
            _ExtentX        =   6006
            _ExtentY        =   635
            ButtonWidth     =   2302
            ButtonHeight    =   582
            Appearance      =   1
            Style           =   1
            TextAlignment   =   1
            ImageList       =   "imglist"
            _Version        =   393216
            BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
               NumButtons      =   2
               BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Caption         =   "&Agregar  "
                  Key             =   "Add"
                  Object.Tag             =   "Agregar nuevo servicio"
                  ImageIndex      =   1
                  Style           =   5
                  Object.Width           =   1e-4
                  BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
                     NumButtonMenus  =   3
                     BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                        Key             =   "Elec"
                        Text            =   "Electricidad"
                     EndProperty
                     BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                        Key             =   "Agua"
                        Text            =   "Agua"
                     EndProperty
                     BeginProperty ButtonMenu3 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                        Key             =   "Tele"
                        Text            =   "Teléfono"
                     EndProperty
                  EndProperty
               EndProperty
               BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Caption         =   "A&ctualizar  "
                  Key             =   "Update"
                  ImageIndex      =   2
                  Style           =   5
                  BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
                     NumButtonMenus  =   4
                     BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                        Key             =   "All"
                        Text            =   "Todo"
                     EndProperty
                     BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                        Key             =   "Elec"
                        Text            =   "Electricidad"
                     EndProperty
                     BeginProperty ButtonMenu3 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                        Key             =   "Agua"
                        Text            =   "Agua"
                     EndProperty
                     BeginProperty ButtonMenu4 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                        Key             =   "Tele"
                        Text            =   "Teléfono"
                     EndProperty
                  EndProperty
               EndProperty
            EndProperty
            BorderStyle     =   1
            MousePointer    =   99
            MouseIcon       =   "FrmInmueble.frx":0AE6
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid gridServicios 
            Height          =   3000
            Index           =   0
            Left            =   165
            TabIndex        =   75
            Top             =   720
            Width           =   2685
            _ExtentX        =   4736
            _ExtentY        =   5292
            _Version        =   393216
            FixedCols       =   0
            BackColorFixed  =   -2147483646
            ForeColorFixed  =   -2147483639
            BackColorSel    =   8454016
            WordWrap        =   -1  'True
            Appearance      =   0
            FormatString    =   "^IDCS |^Cod.Gasto"
            _NumberOfBands  =   1
            _Band(0).Cols   =   2
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid gridServicios 
            Height          =   3000
            Index           =   1
            Left            =   3180
            TabIndex        =   76
            Top             =   720
            Width           =   2670
            _ExtentX        =   4710
            _ExtentY        =   5292
            _Version        =   393216
            FixedCols       =   0
            BackColorFixed  =   -2147483646
            ForeColorFixed  =   -2147483639
            BackColorSel    =   8454016
            Appearance      =   0
            FormatString    =   "^IDCS |^Cod.Gasto"
            _NumberOfBands  =   1
            _Band(0).Cols   =   2
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid gridServicios 
            Height          =   3000
            Index           =   2
            Left            =   6300
            TabIndex        =   77
            Top             =   705
            Width           =   2670
            _ExtentX        =   4710
            _ExtentY        =   5292
            _Version        =   393216
            FixedCols       =   0
            BackColorFixed  =   -2147483646
            ForeColorFixed  =   -2147483639
            BackColorSel    =   8454016
            Appearance      =   0
            FormatString    =   "^IDCS |^Cod.Gasto"
            _NumberOfBands  =   1
            _Band(0).Cols   =   2
         End
         Begin VB.Label lblInm 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000002&
            Caption         =   "3.- Servicio Telefónico:"
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
            Height          =   330
            Index           =   37
            Left            =   6300
            TabIndex        =   71
            Top             =   390
            Width           =   2665
         End
         Begin VB.Label lblInm 
            Alignment       =   2  'Center
            BackColor       =   &H80000002&
            Caption         =   "2.- Servicio Agua Potable:"
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
            Height          =   330
            Index           =   36
            Left            =   3180
            TabIndex        =   70
            Top             =   405
            Width           =   2670
         End
         Begin VB.Label lblInm 
            Alignment       =   2  'Center
            BackColor       =   &H80000002&
            Caption         =   "1.- Servicio de Electricidad:"
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
            Height          =   330
            Index           =   35
            Left            =   180
            TabIndex        =   69
            Top             =   420
            Width           =   2670
         End
      End
      Begin VB.Frame fraInm 
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
         Left            =   -74790
         TabIndex        =   60
         Top             =   5550
         Width           =   1485
         Begin VB.OptionButton OptBusca 
            Caption         =   "&Nombre"
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
            Left            =   195
            TabIndex        =   62
            Tag             =   "Nombre"
            Top             =   870
            Width           =   1080
         End
         Begin VB.OptionButton OptBusca 
            Caption         =   "&Código"
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
            Left            =   180
            TabIndex        =   61
            Tag             =   "CodInm"
            Top             =   450
            Value           =   -1  'True
            Width           =   1080
         End
      End
      Begin VB.Frame fraInm 
         Caption         =   "Busqueda:"
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
         Left            =   -71160
         TabIndex        =   54
         Top             =   5565
         Width           =   5520
         Begin VB.TextBox txtInm 
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
            Index           =   32
            Left            =   1785
            Locked          =   -1  'True
            TabIndex        =   57
            Top             =   810
            Width           =   690
         End
         Begin VB.TextBox txtInm 
            BackColor       =   &H00FFFFFF&
            Height          =   285
            Index           =   31
            Left            =   1005
            TabIndex        =   56
            Top             =   285
            Width           =   4335
         End
         Begin VB.CommandButton cmdInm 
            Height          =   400
            Index           =   1
            Left            =   4710
            Style           =   1  'Graphical
            TabIndex        =   55
            ToolTipText     =   "Buscar"
            Top             =   700
            Width           =   615
         End
         Begin VB.Label lblInm 
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
            Index           =   27
            Left            =   165
            TabIndex        =   59
            Top             =   315
            Width           =   630
         End
         Begin VB.Label lblInm 
            Caption         =   "Total Registros:"
            Height          =   240
            Index           =   28
            Left            =   180
            TabIndex        =   58
            Top             =   825
            Width           =   1440
         End
      End
      Begin VB.Frame fraInm 
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
         Height          =   4830
         Index           =   3
         Left            =   180
         TabIndex        =   7
         Top             =   2050
         Width           =   9195
         Begin VB.CommandButton cmdInm 
            Caption         =   "Asigna Caja"
            Height          =   330
            Index           =   4
            Left            =   2850
            TabIndex        =   74
            ToolTipText     =   "Buscar"
            Top             =   4140
            Width           =   1275
         End
         Begin VB.TextBox txtInm 
            DataField       =   "Cobrador"
            DataSource      =   "AdoInmueble"
            Height          =   315
            Index           =   45
            Left            =   1275
            TabIndex        =   68
            Top             =   3630
            Width           =   2850
         End
         Begin VB.TextBox txtInm 
            DataField       =   "Postal"
            Height          =   315
            Index           =   3
            Left            =   1275
            TabIndex        =   65
            Top             =   2190
            Width           =   1200
         End
         Begin VB.TextBox txtInm 
            DataField       =   "Notas"
            DataSource      =   "AdoInmueble"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   1170
            Index           =   22
            Left            =   4590
            MaxLength       =   249
            MultiLine       =   -1  'True
            TabIndex        =   36
            Top             =   3300
            Width           =   4350
         End
         Begin VB.TextBox txtInm 
            DataField       =   "Torres"
            DataSource      =   "AdoInmueble"
            Height          =   315
            Index           =   4
            Left            =   6015
            TabIndex        =   1
            Top             =   2190
            Width           =   480
         End
         Begin VB.TextBox txtInm 
            DataField       =   "Pisos"
            DataSource      =   "AdoInmueble"
            Height          =   315
            Index           =   6
            Left            =   8475
            TabIndex        =   2
            Top             =   2190
            Width           =   480
         End
         Begin VB.TextBox txtInm 
            DataField       =   "Direccion"
            ForeColor       =   &H00404040&
            Height          =   1665
            Index           =   2
            Left            =   180
            MaxLength       =   190
            MultiLine       =   -1  'True
            TabIndex        =   0
            Top             =   375
            Width           =   8760
         End
         Begin VB.TextBox txtInm 
            DataField       =   "Unidad"
            DataSource      =   "AdoInmueble"
            Height          =   315
            Index           =   7
            Left            =   8475
            TabIndex        =   4
            Top             =   2640
            Width           =   480
         End
         Begin VB.TextBox txtInm 
            DataField       =   "Locales"
            DataSource      =   "AdoInmueble"
            Height          =   315
            Index           =   5
            Left            =   6015
            TabIndex        =   3
            Top             =   2640
            Width           =   480
         End
         Begin MSDataListLib.DataCombo cmbInm 
            DataField       =   "Ciudad"
            Height          =   315
            Index           =   1
            Left            =   1275
            TabIndex        =   66
            Top             =   2640
            Width           =   2850
            _ExtentX        =   5027
            _ExtentY        =   556
            _Version        =   393216
            BackColor       =   -2147483643
            ListField       =   "Nombre"
            BoundColumn     =   "Nombre"
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo cmbInm 
            DataField       =   "Estado"
            Height          =   315
            Index           =   2
            Left            =   1275
            TabIndex        =   67
            Top             =   3090
            Width           =   2850
            _ExtentX        =   5027
            _ExtentY        =   556
            _Version        =   393216
            BackColor       =   -2147483643
            ListField       =   "Nombre"
            BoundColumn     =   "Nombre"
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo cmbInm 
            DataField       =   "TipoCta"
            Height          =   315
            Index           =   3
            Left            =   1275
            TabIndex        =   73
            Top             =   4140
            Width           =   1470
            _ExtentX        =   2593
            _ExtentY        =   556
            _Version        =   393216
            BackColor       =   -2147483643
            ListField       =   "TipoCta"
            BoundColumn     =   "Nombre"
            Text            =   ""
         End
         Begin VB.Label lblInm 
            Caption         =   "Tipo Cuenta:"
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
            Index           =   42
            Left            =   180
            TabIndex        =   72
            Top             =   4185
            Width           =   1185
         End
         Begin VB.Label lblInm 
            Caption         =   "Cobrador:"
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
            Index           =   44
            Left            =   180
            TabIndex        =   64
            Top             =   3135
            Width           =   840
         End
         Begin VB.Label lblInm 
            Caption         =   "Comentario:"
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
            Index           =   26
            Left            =   4590
            TabIndex        =   37
            Top             =   3075
            Width           =   2190
         End
         Begin VB.Label lblInm 
            Caption         =   "N° de Unidades (Apto./Casas/Ofc.)"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   585
            Index           =   11
            Left            =   6735
            TabIndex        =   19
            Top             =   2595
            Width           =   1815
         End
         Begin VB.Label lblInm 
            Caption         =   "N° de Locales:"
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
            Left            =   4680
            TabIndex        =   18
            Top             =   2685
            Width           =   1215
         End
         Begin VB.Label lblInm 
            Caption         =   "N° de Pisos:"
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
            Left            =   6735
            TabIndex        =   17
            Top             =   2235
            Width           =   1575
         End
         Begin VB.Label lblInm 
            Caption         =   "N° de Torres:"
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
            Left            =   4650
            TabIndex        =   16
            Top             =   2235
            Width           =   1230
         End
         Begin VB.Label lblInm 
            Caption         =   "Estado:"
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
            Left            =   180
            TabIndex        =   15
            Top             =   3660
            Width           =   840
         End
         Begin VB.Label lblInm 
            Caption         =   "Ciudad:"
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
            Left            =   180
            TabIndex        =   14
            Top             =   2715
            Width           =   840
         End
         Begin VB.Label lblInm 
            Caption         =   "Zona Postal:"
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
            Left            =   180
            TabIndex        =   13
            Top             =   2235
            Width           =   1140
         End
         Begin VB.Label lblInm 
            Caption         =   "Direccion:"
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
            Left            =   210
            TabIndex        =   12
            Top             =   150
            Width           =   840
         End
      End
      Begin VB.Frame fraInm 
         Caption         =   "Indique los códigos de los gastos:"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4785
         Index           =   1
         Left            =   -74805
         TabIndex        =   6
         Top             =   2050
         Width           =   9165
         Begin VB.TextBox txtInm 
            Alignment       =   1  'Right Justify
            DataField       =   "CodPagoCondominio"
            BeginProperty DataFormat 
               Type            =   0
               Format          =   "#.##0,00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   0
            EndProperty
            DataSource      =   "AdoInmueble"
            Height          =   315
            Index           =   26
            Left            =   7845
            TabIndex        =   108
            Top             =   1560
            Width           =   1125
         End
         Begin VB.TextBox txtInm 
            Alignment       =   1  'Right Justify
            DataField       =   "CodIngresosVarios"
            BeginProperty DataFormat 
               Type            =   0
               Format          =   "#.##0,00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   0
            EndProperty
            DataSource      =   "AdoInmueble"
            Height          =   315
            Index           =   25
            Left            =   7845
            TabIndex        =   107
            Top             =   1185
            Width           =   1125
         End
         Begin VB.TextBox txtInm 
            Alignment       =   1  'Right Justify
            DataField       =   "CodDedInt"
            BeginProperty DataFormat 
               Type            =   0
               Format          =   "#.##0,00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   0
            EndProperty
            DataSource      =   "AdoInmueble"
            Height          =   315
            Index           =   24
            Left            =   7845
            TabIndex        =   106
            Top             =   810
            Width           =   1125
         End
         Begin VB.TextBox txtInm 
            Alignment       =   1  'Right Justify
            DataField       =   "CodIVA"
            BeginProperty DataFormat 
               Type            =   0
               Format          =   "#.##0,00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   0
            EndProperty
            DataSource      =   "AdoInmueble"
            Height          =   315
            Index           =   23
            Left            =   7845
            TabIndex        =   105
            Top             =   435
            Width           =   1125
         End
         Begin VB.TextBox txtInm 
            Alignment       =   1  'Right Justify
            DataField       =   "CodFondoE"
            BeginProperty DataFormat 
               Type            =   0
               Format          =   """Bs"" #,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   8202
               SubFormatType   =   0
            EndProperty
            DataSource      =   "AdoInmueble"
            Height          =   315
            Index           =   18
            Left            =   4995
            TabIndex        =   100
            Top             =   1950
            Width           =   1125
         End
         Begin VB.TextBox txtInm 
            Alignment       =   1  'Right Justify
            DataField       =   "CodIDB"
            BeginProperty DataFormat 
               Type            =   0
               Format          =   """Bs"" #,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   8202
               SubFormatType   =   0
            EndProperty
            DataSource      =   "AdoInmueble"
            Height          =   315
            Index           =   19
            Left            =   4995
            TabIndex        =   99
            Top             =   2325
            Width           =   1125
         End
         Begin VB.TextBox txtInm 
            Alignment       =   1  'Right Justify
            DataField       =   "CodAbonoCta"
            BeginProperty DataFormat 
               Type            =   0
               Format          =   "#.##0,00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   0
            EndProperty
            DataSource      =   "AdoInmueble"
            Height          =   315
            Index           =   35
            Left            =   4995
            TabIndex        =   98
            Top             =   450
            Width           =   1125
         End
         Begin VB.TextBox txtInm 
            Alignment       =   1  'Right Justify
            DataField       =   "CodAbonoFut"
            BeginProperty DataFormat 
               Type            =   0
               Format          =   "#.##0,00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   0
            EndProperty
            DataSource      =   "AdoInmueble"
            Height          =   315
            Index           =   36
            Left            =   4995
            TabIndex        =   97
            Top             =   825
            Width           =   1125
         End
         Begin VB.TextBox txtInm 
            Alignment       =   1  'Right Justify
            DataField       =   "CodCCheq"
            BeginProperty DataFormat 
               Type            =   0
               Format          =   "#.##0,00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   0
            EndProperty
            DataSource      =   "AdoInmueble"
            Height          =   315
            Index           =   37
            Left            =   4995
            TabIndex        =   96
            Top             =   1200
            Width           =   1125
         End
         Begin VB.TextBox txtInm 
            Alignment       =   1  'Right Justify
            DataField       =   "CodRCheq"
            BeginProperty DataFormat 
               Type            =   0
               Format          =   "#.##0,00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   0
            EndProperty
            DataSource      =   "AdoInmueble"
            Height          =   315
            Index           =   38
            Left            =   4995
            TabIndex        =   95
            Top             =   1575
            Width           =   1125
         End
         Begin VB.CommandButton cmdInm 
            Caption         =   "Cógidos Predeterminados"
            Enabled         =   0   'False
            Height          =   330
            Index           =   2
            Left            =   4710
            TabIndex        =   94
            ToolTipText     =   "Buscar"
            Top             =   2910
            Width           =   2475
         End
         Begin VB.TextBox txtInm 
            Alignment       =   1  'Right Justify
            DataField       =   "CodGastoPenal"
            BeginProperty DataFormat 
               Type            =   0
               Format          =   "#.##0,00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   0
            EndProperty
            DataSource      =   "AdoInmueble"
            Height          =   315
            Index           =   12
            Left            =   1905
            TabIndex        =   93
            Top             =   1935
            Width           =   1125
         End
         Begin VB.TextBox txtInm 
            Alignment       =   1  'Right Justify
            DataField       =   "CodTelegrama"
            BeginProperty DataFormat 
               Type            =   0
               Format          =   "#.##0,00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   0
            EndProperty
            DataSource      =   "AdoInmueble"
            Height          =   315
            Index           =   16
            Left            =   1905
            TabIndex        =   92
            Top             =   3435
            Width           =   1125
         End
         Begin VB.TextBox txtInm 
            Alignment       =   1  'Right Justify
            DataField       =   "CodCarta"
            BeginProperty DataFormat 
               Type            =   0
               Format          =   "#.##0,00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   0
            EndProperty
            DataSource      =   "AdoInmueble"
            Height          =   315
            Index           =   15
            Left            =   1905
            TabIndex        =   91
            Top             =   3060
            Width           =   1125
         End
         Begin VB.TextBox txtInm 
            Alignment       =   1  'Right Justify
            DataField       =   "CodGesChDev"
            BeginProperty DataFormat 
               Type            =   0
               Format          =   "#.##0,00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   0
            EndProperty
            DataSource      =   "AdoInmueble"
            Height          =   315
            Index           =   14
            Left            =   1905
            TabIndex        =   90
            Top             =   2685
            Width           =   1125
         End
         Begin VB.TextBox txtInm 
            Alignment       =   1  'Right Justify
            DataField       =   "CodGastoAdmin"
            BeginProperty DataFormat 
               Type            =   0
               Format          =   "#.##0,00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   0
            EndProperty
            DataSource      =   "AdoInmueble"
            Height          =   315
            Index           =   11
            Left            =   1905
            TabIndex        =   89
            Top             =   1560
            Width           =   1125
         End
         Begin VB.TextBox txtInm 
            Alignment       =   1  'Right Justify
            DataField       =   "CodGestion"
            BeginProperty DataFormat 
               Type            =   0
               Format          =   "#.##0,00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   0
            EndProperty
            DataSource      =   "AdoInmueble"
            Height          =   315
            Index           =   9
            Left            =   1905
            TabIndex        =   88
            Top             =   810
            Width           =   1125
         End
         Begin VB.TextBox txtInm 
            Alignment       =   1  'Right Justify
            DataField       =   "CodIntMora"
            BeginProperty DataFormat 
               Type            =   0
               Format          =   "#.##0,00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   0
            EndProperty
            DataSource      =   "AdoInmueble"
            Height          =   315
            Index           =   10
            Left            =   1905
            TabIndex        =   87
            Top             =   1185
            Width           =   1125
         End
         Begin VB.TextBox txtInm 
            Alignment       =   1  'Right Justify
            DataField       =   "CodChDev"
            BeginProperty DataFormat 
               Type            =   0
               Format          =   "#.##0,00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   0
            EndProperty
            DataSource      =   "AdoInmueble"
            Height          =   315
            Index           =   13
            Left            =   1905
            TabIndex        =   86
            Top             =   2310
            Width           =   1125
         End
         Begin VB.TextBox txtInm 
            Alignment       =   1  'Right Justify
            DataField       =   "CodFondo"
            BeginProperty DataFormat 
               Type            =   0
               Format          =   "#.##0,00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   0
            EndProperty
            DataSource      =   "AdoInmueble"
            Height          =   315
            Index           =   8
            Left            =   1905
            TabIndex        =   85
            Top             =   435
            Width           =   1125
         End
         Begin VB.TextBox txtInm 
            Alignment       =   1  'Right Justify
            DataField       =   "CodHA"
            BeginProperty DataFormat 
               Type            =   0
               Format          =   "#.##0,00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   0
            EndProperty
            DataSource      =   "AdoInmueble"
            Height          =   315
            Index           =   33
            Left            =   1905
            TabIndex        =   84
            Top             =   3810
            Width           =   1125
         End
         Begin VB.TextBox txtInm 
            Alignment       =   1  'Right Justify
            DataField       =   "CodRebHA"
            BeginProperty DataFormat 
               Type            =   0
               Format          =   "#.##0,00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   0
            EndProperty
            DataSource      =   "AdoInmueble"
            Height          =   315
            Index           =   34
            Left            =   1905
            TabIndex        =   83
            Top             =   4185
            Width           =   1125
         End
         Begin VB.TextBox txtInm 
            Alignment       =   1  'Right Justify
            DataField       =   "Deuda"
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "#,##0.00 ""Bs. """
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   8202
               SubFormatType   =   0
            EndProperty
            DataSource      =   "AdoInmueble"
            Height          =   315
            Index           =   17
            Left            =   7200
            Locked          =   -1  'True
            TabIndex        =   8
            Top             =   3435
            Width           =   1650
         End
         Begin VB.TextBox txtInm 
            Alignment       =   1  'Right Justify
            DataField       =   "FondoAct"
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "#,##0.00 ""Bs. """
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   8202
               SubFormatType   =   0
            EndProperty
            DataSource      =   "AdoInmueble"
            Height          =   315
            Index           =   21
            Left            =   7200
            Locked          =   -1  'True
            TabIndex        =   10
            Top             =   4230
            Width           =   1650
         End
         Begin VB.TextBox txtInm 
            Alignment       =   1  'Right Justify
            DataField       =   "FondoIni"
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "#,##0.00 ""Bs. """
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   8202
               SubFormatType   =   0
            EndProperty
            DataSource      =   "AdoInmueble"
            Height          =   315
            Index           =   20
            Left            =   7200
            Locked          =   -1  'True
            TabIndex        =   9
            Top             =   3840
            Width           =   1650
         End
         Begin VB.Label lblInm 
            Caption         =   "I.V.A.:"
            Height          =   210
            Index           =   41
            Left            =   6375
            TabIndex        =   112
            Top             =   480
            Width           =   1470
         End
         Begin VB.Label lblInm 
            Caption         =   "Ded. Intereses:"
            Height          =   210
            Index           =   40
            Left            =   6375
            TabIndex        =   111
            Top             =   855
            Width           =   1470
         End
         Begin VB.Label lblInm 
            Caption         =   "Ing. Varios:"
            Height          =   210
            Index           =   39
            Left            =   6375
            TabIndex        =   110
            Top             =   1230
            Width           =   1470
         End
         Begin VB.Label lblInm 
            Caption         =   "Pago Condominio;"
            Height          =   210
            Index           =   38
            Left            =   6375
            TabIndex        =   109
            Top             =   1605
            Width           =   1470
         End
         Begin VB.Label lblInm 
            Caption         =   "Reposición Cheq.Dev.:"
            Height          =   210
            Index           =   34
            Left            =   3300
            TabIndex        =   52
            Top             =   1620
            Width           =   1755
         End
         Begin VB.Label lblInm 
            Caption         =   "Cambio de Cheques:"
            Height          =   210
            Index           =   33
            Left            =   3300
            TabIndex        =   51
            Top             =   1245
            Width           =   1755
         End
         Begin VB.Label lblInm 
            Caption         =   "Abono Próxima Fact.:"
            Height          =   210
            Index           =   32
            Left            =   3300
            TabIndex        =   50
            Top             =   870
            Width           =   1755
         End
         Begin VB.Label lblInm 
            Caption         =   "Abono a Cuenta:"
            Height          =   210
            Index           =   31
            Left            =   3300
            TabIndex        =   49
            Top             =   495
            Width           =   1755
         End
         Begin VB.Label lblInm 
            Caption         =   "Desc.Hono.Abogado:"
            Height          =   210
            Index           =   30
            Left            =   225
            TabIndex        =   48
            Top             =   4230
            Width           =   1785
         End
         Begin VB.Label lblInm 
            Caption         =   "Honorarios Abogado:"
            Height          =   210
            Index           =   29
            Left            =   225
            TabIndex        =   47
            Top             =   3855
            Width           =   1785
         End
         Begin VB.Label lblInm 
            Caption         =   "Fondo de Reserva:"
            Height          =   210
            Index           =   12
            Left            =   225
            TabIndex        =   46
            Top             =   480
            Width           =   1785
         End
         Begin VB.Label lblInm 
            Caption         =   "Gestión de cobro:"
            Height          =   210
            Index           =   13
            Left            =   225
            TabIndex        =   45
            Top             =   855
            Width           =   1785
         End
         Begin VB.Label lblInm 
            Caption         =   "Intereses de Mora:"
            Height          =   210
            Index           =   14
            Left            =   225
            TabIndex        =   44
            Top             =   1230
            Width           =   1785
         End
         Begin VB.Label lblInm 
            Caption         =   "Gastos de Admin.:"
            Height          =   210
            Index           =   15
            Left            =   225
            TabIndex        =   43
            Top             =   1605
            Width           =   1785
         End
         Begin VB.Label lblInm 
            Caption         =   "Clausula Penal:"
            Height          =   210
            Index           =   16
            Left            =   225
            TabIndex        =   42
            Top             =   1980
            Width           =   1785
         End
         Begin VB.Label lblInm 
            Caption         =   "Cheque Devuelto:"
            Height          =   210
            Index           =   17
            Left            =   225
            TabIndex        =   41
            Top             =   2355
            Width           =   1785
         End
         Begin VB.Label lblInm 
            Caption         =   "Gestión Cheq. Dev.:"
            Height          =   210
            Index           =   18
            Left            =   225
            TabIndex        =   40
            Top             =   2730
            Width           =   1785
         End
         Begin VB.Label lblInm 
            Caption         =   "Emisión Carta de Mora:"
            Height          =   210
            Index           =   19
            Left            =   225
            TabIndex        =   39
            Top             =   3105
            Width           =   1785
         End
         Begin VB.Label lblInm 
            Caption         =   "Telegrama de Mora:"
            Height          =   210
            Index           =   20
            Left            =   225
            TabIndex        =   38
            Top             =   3480
            Width           =   1785
         End
         Begin VB.Label lblInm 
            Caption         =   "Fondo de Reserva Actual:"
            Height          =   210
            Index           =   25
            Left            =   4725
            TabIndex        =   24
            Top             =   4275
            Width           =   1995
         End
         Begin VB.Label lblInm 
            Caption         =   "Fondo de Reserva Inicial:"
            Height          =   210
            Index           =   24
            Left            =   4725
            TabIndex        =   23
            Top             =   3885
            Width           =   1995
         End
         Begin VB.Label lblInm 
            BackStyle       =   0  'Transparent
            Caption         =   "I.D.B"
            Height          =   210
            Index           =   23
            Left            =   3300
            TabIndex        =   22
            Top             =   2370
            Width           =   1755
         End
         Begin VB.Label lblInm 
            BackStyle       =   0  'Transparent
            Caption         =   "Fondo Especial:"
            Height          =   210
            Index           =   22
            Left            =   3285
            TabIndex        =   21
            Top             =   1995
            Width           =   1755
         End
         Begin VB.Label lblInm 
            BackStyle       =   0  'Transparent
            Caption         =   "Deuda a la Fecha:"
            Height          =   210
            Index           =   21
            Left            =   4725
            TabIndex        =   20
            Top             =   3480
            Width           =   1995
         End
      End
      Begin MSDataGridLib.DataGrid dtgInm 
         Bindings        =   "FrmInmueble.frx":0C48
         Height          =   4635
         Left            =   -74760
         TabIndex        =   53
         Top             =   600
         Width           =   9090
         _ExtentX        =   16034
         _ExtentY        =   8176
         _Version        =   393216
         AllowUpdate     =   -1  'True
         BackColor       =   16777215
         BorderStyle     =   0
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
         ColumnCount     =   5
         BeginProperty Column00 
            DataField       =   "CodInm"
            Caption         =   "Código"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   """Bs"" #,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   "Nombre"
            Caption         =   "Nombre"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column02 
            DataField       =   "Unidad"
            Caption         =   "Apartamentos"
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
            DataField       =   "Honorarios"
            Caption         =   "Honorarios"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   """Bs."" #,##0.00 "
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   8202
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column04 
            DataField       =   "Deuda"
            Caption         =   "Deuda Actual"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   """Bs."" #,##0.00 "
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   0
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            AllowRowSizing  =   0   'False
            AllowSizing     =   0   'False
            BeginProperty Column00 
               Alignment       =   2
               ColumnWidth     =   689,953
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   3795,024
            EndProperty
            BeginProperty Column02 
               Alignment       =   2
               ColumnWidth     =   1065,26
            EndProperty
            BeginProperty Column03 
               Alignment       =   1
               ColumnWidth     =   1214,929
            EndProperty
            BeginProperty Column04 
               Alignment       =   1
               ColumnWidth     =   1530,142
            EndProperty
         EndProperty
      End
   End
   Begin ComctlLib.ImageList ImageList1 
      Left            =   10020
      Top             =   1620
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
            Picture         =   "FrmInmueble.frx":0C5D
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmInmueble.frx":0DDF
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmInmueble.frx":0F61
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmInmueble.frx":10E3
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmInmueble.frx":1265
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmInmueble.frx":13E7
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmInmueble.frx":1569
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmInmueble.frx":16EB
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmInmueble.frx":186D
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmInmueble.frx":19EF
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmInmueble.frx":1B71
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmInmueble.frx":1CF3
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "FrmInmueble"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
    Dim rstInmueble(4) As New ADODB.Recordset  'Matriz de ADODB.Recordset utilizados dentro de este módulo
Attribute rstInmueble.VB_VarHelpID = -1
    Private Enum rst                        'Valor elemento dentro de la matriz de recodset's
        Inm = 0
        Ciudad
        Estado
        TipoInm
        tipoCta
    End Enum
    Dim CnTabla As New ADODB.Connection
    Public Caja$
    'rem
    '---------------------------------------------------------------------------------------------
    Private Sub AdoInm_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, _
    ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
    '---------------------------------------------------------------------------------------------
    On Error Resume Next
    rstInmueble(Inm).MoveFirst
    rstInmueble(Inm).Find "CodInm='" & AdoInm.Recordset.Fields!CodInm & "'"
    'Call rtnServicios
    If adReason = adRsnRequery Then txtInm(32) = AdoInm.Recordset.RecordCount
    End Sub


    Private Sub cmbInm_KeyPress(Index%, KeyAscii%): If Index = 3 Then KeyAscii = 0
    End Sub

    '---------------------------------------------------------------------------------------------
    Private Sub cmdInm_Click(Index As Integer)  '
    '---------------------------------------------------------------------------------------------
    'variables locales
    Dim rstCodigo As New ADODB.Recordset
    Dim strCriterio$
    '
    Select Case Index
        Case 1  'Botón Buscar
    '   --------------------
        If txtInm(31) = "" Then
            MsgBox "Debe especificar algún valor en el cuadro 'Buscar'", vbInformation, _
            App.ProductName
        Else
            
            If OptBusca(0) Then 'Buscar por código
                strCriterio = "CodInm='" & txtInm(31) & "'"
            ElseIf OptBusca(1) Then 'Buscar por Nombre
                strCriterio = "Nombre LIKE '*" & txtInm(31) & "*'"
            End If
            AdoInm.Recordset.Find strCriterio
            If AdoInm.Recordset.EOF Then
                AdoInm.Recordset.MoveFirst
                AdoInm.Recordset.Find strCriterio
                If AdoInm.Recordset.EOF Then MsgBox "Criterio de Búsqueda No Encontrado", _
                vbExclamation, App.ProductName
            End If
        End If
        
        Case 2  'Seleccionar Codigos Pre-determinados
    '   --------------------
        
        On Error Resume Next
        rstCodigo.Open "SELECT * FROM Inmueble WHERE CodInm='" & sysCodInm & "';", cnnConexion, _
        adOpenKeyset, adLockOptimistic, adCmdText
        '
        With rstCodigo
            '
            txtInm(8) = !CodFondo   'código del fondo de reserva
            txtInm(9) = !CodGestion 'Gestiónes de cobranzas
            txtInm(10) = !CodIntMora    'Intereses de mora
            txtInm(11) = !CodGastoAdmin 'Gastos administrativos
            txtInm(12) = !CodGastoPenal 'Gasto penal
            txtInm(13) = !CodChdev  'Cheques devueltos
            txtInm(14) = !CodGesChdev   'gestión cheques devueltos
            txtInm(15) = !CodCarta  'cartas de cobranza
            txtInm(16) = !CodTelegrama  'telegramas
            txtInm(18) = !CodFondoE 'fondo especial de emergencia
            txtInm(19) = !CodIDB    'Impuesto al débito bancario
            txtInm(33) = !CodHA 'honorarios de abogado
            txtInm(34) = !CodRebHA  'rebaja honorarios de abogado
            txtInm(35) = !CodAbonoCta   'abono a cuenta
            txtInm(36) = !CodAbonoFut   'abono a futuro
            txtInm(37) = !CodCCheq  'cambio de cheque
            txtInm(38) = !CodRcheq  'reposición cheque devuelto
            '
        End With
        rstCodigo.Close
        Set rstCodigo = Nothing
        
        Case 3  'Actualizar servicios Inmuebles
            Call rtnDomiciliar_Servicios
            
        Case 4  'Registrar Caja
            With FrmAsignaCaja
                .ctlCodInm = txtInm(0)
                .Show vbModal
            End With
        '
    End Select
    '
    End Sub

    Private Sub Form_Resize()
    If WindowState <> vbMinimized Then
        SSTab1.Left = ((ScaleWidth - ScaleLeft) / 2) - SSTab1.Width / 2
        fraInm(0).Left = SSTab1.Left + 240
    End If
    End Sub

    Private Sub Form_Unload(Cancel As Integer)
    Dim I%
    On Error Resume Next
    For I = 0 To 4
        rstInmueble(I).Close
        Set rstInmueble(I) = Nothing
    Next
    CnTabla.Close
    Set CnTabla = Nothing
    Set FrmInmueble = Nothing
    End Sub



    Private Sub gridServicios_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = 46 Then gridServicios(Index).Text = ""
    If Shift = 2 And gridServicios(Index).RowSel >= 1 Then
        If KeyCode = 67 Then
            Clipboard.Clear
            Clipboard.SetText gridServicios(Index).Text
        ElseIf KeyCode = 86 Then
            gridServicios(Index).Text = Clipboard.GetText
        End If
        '
    End If
    End Sub

    Private Sub gridServicios_KeyPress(Index As Integer, KeyAscii As Integer)
    '
    Call Validacion(KeyAscii, "1234567890")
    With gridServicios(Index)
        If KeyAscii = 8 And Len(.Text) > 0 Then .Text = Left(.Text, Len(.Text) - 1)
        If KeyAscii > 26 Then .Text = .Text & Chr(KeyAscii)
    End With
    '
    End Sub

    Private Sub OptBusca_Click(Index As Integer)
    '
    Select Case Index
        Case 0, 1
            AdoInm.Recordset.Sort = OptBusca(Index).Tag
            txtInm(31).SetFocus
            '
        Case 2, 3, 4
            AdoInm.Recordset.Filter = IIf(Index = 3, 0, OptBusca(Index).Tag)
            '
    End Select
    '
    End Sub

    '---------------------------------------------------------------------------------------------
    Private Sub SSTab1_Click(PreviousTab As Integer)    '
    '---------------------------------------------------------------------------------------------
    '
    Select Case SSTab1.tab
        'Datos Generales,Administrativos,Servicios,Adicionales
        Case 0, 1, 2, 3: fraInm(0).Visible = True
        '   ---------------------
        'Lista
        Case 4
            fraInm(0).Visible = Fasle
            txtInm(32) = AdoInm.Recordset.RecordCount
            
    End Select
    '
    End Sub


    '---------------------------------------------------------------------------------------------
    Private Sub Form_Load() '
    '---------------------------------------------------------------------------------------------
    'espero
    Dim ctlTemp As Control
    Dim Registros As Boolean
    AdoInm.ConnectionString = cnnOLEDB + gcPath + "\sac.mdb"
    AdoInm.CursorLocation = adUseClient
    AdoInm.CommandType = adCmdTable
    AdoInm.RecordSource = "Inmueble"
    AdoInm.Refresh
    '
    cmdInm(1).Picture = LoadResPicture("Buscar", vbResIcon)
    CnTabla.CursorLocation = adUseClient
    CnTabla.Open cnnOLEDB + gcPath + "\Tablas.mdb"
    
    rstInmueble(Inm).Open "SELECT * FROM Inmueble ORDER BY CodInm", _
    cnnConexion, adOpenKeyset, adLockOptimistic, adCmdText
    If Not rstInmueble(Inm).EOF Or Not rstInmueble(Inm).BOF Then Registros = True
    If Not gcCodInm = "" Then rstInmueble(Inm).Find "CodInm='" & gcCodInm & "'"
    For Each ctlTemp In Controls
        If TypeName(ctlTemp) = "TextBox" Or TypeName(ctlTemp) = "DataCombo" Then _
        Set ctlTemp.DataSource = rstInmueble(Inm)
    Next
    Set chkADM.DataSource = rstInmueble(Inm)
    '
    Set mskInm(0).DataSource = rstInmueble(Inm)
    Set cnkInm.DataSource = rstInmueble(Inm)
    '
    rstInmueble(TipoInm).Open "SELECT * FROM TipoInm ORDER BY Nombre", _
    CnTabla, adOpenKeyset, adLockOptimistic, adCmdText
    Set cmbInm(0).RowSource = rstInmueble(TipoInm)
    '
    rstInmueble(Ciudad).Open "SELECT * FROM Ciudades ORDER BY Nombre", _
    CnTabla, adOpenKeyset, adLockOptimistic, adCmdText
    Set cmbInm(1).RowSource = rstInmueble(Ciudad)
    '
    rstInmueble(Estado).Open "SELECT * FROM Estados ORDER BY Nombre", _
    CnTabla, adOpenKeyset, adLockOptimistic, adCmdText
    Set cmbInm(2).RowSource = rstInmueble(Estado)
    '
    rstInmueble(tipoCta).Open "SELECT DISTINCT TipoCta FROM Inmueble;", _
    cnnConexion, adOpenStatic, adLockReadOnly, adCmdText
    Set cmbInm(3).RowSource = rstInmueble(tipoCta)
    '
    SSTab1.tab = 0
    For H = 0 To 2
        Set gridServicios(H).FontFixed = LetraTitulo(LoadResString(527), 7, , True)
        gridServicios(H).ColWidth(0) = 1790
        For j = 0 To 1
            gridServicios(H).Row = 0
            gridServicios(H).Col = j
            gridServicios(H).CellAlignment = flexAlignCenterCenter
        Next
    Next
    Set dtgInm.HeadFont = LetraTitulo(LoadResString(527), 7.5, True)
    Set dtgInm.Font = LetraTitulo(LoadResString(528), 8)
    adoJC.CursorLocation = adUseClient
    adoJC.RecordSource = "SELECT Propietarios.* FROM Propietarios INNER JOIN CargoJC ON Propiet" _
    & "arios.CarJunta = CargoJC.Descripcion ORDER BY CargoJC.IDCargo;"
    adoJC.CommandType = adCmdText
    adoJC.LockType = adLockOptimistic
    
    Call rtnServicios
    Call RtnEstado(6, Toolbar1, rstInmueble(Inm).EOF Or rstInmueble(Inm).BOF)
    '
    Set dtgJC.DataSource = adoJC
    '
    End Sub

    '---------------------------------------------------------------------------------------------
    Private Sub Toolbar1_ButtonClick(ByVal Button As Button)    '
    '---------------------------------------------------------------------------------------------
    '
    Dim strDir$, strOrigen$, strDestino$, strAccion$
    ' Utiliza la propiedad Key con la instrucción
    ' SelectCase para especificar una acción.
    '
    With rstInmueble(Inm)
    
        Select Case UCase(Button.Key)
            Case "FIRST"    ' Primer Registro
        '   -----------------
                .MoveFirst
                Call rtnServicios
                
            Case "PREVIOUS"     ' Registro Anterior
        '   -----------------
                .MovePrevious
                If .BOF Then .MoveLast
                Call rtnServicios
                
            Case "NEXT"         ' Registro Siguiente
        '   -----------------
                .MoveNext
                If .EOF Then .MoveFirst
                Call rtnServicios
                
            Case "END"          ' Ultimo Registro
        '   -----------------
                .MoveLast
                Call rtnServicios
                
            Case "NEW"          ' Nuevo FrmInmueble.
        '   -----------------
                For I = 0 To 3: fraInm(I).Enabled = True
                Next
                cmdInm(2).Enabled = True
                .AddNew
                txtInm(0).SetFocus
                txtInm(17).Locked = False
                txtInm(20).Locked = False
                Call RtnEstado(5, Toolbar1, True)
                Call rtnBitacora("Agregar Nuevo Inmueble")
                
            Case "SAVE"         ' Actualizar.
        '   -----------------
                
                MousePointer = vbHourglass
                mskInm(0).PromptInclude = True
                !Ubica = "\" & Trim(txtInm(0)) & "\"
                strDir = gcPath & "\" & Trim(txtInm(0))
                !Usuario = gcUsuario
                !Freg = Date
                If ftnUpdate Then
                '
                    If Respuesta("Desea Guardar la información") Then
                    '
                        If .EditMode = adEditAdd Then
                            !Facturacion = 0
                            MsgBox "Se le Recuerda que debe completar la información del inmueb" _
                            & "le" & vbCrLf & "Parametros de Facturación...", vbInformation, _
                            App.ProductName
                        End If
                        '
                    Else
                        MousePointer = vbDefault
                        Exit Sub
                    End If
                    '
                End If
                '
                strAccion = "Actualizado...."
                '
                If .EditMode = adEditAdd Then
                    mskInm(0) = Date
                    !Caja = Caja
                    If Dir(strDir, vbDirectory) = "" Then   'si no existe el directorio del inm
                        MkDir (strDir)  'crea la carpeta del inmueble
                        MkDir (strDir & "\Reportes")    'crea la carpeta de sus reportes
                        strOrigen = Left(gcPath, 7) & "DESING\inm.mdb"
                        strDestino = gcPath & !Ubica
                        FileCopy "\\servidor\sac\DESING\inm.mdb", strDestino & "inm.mdb"
                    End If
                    strAccion = " Registrado..."
                End If
                '
                Call rtnBitacora("Inm:" & txtInm(0) & "/" & txtInm(1) & strAccion)
                .Update
                txtInm(17).Locked = True
                txtInm(20).Locked = True
                mskInm(0).PromptInclude = False
                'AdoInm.Refresh
                FrmAdmin.objRst.Requery
                Call RtnEstado(6, Toolbar1, rstInmueble(Inm).EOF Or rstInmueble(Inm).BOF)
                For I = 0 To 3: fraInm(I).Enabled = False
                Next
                cmdInm(2).Enabled = False
                MsgBox "Inmueble " & strAccion, vbInformation, App.ProductName
                MousePointer = vbDefault
                
            Case "FIND"         ' Buscar.
        '   -----------------
        
            Case "UNDO"         ' Deshacer Registro
        '   -----------------
                'mskInm(0).PromptInclude = False
                .CancelUpdate
                For I = 0 To 3: fraInm(I).Enabled = False
                Next
                cmdInm(2).Enabled = False
                Call rtnBitacora("Inm:" & txtInm(0) & " cambios cancelados...")
                '.Requery
                Call RtnEstado(6, Toolbar1, .EOF Or .BOF)
                
                
            Case "DELETE"       'Eliminar Registro
        '   -----------------
                If gcNivel > nuAdministrador Then
                    MsgBox "Coño " & gcUsuario & " tú si eres entrepito(a). ¿Quien te dijo pre" _
                    & "sionara este botón?", vbInformation, App.ProductName
                Else
                    MsgBox "Opción no esta disponible,....por ahora....", vbInformation, _
                    App.ProductName
                End If
                
            Case "EDIT1"        'Editar
        '   -----------------
                For I = 0 To 3: fraInm(I).Enabled = True
                Next
                cmdInm(2).Enabled = True
                Call RtnEstado(5, Toolbar1, True)
                
            Case "CLOSE"   'Cerrar
        '   -----------------
                Unload Me
                
            Case "PRINT"        ' Imprimir
        '   -----------------
                MsgBox "Opción no disponible.....por ahora....", vbInformation, App.ProductName
        '
        End Select
    '
    End With
    '
    End Sub


    Private Sub Toolbar2_ButtonClick(ByVal Button As MSComctlLib.Button)
    '
    Select Case Button.Key
        '
        Case "Add"  'Agregar Linea
        '----------
            If Respuesta(LoadResString(524)) Then
                For I = 0 To 2
                    gridServicios(I).AddItem ("")
                    gridServicios(I).Col = 0
                    gridServicios(I).Row = gridServicios(I).Rows - 1
                Next
            End If
            '
            'Actualizar Todo
        Case "Update": If Respuesta(LoadResString(525)) Then Call rtnDomiciliar_Servicios
        '-----------
    End Select
    '
    End Sub

    Private Sub Toolbar2_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
    '
    Dim Indice As Integer   'variables locales
    Dim strCriterio As String
    '
    Select Case ButtonMenu.Parent.Index
    
        Case 1  'Agregar
            Select Case ButtonMenu.Key
                Case "Elec": Indice = 0
                Case "Agua": Indice = 1
                Case "Tele": Indice = 2
            End Select
            gridServicios(Indice).AddItem ("")
            gridServicios(Indice).Row = gridServicios(Indice).Rows - 1
            gridServicios(Indice).Col = 0
        
        Case 2  'Actualizar
            Call rtnDomiciliar_Servicios
            
    End Select
    
    End Sub

    '---------------------------------------------------------------------------------------------
    Private Sub txtInm_KeyPress(Index%, KeyAscii%)  '
    '---------------------------------------------------------------------------------------------
    '
    KeyAscii = Asc(UCase(Chr(KeyAscii)))    'Convierte todo en mayúsculas
    Select Case Index
        Case 0, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 33, 34, 35, 36, 37, 38
    '   ---------------------
            Call Validacion(KeyAscii, "1234567890")
            
        Case 17, 18, 19, 20, 21
    '   ---------------------
            If KeyAscii = 46 Then KeyAscii = 44 'convierte punto en coma
            Call Validacion(KeyAscii, "0123456789,")
        Case 22
    '   ---------------------
            If KeyAscii = 13 Or KeyAscii = 34 Or KeyAscii = 39 Then KeyAscii = 0
        
        Case 27
            Call Validacion(KeyAscii, "0123456789-J")
    End Select
    '
    End Sub
    
    '---------------------------------------------------------------------------------------------
    '
    '   Rutina:     rtnDomicilar_Servisios
    '
    '   Registra en la tabla servicios las cuentas de servicios docimiliadas
    '   para emitir el pago por la administradora
    '---------------------------------------------------------------------------------------------
    Private Sub rtnDomiciliar_Servicios()
    '
    On Error Resume Next
    Dim strSQL$, strInm$
    strInm = txtInm(0)
    cnnConexion.BeginTrans
    cnnConexion.Execute "DELETE FROM Servicios WHERE Inmueble='" & strInm & "'"
    For I = 0 To 2
        With gridServicios(I)
            If .Rows > 1 Then
                For K = 1 To .Rows - 1
                    If Not .TextMatrix(K, 0) = "" Then
                    strSQL = "INSERT INTO SERVICIOS(IDCS,TipoServ,Inmueble,CodGasto) VALUES ('" _
                    & .TextMatrix(K, 0) & "'," & I & ",'" & strInm & "','" & .TextMatrix(K, 1) _
                    & "')"
                    cnnConexion.Execute strSQL
                    End If
                Next
            End If
        End With
    Next
    '
    If Err.Number <> 0 Then
        cnnConexion.RollbackTrans
        MsgBox "Error al intentar actualizar las cuentas de servicios.." & vbCrLf & _
        Err.Description, vbExclamation, App.ProductName
    Else
        cnnConexion.CommitTrans
        MsgBox "Cuentas de Servicios Actualizados", vbInformation, App.ProductName
        rtnBitacora ("Actulizar Servicios Inm.:" & txtInm(0) & "'")
    End If
    '
    End Sub

    '---------------------------------------------------------------------------------------------
    '
    '   Rutina:     rtnServicios
    '
    '   Busca la información de los servicios domiciliados a la administradora
    '   los muestra en pantalla
    '---------------------------------------------------------------------------------------------
    Private Sub rtnServicios()
    'variables locales
    Dim rstServ(2) As New ADODB.Recordset
    Dim l%
    '
    For I = 0 To 2
        rstServ(I).Open "SELECT IDCS,CodGasto FROM Servicios WHERE Inmueble ='" & txtInm(0) & _
        "' AND TipoServ=" & I, cnnConexion, adOpenStatic, adLockReadOnly
        With gridServicios(I)
            If Not rstServ(I).EOF Or Not rstServ(I).BOF Then
                rstServ(I).MoveFirst
                .Rows = rstServ(I).RecordCount + 1
                l = 1
                Do
                    .TextMatrix(l, 0) = IIf(IsNull(rstServ(I)("idcs")), "", rstServ(I)("idcs"))
                    .TextMatrix(l, 1) = IIf(IsNull(rstServ(I)("CodGasto")), "", rstServ(I)("CodGasto"))
                    l = l + 1
                    rstServ(I).MoveNext
                Loop Until rstServ(I).EOF
            
            Else
                gridServicios(I).Rows = 2
                Call rtnLimpiar_Grid(gridServicios(I))
            End If
            If .Rows > 1 Then .Row = 1
            rstServ(I).Close
            Set rstServ(I) = Nothing
        End With
    Next
    'Junta de Condominio
    adoJC.ConnectionString = cnnOLEDB & gcPath & "\" & rstInmueble(Inm).Fields("COdInm") _
    & "\inm.mdb"
    adoJC.Refresh
    adoJC.Recordset.Filter = "CarJunta<>''"
    End Sub


    '---------------------------------------------------------------------------------------------
    '   Funcion: ftnUpdate
    '
    '   Devuelve Verdadera si falta algún dato requerido para guardar
    '   la información mínima necesaria de un inmueble
    '---------------------------------------------------------------------------------------------
    Private Function ftnUpdate() As Boolean
    '
    If txtInm(0) = "" Or Not IsNumeric(txtInm(0)) Then
        ftnUpdate = MsgBox("Valor No Válido en Código del Inmueble..", vbExclamation, _
        App.ProductName)
    End If
    If txtInm(1) = "" Then
        ftnUpdate = MsgBox("Falta Nombre del Inmueble..", vbExclamation, App.ProductName)
    End If
    If cmbInm(0) = "" Then
        ftnUpdate = MsgBox("Defina Tipo de Condominio...", vbExclamation, App.ProductName)
    End If
    If cmbInm(3) = "" Then
        ftnUpdate = MsgBox("Falta Modalidad Cuenta: (POTE/PARTICULAR)", vbExclamation, _
        App.ProductName)
    End If
    For I = 8 To 16
        If txtInm(I) = "" Then
            ftnUpdate = MsgBox("Falta algún valor en los datos administrativos.", vbExclamation, _
            App.ProductName)
            Exit For
        End If
    Next
    For I = 33 To 38
        If txtInm(I) = "" Then
            ftnUpdate = MsgBox("Falta algún valor en los datos administrativos.", vbExclamation, _
            App.ProductName)
            Exit For
        End If
    Next
'    If Caja = "" Then
'        ftnUpdate = MsgBox("Falta completar la información de la caja del inmueble", _
'        vbInformation, App.ProductName)
'    End If
    '
    End Function


    
