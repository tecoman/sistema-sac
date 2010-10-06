VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmSSO 
   Caption         =   "Remesa S.S.O"
   ClientHeight    =   7380
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   11880
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7380
   ScaleWidth      =   11880
   WindowState     =   2  'Maximized
   Begin ComctlLib.Toolbar BarH 
      Align           =   1  'Align Top
      Height          =   480
      Left            =   0
      TabIndex        =   18
      Top             =   0
      Width           =   11880
      _ExtentX        =   20955
      _ExtentY        =   847
      ButtonWidth     =   714
      ButtonHeight    =   688
      AllowCustomize  =   0   'False
      Appearance      =   1
      ImageList       =   "ImageList1"
      _Version        =   327682
      BeginProperty Buttons {0713E452-850A-101B-AFC0-4210102A8DA7} 
         NumButtons      =   15
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
            Object.ToolTipText     =   "Buscar en la lista"
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
         BeginProperty Button13 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button14 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Copiar"
            Object.ToolTipText     =   "Copiar "
            Object.Tag             =   ""
            ImageIndex      =   13
         EndProperty
         BeginProperty Button15 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Pegar"
            Object.ToolTipText     =   "Pegar Copia"
            Object.Tag             =   ""
            ImageIndex      =   14
         EndProperty
      EndProperty
      BorderStyle     =   1
      MousePointer    =   99
      MouseIcon       =   "frmSSO.frx":0000
   End
   Begin TabDlg.SSTab Fichas 
      Height          =   6555
      Left            =   180
      TabIndex        =   0
      Top             =   645
      Width           =   11550
      _ExtentX        =   20373
      _ExtentY        =   11562
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   617
      ShowFocusRect   =   0   'False
      TabCaption(0)   =   "Datos Generales"
      TabPicture(0)   =   "frmSSO.frx":031A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Fra(0)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Lista"
      TabPicture(1)   =   "frmSSO.frx":0336
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Fra(1)"
      Tab(1).ControlCount=   1
      Begin VB.Frame Fra 
         Height          =   5850
         Index           =   1
         Left            =   -74790
         TabIndex        =   26
         Top             =   450
         Width           =   11100
         Begin MSDataGridLib.DataGrid DataGrid1 
            Bindings        =   "frmSSO.frx":0490
            Height          =   4245
            Left            =   135
            TabIndex        =   27
            Top             =   210
            Width           =   10815
            _ExtentX        =   19076
            _ExtentY        =   7488
            _Version        =   393216
            AllowUpdate     =   0   'False
            DefColWidth     =   1
            HeadLines       =   1
            RowHeight       =   15
            FormatLocked    =   -1  'True
            BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
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
            Caption         =   "F A C T U R A S    R E M E S A D A S   S.O.S"
            ColumnCount     =   8
            BeginProperty Column00 
               DataField       =   "RemSSO"
               Caption         =   "Nº Rem"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   1
                  Format          =   "0000"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   8202
                  SubFormatType   =   0
               EndProperty
            EndProperty
            BeginProperty Column01 
               DataField       =   "Ndoc"
               Caption         =   "Nº Doc."
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   1
                  Format          =   "00000"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   8202
                  SubFormatType   =   0
               EndProperty
            EndProperty
            BeginProperty Column02 
               DataField       =   "Fact"
               Caption         =   "Nº Fact."
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
               DataField       =   "CodInm"
               Caption         =   "Inm"
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
               DataField       =   "CodProv"
               Caption         =   "Cod.Prov."
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   1
                  Format          =   "0000"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   8202
                  SubFormatType   =   0
               EndProperty
            EndProperty
            BeginProperty Column05 
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
            BeginProperty Column06 
               DataField       =   "Total"
               Caption         =   "Monto"
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
            BeginProperty Column07 
               DataField       =   "FRecep"
               Caption         =   "Fecha"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   1
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
                  ColumnWidth     =   810,142
               EndProperty
               BeginProperty Column01 
                  Alignment       =   2
                  ColumnWidth     =   975,118
               EndProperty
               BeginProperty Column02 
                  Alignment       =   2
                  ColumnWidth     =   900,284
               EndProperty
               BeginProperty Column03 
                  Alignment       =   2
                  ColumnWidth     =   750,047
               EndProperty
               BeginProperty Column04 
                  Alignment       =   2
                  ColumnWidth     =   915,024
               EndProperty
               BeginProperty Column05 
                  ColumnWidth     =   3240
               EndProperty
               BeginProperty Column06 
                  Alignment       =   1
                  ColumnWidth     =   1349,858
               EndProperty
               BeginProperty Column07 
                  Alignment       =   2
                  ColumnWidth     =   1184,882
               EndProperty
            EndProperty
         End
         Begin MSAdodcLib.Adodc Ado 
            Height          =   330
            Left            =   135
            Top             =   210
            Visible         =   0   'False
            Width           =   3705
            _ExtentX        =   6535
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
            Caption         =   "Facturas Remesadas"
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
      Begin VB.Frame Fra 
         Enabled         =   0   'False
         Height          =   5850
         Index           =   0
         Left            =   210
         TabIndex        =   19
         Top             =   450
         Width           =   11100
         Begin VB.TextBox Txt 
            Alignment       =   1  'Right Justify
            DataField       =   "Titulo"
            DataSource      =   "AdoRemesa"
            Height          =   315
            Index           =   6
            Left            =   7710
            Locked          =   -1  'True
            TabIndex        =   29
            Text            =   "0,00"
            Top             =   315
            Width           =   1425
         End
         Begin VB.TextBox Txt 
            BackColor       =   &H00FFFFFF&
            DataField       =   "Codigo"
            DataSource      =   "AdoRemesa"
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
            Left            =   1560
            Locked          =   -1  'True
            MaxLength       =   5
            TabIndex        =   2
            Top             =   315
            Width           =   1335
         End
         Begin VB.TextBox Txt 
            DataField       =   "Titulo"
            DataSource      =   "AdoRemesa"
            Height          =   315
            Index           =   2
            Left            =   1785
            TabIndex        =   9
            Top             =   2115
            Width           =   1035
         End
         Begin VB.TextBox Txt 
            DataField       =   "Titulo"
            DataSource      =   "AdoRemesa"
            Height          =   315
            Index           =   3
            Left            =   2940
            TabIndex        =   11
            Top             =   2115
            Width           =   3120
         End
         Begin VB.TextBox Txt 
            DataField       =   "Titulo"
            DataSource      =   "AdoRemesa"
            Height          =   315
            Index           =   4
            Left            =   6180
            MaxLength       =   6
            TabIndex        =   13
            Top             =   2115
            Width           =   1245
         End
         Begin VB.TextBox Txt 
            Alignment       =   1  'Right Justify
            DataField       =   "Titulo"
            DataSource      =   "AdoRemesa"
            Height          =   315
            Index           =   5
            Left            =   8700
            TabIndex        =   17
            Text            =   "0,00"
            Top             =   2115
            Width           =   1575
         End
         Begin VB.TextBox Txt 
            DataField       =   "Titulo"
            DataSource      =   "AdoRemesa"
            Height          =   315
            Index           =   1
            Left            =   360
            TabIndex        =   7
            Top             =   2115
            Width           =   1305
         End
         Begin VB.CommandButton Cmd 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Index           =   0
            Left            =   10350
            Picture         =   "frmSSO.frx":04A2
            Style           =   1  'Graphical
            TabIndex        =   20
            ToolTipText     =   "Agregar linea"
            Top             =   1815
            Width           =   525
         End
         Begin VB.CommandButton Cmd 
            Caption         =   "Editar Parámetros"
            Height          =   390
            Index           =   1
            Left            =   7710
            TabIndex        =   5
            ToolTipText     =   "Agregar linea"
            Top             =   1155
            Width           =   1440
         End
         Begin MSMask.MaskEdBox Msk 
            Height          =   315
            Index           =   0
            Left            =   4650
            TabIndex        =   4
            Top             =   315
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   556
            _Version        =   393216
            MaxLength       =   7
            Format          =   "mm-yyyy"
            Mask            =   "##-####"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox Msk 
            Height          =   315
            Index           =   1
            Left            =   7545
            TabIndex        =   15
            Top             =   2115
            Width           =   1035
            _ExtentX        =   1826
            _ExtentY        =   556
            _Version        =   393216
            MaxLength       =   7
            Format          =   "mm-yyyy"
            Mask            =   "##-####"
            PromptChar      =   "_"
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid Grid 
            Height          =   2835
            Left            =   360
            TabIndex        =   21
            Tag             =   "800|3500|1200|1200|1000|1200|350|650"
            Top             =   2670
            Width           =   10515
            _ExtentX        =   18547
            _ExtentY        =   5001
            _Version        =   393216
            Cols            =   8
            FixedCols       =   0
            BackColorFixed  =   -2147483646
            ForeColorFixed  =   -2147483639
            BackColorBkg    =   -2147483636
            Enabled         =   0   'False
            FormatString    =   "Cod.Inm|Nom. Inmueble|Nº Empresa|Nº Factura|Período|>Monto|No|Eliminar"
            _NumberOfBands  =   1
            _Band(0).Cols   =   8
         End
         Begin VB.Label Lbl 
            Caption         =   "Total Remesa:"
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
            Left            =   6375
            TabIndex        =   28
            Top             =   345
            Width           =   1320
         End
         Begin VB.Label Lbl 
            Caption         =   "Nº &Remesa:"
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
            Left            =   285
            TabIndex        =   1
            Top             =   367
            Width           =   1035
         End
         Begin VB.Label Lbl 
            Caption         =   "&Cargado:"
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
            Left            =   3750
            TabIndex        =   3
            Top             =   360
            Width           =   1035
         End
         Begin VB.Label Lbl 
            Caption         =   "Inf. del Gasto:"
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
            Left            =   285
            TabIndex        =   25
            Top             =   795
            Width           =   1275
         End
         Begin VB.Label Lbl 
            Caption         =   "Proveedor:"
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
            Left            =   285
            TabIndex        =   24
            Top             =   1230
            Width           =   1245
         End
         Begin VB.Label Lbl 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            Caption         =   " Gasto:"
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
            Left            =   1560
            TabIndex        =   23
            Top             =   750
            Width           =   5970
         End
         Begin VB.Label Lbl 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            Caption         =   " Proveedor:"
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
            Index           =   5
            Left            =   1560
            TabIndex        =   22
            Top             =   1185
            Width           =   5970
         End
         Begin VB.Label Lbl 
            Alignment       =   2  'Center
            BackColor       =   &H80000002&
            Caption         =   "Cod. &Inm"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000009&
            Height          =   285
            Index           =   7
            Left            =   1755
            TabIndex        =   8
            Top             =   1815
            Width           =   1110
         End
         Begin VB.Label Lbl 
            Alignment       =   2  'Center
            BackColor       =   &H80000002&
            Caption         =   "&Nom. Inm"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000009&
            Height          =   285
            Index           =   8
            Left            =   2895
            TabIndex        =   10
            Top             =   1815
            Width           =   3210
         End
         Begin VB.Label Lbl 
            Alignment       =   2  'Center
            BackColor       =   &H80000002&
            Caption         =   "Nº &Factura"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000009&
            Height          =   285
            Index           =   9
            Left            =   6135
            TabIndex        =   12
            Top             =   1815
            Width           =   1350
         End
         Begin VB.Label Lbl 
            Alignment       =   2  'Center
            BackColor       =   &H80000002&
            Caption         =   "&Período"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000009&
            Height          =   285
            Index           =   10
            Left            =   7515
            TabIndex        =   14
            Top             =   1815
            Width           =   1125
         End
         Begin VB.Label Lbl 
            Alignment       =   2  'Center
            BackColor       =   &H80000002&
            Caption         =   "&Monto"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000009&
            Height          =   285
            Index           =   11
            Left            =   8670
            TabIndex        =   16
            Top             =   1815
            Width           =   1605
         End
         Begin VB.Label Lbl 
            Alignment       =   2  'Center
            BackColor       =   &H80000002&
            Caption         =   "Nº &Empresa"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000009&
            Height          =   285
            Index           =   6
            Left            =   360
            TabIndex        =   6
            Top             =   1815
            Width           =   1365
         End
         Begin VB.Image img 
            Enabled         =   0   'False
            Height          =   480
            Index           =   1
            Left            =   0
            Top             =   0
            Visible         =   0   'False
            Width           =   480
         End
         Begin VB.Image img 
            Enabled         =   0   'False
            Height          =   480
            Index           =   0
            Left            =   100
            Top             =   100
            Visible         =   0   'False
            Width           =   480
         End
      End
   End
   Begin ComctlLib.ImageList ImageList1 
      Left            =   11340
      Top             =   675
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   20
      ImageHeight     =   20
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   14
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmSSO.frx":05EC
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmSSO.frx":076E
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmSSO.frx":08F0
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmSSO.frx":0A72
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmSSO.frx":0BF4
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmSSO.frx":0D76
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmSSO.frx":0EF8
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmSSO.frx":107A
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmSSO.frx":11FC
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmSSO.frx":137E
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmSSO.frx":1500
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmSSO.frx":1682
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmSSO.frx":1804
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmSSO.frx":1B1E
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmSSO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'variables locales a nivel de formulario
Dim gProveedor(1) As String
Dim gGasto(1) As String


Private Sub BarH_ButtonClick(ByVal Button As ComctlLib.Button)
'
Dim Copia As rSSO
Dim N As Long, I As Long
Dim rstLocal As ADODB.Recordset

With Ado.Recordset

    Select Case UCase(Button.Key)
        Case "FIRST"
            .MoveFirst
        
        Case "NEXT"
            .MoveNext
            If .EOF Then .MoveFirst
        
        Case "PREVIOUS"
            .MovePrevious
            If .BOF Then .MoveLast
        
        Case "END"
            .MoveLast
            
        Case "SAVE"
            Call Guardar_Remesa
            Call RtnEstado(Button.Index, BarH, Ado.Recordset.EOF Or Ado.Recordset.BOF)
            BarH.Buttons("Delete").Enabled = False
            For I = 1 To 4: BarH.Buttons(I).Enabled = False
            Next
            
        Case "CLOSE"
            Unload Me
            Set frmSSO = Nothing
            
        Case "DELETE"
            Call EliminarITem
            
        Case "NEW"
            Call RtnEstado(Button.Index, BarH, False)
            Call rtnLimpiar_Grid(Grid)
            BarH.Buttons("Delete").Enabled = True
            Fichas.Tab = 0
            Fichas.TabEnabled(1) = False
            Fra(0).Enabled = True
            Grid.Enabled = False
            Txt(6) = "0,00"
            Txt(0) = NRemSSO
            Msk(0).SetFocus
            
        Case "UNDO"
            Call CancelarEdicion
            Call RtnEstado(Button.Index, BarH, True)
            BarH.Buttons("Delete").Enabled = False
            Txt(6) = "0,00"
            Fra(0).Enabled = False
            Fichas.TabEnabled(1) = True
            
        Case "FIND"
            Fichas.Tab = 1
                
        Case "COPIAR"
        'escribe un archivo temporal de longitud fija
        N = FreeFile
        If Dir(gcPath & "\rSSO.prn", vbArchive) <> "" Then Kill gcPath & "\rSSO.prn"
        
        Open gcPath & "\rSSO.prn" For Random As #N Len = Len(Copia)
            For I = 1 To Grid.Rows - 1
                Copia.CodInm = Grid.TextMatrix(I, 0)
                Copia.Nfact = Grid.TextMatrix(I, 3)
                Copia.Monto = Grid.TextMatrix(I, 5) * 100
                Put #N, , Copia
            Next
        Close #N
            
        Case "PEGAR"
        
        N = FreeFile
        Open gcPath & "\rSSO.prn" For Random As #N Len = Len(Copia)
        Set rstLocal = New ADODB.Recordset
        rstLocal.Open "Inmueble", cnnConexion, adOpenKeyset, adLockOptimistic, adCmdTable
            Call rtnLimpiar_Grid(Grid)
            Do
            '
                I = I + 1
                Get #N, , Copia
                If Asc(Copia.CodInm) > 0 Then
                    rstLocal.Filter = "CodInm ='" & Trim(Copia.CodInm) & "'"
                    If Not rstLocal.EOF And Not .BOF Then
                        If I > 1 Then Grid.AddItem ("")
                        Grid.TextMatrix(I, 0) = rstLocal("CodInm")
                        Grid.TextMatrix(I, 1) = rstLocal("Nombre")
                        Grid.TextMatrix(I, 2) = rstLocal("Campo3")
                        Grid.TextMatrix(I, 3) = Trim(Copia.Nfact)
                        Grid.TextMatrix(I, 4) = ""
                        Grid.TextMatrix(I, 5) = Format((Trim(Copia.Monto) / 100), "#,##0.00")
                    End If
                End If
            '
            Loop Until EOF(N)
            'ahora llena el flex con la información del temporal
            '
            
            '
        Close #N
        '
    End Select
    
End With
'
End Sub

Private Sub cmd_Click(Index As Integer)
'Variables locales
Select Case Index
    'modificar los parámetros del sistema
    Case 1: Call FrmAdmin.AC6016_Click(2)
    'agregar línea
    Case 0: Call AgregarLinea
End Select
'
End Sub

Private Sub Fichas_Click(PreviousTab As Integer)
If Fichas.Tab = 0 Then
    For I = 1 To 4: BarH.Buttons(I).Enabled = False
    Next
    BarH.Buttons(14).Enabled = True
    BarH.Buttons(15).Enabled = Dir(gcPath & "\rSSO.prn", vbArchive) <> ""
Else
    For I = 1 To 4: BarH.Buttons(I).Enabled = True
    Next
    BarH.Buttons(14).Enabled = False
    BarH.Buttons(15).Enabled = False
End If
End Sub

Private Sub Form_Load()
'variables locales
Dim rstLocal As New ADODB.Recordset
Dim Cod0$, Cod2$
Dim Respuesta&
'
img(0).Picture = LoadResPicture("Unchecked", vbResBitmap)
img(1).Picture = LoadResPicture("Checked", vbResBitmap)
Call centra_titulo(Grid, True)
Lbl(5) = ""
Lbl(4) = ""
'
With rstLocal

    .Open "ParamSSO", cnnConexion, adOpenKeyset, adLockOptimistic, adCmdTable
    If Not .EOF And Not .BOF Then
        Cod0 = !CodPro
        Cod2 = !codGasto
        .Close
        .Open "Proveedores", cnnConexion, adOpenKeyset, adLockOptimistic, adCmdTable
        If Not .EOF And Not .BOF Then
            .Find "Codigo='" & Cod0 & "'"
            If Not .EOF And Not .BOF Then
                Lbl(5) = Space(1) & !Codigo & " - " & !NombProv
                gProveedor(0) = !Codigo
                gProveedor(1) = !NombProv
            End If
        End If
        .Close
        FrmAdmin.objRst.MoveFirst
        FrmAdmin.objRst.Find "Inactivo = False"
        '
        If Not FrmAdmin.objRst.EOF Then
            .Open "TGastos", cnnOLEDB & gcPath & FrmAdmin.objRst!Ubica & "inm.mdb", _
            adOpenKeyset, adLockOptimistic, adCmdTable
            
            If Not .EOF And Not .BOF Then
            
                .Filter = "CodGasto='" & Cod2 & "'"
                If Not .EOF And Not .BOF Then
                    Lbl(4) = Space(1) & .Fields("CodGasto") & " - " & !Titulo
                    gGasto(0) = !codGasto
                    gGasto(1) = !Titulo
                End If
                
            End If
            
        End If
        
    Else
        Respuesta = MsgBox("Faltan parámetros necesarios para procesar la remesa del " _
        & "S.S.O." & vbCrLf & vbCrLf & "¿Desea completar esta información ahora?", vbQuestion _
        + vbYesNo, "Falta Información de los parámetros")
        If Respuesta = vbYes Then Call FrmAdmin.AC6016_Click(2)
    End If
    .Close
End With
Set rstLocal = Nothing
Call Configurar_ADO(Ado, adCmdText, "SELECT * FROM Cpp WHERE Not IsNull (RemSSO) AND RemSSO>0", _
gcPath & "\sac.mdb")
Call RtnEstado(6, BarH, Ado.Recordset.EOF Or Ado.Recordset.BOF)
BarH.Buttons("Delete").Enabled = False
BarH.Buttons("Edit1").Visible = False
'
End Sub

Private Sub grid_Click()
Dim C As Long   'Columna

With Grid
    C = .ColSel
    If C = .Cols - 1 Then
        .Col = C
        .Row = .RowSel
        If .CellPicture = img(0) Then
            Set .CellPicture = img(1)
        Else
            Set .CellPicture = img(0)
        End If
        '.CellPictureAlignment=flexalig
    End If
End With
End Sub

Private Sub Msk_GotFocus(Index As Integer): Msk(Index).BackColor = &HFFC0C0
End Sub

Private Sub msk_KeyPress(Index As Integer, KeyAscii As Integer)
Call Validacion(KeyAscii, "0123456789")
If KeyAscii = 13 Then SendKeys vbTab
End Sub


Private Sub Msk_LostFocus(Index As Integer): Msk(Index).BackColor = &HFFFFFF
End Sub


Private Sub txt_GotFocus(Index As Integer)
    Txt(Index).BackColor = &HFFC0C0
    If Index = 5 And Txt(5) <> "" Then
        Txt(5) = CCur(Txt(5))
    ElseIf Index = 6 Then
        SendKeys vbTab
    End If
    Txt(Index).SelStart = 0
    Txt(Index).SelLength = Len(Txt(Index))
    
End Sub

Private Sub txt_KeyPress(Index As Integer, KeyAscii As Integer)
'convierte todo en mayúscula
KeyAscii = Asc(UCase(Chr(KeyAscii)))
Select Case Index
    Case 0, 2, 4 'n remesa, Cod.Inm,Nª Factura
        Call Validacion(KeyAscii, "0123456789")
        If Index = 2 And KeyAscii = 13 Then
            If Txt(Index) = "" Then
                SendKeys vbTab
            Else
                Call BuscaInf("", Txt(Index), "")
            End If
        ElseIf KeyAscii = 13 Then
            SendKeys vbTab
        End If
        
    Case 1  'nº empresa
        Call Validacion(KeyAscii, "D1234567890")
        If InStr(Txt(1), "D") > 0 And KeyAscii = Asc("D") Then KeyAscii = 0
        If KeyAscii = 13 Then
            If Txt(1) = "" Then
                SendKeys vbTab
            Else
                Call BuscaInf(Txt(1))
            End If
        End If
        
    Case 3  'nombre inm
        Call Validacion(KeyAscii, "0123456789ABCDEFGHIJKLMNÑOPQRSTUVWXYZ-")
        If KeyAscii = 13 Then
            If Txt(3) = "" Then
                SendKeys vbTab
            Else
                Call BuscaInf("", "", Txt(3))
            End If
        End If
        
    Case 5  'Monto
        If KeyAscii = Asc(".") Then KeyAscii = Asc(",")
        Call Validacion(KeyAscii, "0123456789,")
        If KeyAscii = 13 Then SendKeys vbTab
        
End Select

End Sub

'-------------------------------------------------------------------------------------------------
'   Rutina: BuscaInf
'
'   Entradas:   Opcional NEmpresa(Nª de Empresa en SSO)
'               Opcional CodInm (Código del Inmueble)
'               Opcional NomInm (Nombre [parcial o total] del inmueble
'
'   Busca la coincidencia completa los datos del S.S.O
'-------------------------------------------------------------------------------------------------
Private Sub BuscaInf(Optional NEmpresa As String, Optional CodInm As String, _
Optional NomInm As String)
'variables locales
Dim Criterio As String

If NEmpresa <> "" Or CodInm <> "" Or NomInm <> "" Then
    If NEmpresa <> "" Then
        If Left(NEmpresa, 1) <> "D" Then NEmpresa = "D" & NEmpresa
        Criterio = "Campo3 ='" & NEmpresa & "'"
        Txt(2) = ""
        Txt(3) = ""
    ElseIf CodInm <> "" Then
        Criterio = "CodInm='" & CodInm & "'"
        Txt(1) = ""
        Txt(3) = ""
    Else
        Criterio = "Nombre Like '%" & NomInm & "%'"
        Txt(1) = ""
        'Txt(2) = ""
    End If
    '
    With FrmAdmin.objRst
    
        If Not .EOF And Not .BOF Then
        '
            .MoveFirst
            .Find Criterio
            '
            If Not .EOF Then
                Txt(1) = IIf(IsNull(!Campo3), "", !Campo3)
                Txt(2) = !CodInm
                Txt(3) = !Nombre
                Txt(4).SetFocus
            Else
                .MoveFirst
                MsgBox "No existe coincidencia", vbInformation, App.ProductName
            End If
            '
        End If
        '
    End With
    '
End If
'
End Sub

'-------------------------------------------------------------------------------------------------
'
'   Rutina:     AgregarLinea
'
'   Esta rutina adiciona la linea al grid
'-------------------------------------------------------------------------------------------------
Private Sub AgregarLinea()
'variables locales
Dim Fila&
Dim Cadena$
'valida los datos necesarios para agregar a la lista

If Txt(1) = "" Then Cadena = "- Falta Nº de la empresa." & vbCrLf
If Txt(2) = "" Then Cadena = Cadena & "- Falta Código del Inmueble." & vbCrLf
If Txt(3) = "" Then Cadena = Cadena & "- Falta nombre del inmueble." & vbCrLf
If Txt(4) = "" Then Cadena = Cadena & "- Falta el Nº de la factura." & vbCrLf
If Not IsDate("01/" & Msk(1)) Then Cadena = Cadena & "- Período inválido." & vbCrLf
If Txt(5) = "" Then Cadena = Cadena & "- Falta monto de la factura." & vbCrLf
If Txt(5) <> "" Then
    If CCur(Txt(5)) = 0 Then Cadena = Cadena & "- Monto no puede ser igual a cero." & vbCrLf
End If
If Cadena <> "" Then
    MsgBox "No se puede agregar esta línea:" & vbCrLf & vbCrLf & Cadena, vbCritical, _
    App.ProductName
    Txt(1).SetFocus
    Exit Sub
End If
'
With Grid
    Fila = .Rows - 1
    If .TextMatrix(Fila, 0) <> "" And .TextMatrix(Fila, 1) <> "" And .TextMatrix(Fila, 2) <> "" Then
        .Redraw = True
        Fila = Fila + 1
        .AddItem "", Fila
    
    End If
    If .Rows - 1 = 1 Then Grid.Enabled = True
    .TextMatrix(Fila, 0) = Txt(2)   'Codigo del inmueble
    .TextMatrix(Fila, 1) = Txt(3)   'Nombre del Inmueble
    .TextMatrix(Fila, 2) = Txt(1)   'Nª de empresa en el SSO
    .TextMatrix(Fila, 3) = Txt(4)   'Nª de Factura SSO
    .TextMatrix(Fila, 4) = Msk(1)   'Período al cual corresponde la factura
    .TextMatrix(Fila, 5) = Txt(5)   'Monto
    Txt(6) = Format(CCur(Txt(6)) + CCur(Txt(5)), "#,##0.00")
    .Row = Fila
    .Col = 6
    Set .CellPicture = img(0)
    .CellPictureAlignment = flexAlignCenterCenter
    .Col = 7
    Set .CellPicture = img(0)
    .CellPictureAlignment = flexAlignCenterCenter
    
    'LIMPIA EL CONTENIDO DE LA LINEA
    Msk(1).PromptInclude = False
    Txt(2) = "": Txt(3) = "": Txt(1) = "": Txt(4) = "": Msk(1) = "": Txt(5) = "0,00"
    Msk(1).PromptInclude = True
    Txt(1).SetFocus
    
End With
'
End Sub

Private Sub txt_LostFocus(Index As Integer)
    Txt(Index).BackColor = &HFFFFFF
    If Index = 5 And Txt(5) <> "" Then Txt(5) = Format(Txt(5), "#,##0.00")
End Sub

'-------------------------------------------------------------------------------------------------
'   Rutina:     Eliminar Items
'
'   Elimina de la lista todas las faturas marcas en la columna
'   'eliminar'
'-------------------------------------------------------------------------------------------------
Private Sub EliminarITem()
'variables locales
Dim I&
With Grid
    '
    I = 1
    Do
        .Col = .Cols - 1
        .Row = I
        If .CellPicture = img(1) Then
           Set .CellPicture = Nothing
           .Col = .Cols - 2
           Set .CellPicture = Nothing
           .RowHeight(I) = 0
           Txt(6) = Format(CCur(Txt(6)) - CCur(.TextMatrix(I, 5)), "#,##0.00")
        End If
        I = I + 1
    Loop Until I > (.Rows - 1)
    
End With
'
End Sub

'-------------------------------------------------------------------------------------------------
'
'   Rutina:     Guardar_Remesa
'
'   Efecuta el procesamiento de la remesa, registra la remesa,
'   agrega las facturas a Cuentas por Pagar, Asigna los gastos
'-------------------------------------------------------------------------------------------------
Private Sub Guardar_Remesa()
'variables locales
Dim I&, strSql$, NDoc$
If novalida Then Exit Sub
'
On Error GoTo Salir:
'
cnnConexion.BeginTrans

With Grid
    '
    For I = 1 To .Rows - 1
        
        If .RowHeight(I) > 0 Then
            'guarda las facturas en Cpp
            NDoc = FrmFactura.FntStrDoc
            '
            strSql = "INSERT INTO Cpp(Tipo,NDoc,Fact,CodProv,Benef,Detalle,Monto,Ivm,Total," _
            & "Frecep,Fecr,FVen,CodInm,Moneda,Estatus,Usuario,Freg,RemSSO) VALUES ('FA','" & _
            NDoc & "','" & .TextMatrix(I, 3) & "','" & gProveedor(0) & "','" & gProveedor(1) & _
            "','CAN. SSO MES " & .TextMatrix(I, 4) & " " & .TextMatrix(I, 1) & "(" & _
            .TextMatrix(I, 0) & ")','" & .TextMatrix(I, 5) & "',0,'" & .TextMatrix(I, 5) & _
            "',Date(),Date(),Date(),'" & .TextMatrix(I, 0) & "','Bs','ASIGNADO','" & gcUsuario _
            & "',Date()," & Txt(0) & ")"
            cnnConexion.Execute strSql
            '
            'efecuta la asignación del gasto
            strSql = "INSERT INTO AsignaGasto (Ndoc,CodGasto,Cargado,Descripcion,Fijo,Comun," & _
            "Alicuota,Monto,Usuario,Fecha,Hora) in '" & gcPath & "\" & .TextMatrix(I, 0) & _
            "\inm.mdb' VALUES('" & NDoc & "','" & gGasto(0) & "','01/" & Msk(0) & "','" & _
            gGasto(1) & Space(1) & Replace(.TextMatrix(I, 4), "-", "/") & "',0,-1,-1,'" & .TextMatrix(I, 5) & "','" _
            & gcUsuario & "',Date(),Time())"
            cnnConexion.Execute strSql
            '
            'agrega el registro al cargado
            strSql = "INSERT INTO Cargado (Ndoc,CodGasto,Detalle,Periodo,Monto,Fecha,Hora," _
            & "Usuario) in '" & gcPath & "\" & .TextMatrix(I, 0) & "\inm.mdb' VALUES ('" & NDoc _
            & "','" & gGasto(0) & "','" & gGasto(1) & Space(1) & Replace(.TextMatrix(I, 4), "-", "/") & "','01/" & _
            Msk(0) & "','" & .TextMatrix(I, 5) & "',Date(),Time(),'" & gcUsuario & "')"
            cnnConexion.Execute strSql
            '
        End If
        '
    Next
    '
End With
Salir:
If Err = 0 Then
    cnnConexion.CommitTrans
    Fra(0).Enabled = False
    Txt(6) = "0,00"
    Fichas.TabEnabled(1) = True
    Call Printer_Remesa(Txt(0), crPantalla)
    MsgBox "Remesa S.S.O procesada con éxito", vbInformation, App.ProductName
    Call rtnBitacora("Remesa # " & Txt(0) & " procesada con éxito")
    'impresión del reporte
    
Else

    cnnConexion.RollbackTrans
    MsgBox Err.Description, vbCritical, "Error " & Err
    Call rtnBitacora("Error " & Err.Number & " al procesar la remesa # & txt(0)")
    
End If
'
End Sub

'-------------------------------------------------------------------------------------------------
'   Rutina:     NoValida
'   ---------------------------------------------
'   Valida c/u de los datos q se encuentran en el grid (linea x linea)
'   si falta algún datos devuelve 'TRUE', de lo contrario devuelve 'FALSE'
'-------------------------------------------------------------------------------------------------
Private Function novalida() As Boolean
'variables locales
Dim rstLocal As ADODB.Recordset
Dim msg$, I&, strSql$
'
If Txt(0) = "" Then
    msg = "- Falta Nº de remesa." & vbCrLf
Else
    If CLng(Txt(0)) = 0 Then msg = "- Nº de Remesa no puede ser cero." & vbCrLf
End If
'
If Not IsDate("01/" & Msk(0)) Then msg = msg & "- Período Inválido." & vbCrLf
If gGasto(0) = "" Or gGasto(1) = "" Or gProveedor(0) = "" Or gProveedor(1) = "" Then
    msg = msg & "- Revise los parámetros de la remesa S.S.O" & vbCrLf
End If
'
'   Revisa cada una de las lineas del Grid
If msg = "" Then

With Grid
    '
    Set rstLocal = New ADODB.Recordset
    For I = 1 To .Rows - 1
        '
        If .RowHeight(I) > 0 Then
            
            'verifica q el periódo no este facturado
            strSql = "SELECT Max(Periodo) as P FROM Factura IN '" & gcPath & "\" & _
            .TextMatrix(I, 0) & "\inm.mdb' WHERE Fact Not Like 'CHD%'"
            Set rstLocal = cnnConexion.Execute(strSql)
            '
            If Not rstLocal.EOF And Not rstLocal.BOF Then
                '
                If CDate("01/" & Msk(0)) <= rstLocal!P Then
                    msg = msg & "- Período " & UCase(Format("01/" & Msk(0), "MMM-YYYY")) _
                    & " ya facturado inm: " & .TextMatrix(I, 0) & vbCrLf
                End If
            '
            End If
            rstLocal.Close
            '
        End If
        '
    Next
    '
End With
End If
If msg <> "" Then novalida = MsgBox(msg, vbCritical, App.ProductName)
Set rstLocal = Nothing
'
End Function


Private Sub CancelarEdicion()
'variables locales
Msk(0).PromptInclude = False
Msk(1).PromptInclude = False
Txt(0) = "": Txt(1) = "": Txt(2) = "": Txt(3) = "": Txt(4) = "": Txt(5) = "0,00": Msk(0) = "": Msk(1) = ""
Msk(1).PromptInclude = True
Msk(0).PromptInclude = True
Call rtnLimpiar_Grid(Grid)
Grid.Enabled = False
End Sub


Private Function NRemSSO() As Long
'variables locales
Dim rstLocal As New ADODB.Recordset
'
rstLocal.Open "SELECT Max(RemSSO) FROM Cpp", cnnConexion, adOpenKeyset, _
adLockOptimistic, adCmdText
NRemSSO = 1
If Not rstLocal.EOF And Not rstLocal.BOF Then
    NRemSSO = rstLocal.Fields(0) + NRemSSO
End If
rstLocal.Close
Set rstLocal = Nothing
'
End Function

Private Sub Printer_Remesa(N As Long, Salida As crSalida)
'variables locales
Dim rpReporte As ctlReport
On Error GoTo Salir:
Call rtnBitacora("Imprimiendo reporte remsa S.S.O")
Set rpReporte = New ctlReport
With rpReporte
'    .Reset
'    .ProgressDialog = False
    '.WindowShowProgressCtls = False
    .Reporte = gcReport & "\nom_remsso.rpt"
    .OrigenDatos(0) = gcPath & "\sac.mdb"
    .OrigenDatos(1) = gcPath & "\sac.mdb"
    .FormuladeSeleccion = "{Cpp.RemSSO}=" & N
    .Salida = Salida
    If .Salida = crPantalla Then
        '.WindowShowCloseBtn = True
        '.WindowParentHandle = FrmAdmin.hWnd
        .TituloVentana = "Remesa del S.S.O"
        '.WindowState = crptMaximized
    End If
    'errLocal = .PrintReport
    .Imprimir
End With
Set rpReporte = Nothing
Salir:
'    If errLocal <> 0 Then
'        MsgBox "Ocurrió el siguiente error mientras se emitia el reporte: " & vbCrLf _
'        & FrmAdmin.rptReporte.LastErrorString, vbCritical, "Error " & errLocal
'        Call rtnBitacora("Error " & errLocal & " al imprimir el reporte.")
'
'    ElseIf Err <> 0 Then
'        MsgBox "Ocurrió el siguiente error antes de la impresión del reporte: " & vbCrLf _
'        & Err.Description, vbCritical, "Error " & Err.Number
'        Call rtnBitacora("Error " & Err & " antes de imprimir el reporte.")
'    End If
End Sub
