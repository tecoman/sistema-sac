VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "msadodc.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form FrmTfondos 
   AutoRedraw      =   -1  'True
   Caption         =   "Tabla de Fondos"
   ClientHeight    =   420
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   4110
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "FrmTfondos.frx":0000
   KeyPreview      =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   420
   ScaleWidth      =   4110
   Tag             =   "1"
   WindowState     =   2  'Maximized
   Begin ComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   480
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4110
      _ExtentX        =   7250
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
            Object.Visible         =   0   'False
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
            Object.Visible         =   0   'False
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
      Begin VB.CheckBox chkReport 
         Caption         =   "Agregar Campo « Saldo Real »"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   4395
         TabIndex        =   73
         Top             =   75
         Visible         =   0   'False
         Width           =   3570
      End
   End
   Begin MSDataListLib.DataCombo dtcFondo 
      Height          =   315
      Index           =   0
      Left            =   2820
      TabIndex        =   42
      Top             =   1440
      Visible         =   0   'False
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   556
      _Version        =   393216
      Style           =   2
      ListField       =   "CodGasto"
      Text            =   "DataCombo1"
   End
   Begin VB.Frame fraFondo 
      BorderStyle     =   0  'None
      Height          =   1290
      Index           =   6
      Left            =   6750
      TabIndex        =   35
      Top             =   1200
      Visible         =   0   'False
      Width           =   4125
      Begin VB.ComboBox cmbFondo 
         DataField       =   "TipoMovimientoCaja"
         Height          =   315
         Index           =   0
         ItemData        =   "FrmTfondos.frx":000C
         Left            =   1065
         List            =   "FrmTfondos.frx":0037
         Style           =   2  'Dropdown List
         TabIndex        =   36
         Top             =   225
         Width           =   1320
      End
      Begin MSComCtl2.DTPicker dtpfondo 
         Height          =   315
         Index           =   0
         Left            =   1095
         TabIndex        =   37
         Top             =   795
         Width           =   1260
         _ExtentX        =   2223
         _ExtentY        =   556
         _Version        =   393216
         CalendarTitleBackColor=   -2147483646
         CalendarTitleForeColor=   -2147483639
         Format          =   50855937
         CurrentDate     =   37603
      End
      Begin VB.OptionButton optFondo 
         Caption         =   "Mes:"
         Height          =   315
         Index           =   0
         Left            =   150
         TabIndex        =   40
         Tag             =   "5"
         Top             =   225
         Width           =   1020
      End
      Begin VB.OptionButton optFondo 
         Caption         =   "Entre:"
         Height          =   315
         Index           =   1
         Left            =   120
         TabIndex        =   39
         Tag             =   "4"
         Top             =   795
         Width           =   1020
      End
      Begin VB.OptionButton optFondo 
         Caption         =   "Todos"
         Height          =   315
         Index           =   2
         Left            =   2745
         TabIndex        =   38
         Tag             =   "3"
         Top             =   225
         Value           =   -1  'True
         Width           =   1020
      End
      Begin MSComCtl2.DTPicker dtpfondo 
         Height          =   315
         Index           =   1
         Left            =   2745
         TabIndex        =   41
         Top             =   795
         Width           =   1260
         _ExtentX        =   2223
         _ExtentY        =   556
         _Version        =   393216
         CalendarTitleBackColor=   -2147483646
         CalendarTitleForeColor=   -2147483639
         CustomFormat    =   "dd/th/yy"
         Format          =   50855937
         CurrentDate     =   37603
      End
   End
   Begin TabDlg.SSTab sstFondo 
      Height          =   7200
      Left            =   675
      TabIndex        =   1
      Top             =   570
      Width           =   10695
      _ExtentX        =   18865
      _ExtentY        =   12700
      _Version        =   393216
      TabHeight       =   520
      TabCaption(0)   =   "Cuenta de Fondos"
      TabPicture(0)   =   "FrmTfondos.frx":00A0
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "fraFondo(0)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Estado de Cuenta"
      TabPicture(1)   =   "FrmTfondos.frx":00BC
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "fraFondo(2)"
      Tab(1).Control(1)=   "gridFondo(0)"
      Tab(1).Control(2)=   "fraFondo(1)"
      Tab(1).ControlCount=   3
      TabCaption(2)   =   "Cuotas Especiales"
      TabPicture(2)   =   "FrmTfondos.frx":00D8
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "gridFondo(1)"
      Tab(2).Control(1)=   "fraFondo(3)"
      Tab(2).Control(2)=   "fraFondo(4)"
      Tab(2).ControlCount=   3
      Begin VB.Frame fraFondo 
         Caption         =   "Opciones Avanzadas"
         Height          =   4920
         Index           =   4
         Left            =   -67830
         TabIndex        =   31
         Top             =   2070
         Width           =   3075
         Begin VB.Frame fraFondo 
            Caption         =   "Aplicar Filtro:"
            Height          =   1095
            Index           =   7
            Left            =   135
            TabIndex        =   43
            Top             =   1470
            Width           =   2760
            Begin VB.TextBox txtFondo 
               Alignment       =   1  'Right Justify
               BeginProperty DataFormat 
                  Type            =   1
                  Format          =   "#,##0.00"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   8202
                  SubFormatType   =   1
               EndProperty
               DataSource      =   "AdoTfondos"
               Height          =   315
               Index           =   21
               Left            =   1560
               TabIndex        =   45
               Text            =   " "
               Top             =   450
               Width           =   1005
            End
            Begin VB.CheckBox chkFiltro 
               Caption         =   "Apartamento:"
               Height          =   360
               Left            =   225
               TabIndex        =   44
               Top             =   435
               Width           =   1380
            End
         End
         Begin VB.Frame fraFondo 
            Height          =   1080
            Index           =   5
            Left            =   120
            TabIndex        =   32
            Top             =   210
            Width           =   2790
            Begin VB.OptionButton optavan 
               Caption         =   "Por Cobrar"
               Height          =   285
               Index           =   1
               Left            =   195
               TabIndex        =   34
               Top             =   630
               Width           =   1515
            End
            Begin VB.OptionButton optavan 
               Caption         =   "Cobrado"
               Height          =   285
               Index           =   0
               Left            =   195
               TabIndex        =   33
               Top             =   225
               Value           =   -1  'True
               Width           =   1515
            End
         End
      End
      Begin VB.Frame fraFondo 
         Caption         =   "Opciones: "
         ClipControls    =   0   'False
         Height          =   1515
         Index           =   3
         Left            =   -74805
         TabIndex        =   22
         Top             =   510
         Width           =   10035
         Begin VB.TextBox txtFondo 
            Alignment       =   1  'Right Justify
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "#,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   8202
               SubFormatType   =   1
            EndProperty
            DataSource      =   "AdoTfondos"
            Height          =   315
            Index           =   20
            Left            =   1950
            Locked          =   -1  'True
            TabIndex        =   25
            Text            =   " "
            Top             =   945
            Width           =   1260
         End
         Begin VB.TextBox txtFondo 
            Alignment       =   1  'Right Justify
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "#,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   8202
               SubFormatType   =   1
            EndProperty
            DataSource      =   "AdoTfondos"
            Height          =   315
            Index           =   19
            Left            =   4455
            Locked          =   -1  'True
            TabIndex        =   24
            Text            =   " "
            Top             =   375
            Width           =   1260
         End
         Begin VB.TextBox txtFondo 
            Alignment       =   1  'Right Justify
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "#,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   8202
               SubFormatType   =   1
            EndProperty
            DataSource      =   "AdoTfondos"
            Height          =   315
            Index           =   18
            Left            =   4455
            Locked          =   -1  'True
            TabIndex        =   23
            Text            =   " "
            Top             =   945
            Width           =   1260
         End
         Begin VB.Label label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Cod. Fondo"
            Height          =   270
            Index           =   26
            Left            =   240
            TabIndex        =   29
            Top             =   420
            Width           =   1620
         End
         Begin VB.Label label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Disponible al:"
            Height          =   435
            Index           =   24
            Left            =   255
            TabIndex        =   28
            Top             =   900
            Width           =   1665
         End
         Begin VB.Label label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Recaudado:"
            Height          =   270
            Index           =   23
            Left            =   3315
            TabIndex        =   27
            Top             =   390
            Width           =   1050
         End
         Begin VB.Label label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Pagado:"
            Height          =   270
            Index           =   21
            Left            =   3315
            TabIndex        =   26
            Top             =   960
            Width           =   780
         End
      End
      Begin VB.Frame fraFondo 
         Caption         =   "Opciones: "
         ClipControls    =   0   'False
         Height          =   1515
         Index           =   1
         Left            =   -74805
         TabIndex        =   4
         Top             =   510
         Width           =   10035
         Begin VB.TextBox txtFondo 
            Alignment       =   1  'Right Justify
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "#,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   8202
               SubFormatType   =   1
            EndProperty
            DataSource      =   "AdoTfondos"
            Height          =   315
            Index           =   15
            Left            =   4125
            Locked          =   -1  'True
            TabIndex        =   12
            Text            =   " "
            Top             =   945
            Width           =   1260
         End
         Begin VB.TextBox txtFondo 
            Alignment       =   1  'Right Justify
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "#,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   8202
               SubFormatType   =   1
            EndProperty
            DataSource      =   "AdoTfondos"
            Height          =   315
            Index           =   14
            Left            =   4125
            Locked          =   -1  'True
            TabIndex        =   11
            Text            =   " "
            Top             =   375
            Width           =   1260
         End
         Begin VB.TextBox txtFondo 
            Alignment       =   1  'Right Justify
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "#,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   8202
               SubFormatType   =   1
            EndProperty
            DataSource      =   "AdoTfondos"
            Height          =   315
            Index           =   13
            Left            =   1950
            Locked          =   -1  'True
            TabIndex        =   7
            Text            =   " "
            Top             =   945
            Width           =   1260
         End
         Begin VB.Label label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Débitos:"
            Height          =   270
            Index           =   17
            Left            =   3480
            TabIndex        =   10
            Top             =   967
            Width           =   615
         End
         Begin VB.Label label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Créditos:"
            Height          =   270
            Index           =   16
            Left            =   3480
            TabIndex        =   9
            Top             =   397
            Width           =   645
         End
         Begin VB.Label label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Saldo al:"
            Height          =   270
            Index           =   15
            Left            =   255
            TabIndex        =   6
            Top             =   990
            Width           =   1665
         End
         Begin VB.Label label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Cod. Fondo"
            Height          =   270
            Index           =   13
            Left            =   240
            TabIndex        =   5
            Top             =   420
            Width           =   1620
         End
      End
      Begin VB.Frame fraFondo 
         ClipControls    =   0   'False
         Height          =   6495
         Index           =   0
         Left            =   180
         TabIndex        =   2
         Top             =   435
         Width           =   10350
         Begin MSDataGridLib.DataGrid DataGrid1 
            Height          =   3630
            Left            =   120
            TabIndex        =   3
            Top             =   360
            Width           =   10080
            _ExtentX        =   17780
            _ExtentY        =   6403
            _Version        =   393216
            AllowUpdate     =   -1  'True
            HeadLines       =   1
            RowHeight       =   15
            RowDividerStyle =   4
            FormatLocked    =   -1  'True
            AllowAddNew     =   -1  'True
            AllowDelete     =   -1  'True
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
            Caption         =   "Tabla de Fondos Especiales"
            ColumnCount     =   7
            BeginProperty Column00 
               DataField       =   "CodGasto"
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
               DataField       =   "Titulo"
               Caption         =   "Título de la Cuenta"
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
               DataField       =   "Fijo"
               Caption         =   "Fijo"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   5
                  Format          =   ""
                  HaveTrueFalseNull=   1
                  TrueValue       =   "S"
                  FalseValue      =   "N"
                  NullValue       =   ""
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   8202
                  SubFormatType   =   7
               EndProperty
            EndProperty
            BeginProperty Column03 
               DataField       =   "Comun"
               Caption         =   "Común"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   5
                  Format          =   ""
                  HaveTrueFalseNull=   1
                  TrueValue       =   "S"
                  FalseValue      =   "N"
                  NullValue       =   ""
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   8202
                  SubFormatType   =   7
               EndProperty
            EndProperty
            BeginProperty Column04 
               DataField       =   "Alicuota"
               Caption         =   "Alícuota"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   5
                  Format          =   ""
                  HaveTrueFalseNull=   1
                  TrueValue       =   "S"
                  FalseValue      =   "N"
                  NullValue       =   ""
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   8202
                  SubFormatType   =   7
               EndProperty
            EndProperty
            BeginProperty Column05 
               DataField       =   "MontoFijo"
               Caption         =   "Monto Fijo"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   1
                  Format          =   "#,##0.00  "
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   8202
                  SubFormatType   =   0
               EndProperty
            EndProperty
            BeginProperty Column06 
               DataField       =   "Fondo"
               Caption         =   "Fondo"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   5
                  Format          =   ""
                  HaveTrueFalseNull=   1
                  TrueValue       =   "S"
                  FalseValue      =   "N"
                  NullValue       =   ""
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   8202
                  SubFormatType   =   7
               EndProperty
            EndProperty
            SplitCount      =   1
            BeginProperty Split0 
               AllowRowSizing  =   0   'False
               AllowSizing     =   0   'False
               BeginProperty Column00 
                  ColumnWidth     =   705,26
               EndProperty
               BeginProperty Column01 
                  ColumnWidth     =   4710,047
               EndProperty
               BeginProperty Column02 
                  ColumnWidth     =   434,835
               EndProperty
               BeginProperty Column03 
                  ColumnWidth     =   689,953
               EndProperty
               BeginProperty Column04 
                  ColumnWidth     =   870,236
               EndProperty
               BeginProperty Column05 
                  Alignment       =   1
               EndProperty
               BeginProperty Column06 
                  ColumnWidth     =   14,74
               EndProperty
            EndProperty
         End
         Begin VB.Frame fraFondo 
            BorderStyle     =   0  'None
            Height          =   2145
            Index           =   8
            Left            =   105
            TabIndex        =   46
            Top             =   4170
            Width           =   10125
            Begin VB.TextBox txtFondo 
               Alignment       =   1  'Right Justify
               DataField       =   "SaldoActual"
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
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Index           =   12
               Left            =   7590
               TabIndex        =   59
               Text            =   " "
               Top             =   1725
               Width           =   2460
            End
            Begin VB.TextBox txtFondo 
               Alignment       =   1  'Right Justify
               DataField       =   "Saldo3"
               BeginProperty DataFormat 
                  Type            =   1
                  Format          =   "#,##0.00"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   8202
                  SubFormatType   =   0
               EndProperty
               DataSource      =   "AdoTfondos"
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
               Index           =   6
               Left            =   6060
               Locked          =   -1  'True
               TabIndex        =   58
               Text            =   " "
               Top             =   240
               Width           =   1515
            End
            Begin VB.TextBox txtFondo 
               Alignment       =   1  'Right Justify
               DataField       =   "Saldo4"
               BeginProperty DataFormat 
                  Type            =   1
                  Format          =   "#,##0.00"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   8202
                  SubFormatType   =   0
               EndProperty
               DataSource      =   "AdoTfondos"
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
               Index           =   9
               Left            =   8550
               Locked          =   -1  'True
               TabIndex        =   57
               Text            =   " "
               Top             =   240
               Width           =   1515
            End
            Begin VB.TextBox txtFondo 
               Alignment       =   1  'Right Justify
               DataField       =   "Saldo1"
               BeginProperty DataFormat 
                  Type            =   1
                  Format          =   "#,##0.00"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   8202
                  SubFormatType   =   0
               EndProperty
               DataSource      =   "AdoTfondos"
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
               Index           =   0
               Left            =   1095
               Locked          =   -1  'True
               TabIndex        =   56
               Text            =   " "
               Top             =   240
               Width           =   1515
            End
            Begin VB.TextBox txtFondo 
               Alignment       =   1  'Right Justify
               DataField       =   "Saldo2"
               BeginProperty DataFormat 
                  Type            =   1
                  Format          =   "#,##0.00"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   8202
                  SubFormatType   =   0
               EndProperty
               DataSource      =   "AdoTfondos"
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
               Index           =   3
               Left            =   3465
               Locked          =   -1  'True
               TabIndex        =   55
               Text            =   " "
               Top             =   240
               Width           =   1515
            End
            Begin VB.TextBox txtFondo 
               Alignment       =   1  'Right Justify
               DataField       =   "Saldo12"
               BeginProperty DataFormat 
                  Type            =   1
                  Format          =   "#,##0.00"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   8202
                  SubFormatType   =   0
               EndProperty
               DataSource      =   "AdoTfondos"
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
               Index           =   11
               Left            =   8550
               Locked          =   -1  'True
               TabIndex        =   54
               Text            =   " "
               Top             =   1140
               Width           =   1500
            End
            Begin VB.TextBox txtFondo 
               Alignment       =   1  'Right Justify
               DataField       =   "Saldo7"
               BeginProperty DataFormat 
                  Type            =   1
                  Format          =   "#,##0.00"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   8202
                  SubFormatType   =   0
               EndProperty
               DataSource      =   "AdoTfondos"
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
               Index           =   7
               Left            =   6060
               Locked          =   -1  'True
               TabIndex        =   53
               Text            =   " "
               Top             =   690
               Width           =   1515
            End
            Begin VB.TextBox txtFondo 
               Alignment       =   1  'Right Justify
               DataField       =   "Saldo11"
               BeginProperty DataFormat 
                  Type            =   1
                  Format          =   "#,##0.00"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   8202
                  SubFormatType   =   0
               EndProperty
               DataSource      =   "AdoTfondos"
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
               Left            =   6060
               Locked          =   -1  'True
               TabIndex        =   52
               Text            =   " "
               Top             =   1140
               Width           =   1500
            End
            Begin VB.TextBox txtFondo 
               Alignment       =   1  'Right Justify
               DataField       =   "Saldo8"
               BeginProperty DataFormat 
                  Type            =   1
                  Format          =   "#,##0.00"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   8202
                  SubFormatType   =   0
               EndProperty
               DataSource      =   "AdoTfondos"
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
               Index           =   10
               Left            =   8550
               Locked          =   -1  'True
               TabIndex        =   51
               Text            =   " "
               Top             =   690
               Width           =   1515
            End
            Begin VB.TextBox txtFondo 
               Alignment       =   1  'Right Justify
               DataField       =   "Saldo10"
               BeginProperty DataFormat 
                  Type            =   1
                  Format          =   "#,##0.00"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   8202
                  SubFormatType   =   0
               EndProperty
               DataSource      =   "AdoTfondos"
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
               Index           =   5
               Left            =   3465
               Locked          =   -1  'True
               TabIndex        =   50
               Text            =   " "
               Top             =   1140
               Width           =   1500
            End
            Begin VB.TextBox txtFondo 
               Alignment       =   1  'Right Justify
               DataField       =   "Saldo6"
               BeginProperty DataFormat 
                  Type            =   1
                  Format          =   "#,##0.00"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   8202
                  SubFormatType   =   0
               EndProperty
               DataSource      =   "AdoTfondos"
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
               Index           =   4
               Left            =   3465
               Locked          =   -1  'True
               TabIndex        =   49
               Text            =   " "
               Top             =   690
               Width           =   1515
            End
            Begin VB.TextBox txtFondo 
               Alignment       =   1  'Right Justify
               DataField       =   "Saldo9"
               BeginProperty DataFormat 
                  Type            =   1
                  Format          =   "#,##0.00"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   8202
                  SubFormatType   =   0
               EndProperty
               DataSource      =   "AdoTfondos"
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
               Index           =   2
               Left            =   1095
               Locked          =   -1  'True
               TabIndex        =   48
               Text            =   " "
               Top             =   1140
               Width           =   1500
            End
            Begin VB.TextBox txtFondo 
               Alignment       =   1  'Right Justify
               DataField       =   "Saldo5"
               BeginProperty DataFormat 
                  Type            =   1
                  Format          =   "#,##0.00"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   8202
                  SubFormatType   =   0
               EndProperty
               DataSource      =   "AdoTfondos"
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
               Index           =   1
               Left            =   1095
               Locked          =   -1  'True
               TabIndex        =   47
               Text            =   " "
               Top             =   690
               Width           =   1515
            End
            Begin VB.Label label1 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "Mayo :"
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
               Left            =   30
               TabIndex        =   72
               Top             =   735
               Width           =   1065
            End
            Begin VB.Label label1 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "Septiembre :"
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
               Left            =   30
               TabIndex        =   71
               Top             =   1185
               Width           =   1065
            End
            Begin VB.Label label1 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "Enero :"
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
               Left            =   30
               TabIndex        =   70
               Top             =   285
               Width           =   1065
            End
            Begin VB.Label label1 
               AutoSize        =   -1  'True
               Caption         =   "Saldo Actual :"
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
               Left            =   6360
               TabIndex        =   69
               Top             =   1785
               Width           =   1125
            End
            Begin VB.Label label1 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "Marzo :"
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
               Left            =   5070
               TabIndex        =   68
               Top             =   285
               Width           =   990
            End
            Begin VB.Label label1 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "Abril :"
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
               Left            =   7605
               TabIndex        =   67
               Top             =   285
               Width           =   915
            End
            Begin VB.Label label1 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "Febrero :"
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
               Left            =   2655
               TabIndex        =   66
               Top             =   285
               Width           =   795
            End
            Begin VB.Label label1 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "Diciembre :"
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
               Index           =   11
               Left            =   7605
               TabIndex        =   65
               Top             =   1185
               Width           =   915
            End
            Begin VB.Label label1 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "Julio :"
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
               Left            =   5070
               TabIndex        =   64
               Top             =   735
               Width           =   990
            End
            Begin VB.Label label1 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "Noviembre :"
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
               Left            =   5070
               TabIndex        =   63
               Top             =   1185
               Width           =   990
            End
            Begin VB.Label label1 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "Agosto :"
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
               Left            =   7605
               TabIndex        =   62
               Top             =   735
               Width           =   915
            End
            Begin VB.Label label1 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "Octubre :"
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
               Left            =   2655
               TabIndex        =   61
               Top             =   1185
               Width           =   795
            End
            Begin VB.Label label1 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "Junio :"
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
               Left            =   2655
               TabIndex        =   60
               Top             =   735
               Width           =   795
            End
         End
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid gridFondo 
         Height          =   4815
         Index           =   0
         Left            =   -74805
         TabIndex        =   8
         Tag             =   "900|400|4000|1450|1450|1450|1450"
         Top             =   2175
         Width           =   10065
         _ExtentX        =   17754
         _ExtentY        =   8493
         _Version        =   393216
         ForeColor       =   -2147483646
         Cols            =   6
         FixedCols       =   0
         BackColorFixed  =   -2147483646
         ForeColorFixed  =   -2147483643
         BackColorSel    =   65280
         ForeColorSel    =   0
         BackColorBkg    =   -2147483636
         GridColor       =   -2147483633
         GridLinesFixed  =   1
         SelectionMode   =   1
         AllowUserResizing=   1
         MousePointer    =   99
         FormatString    =   "Fecha |Tp |<Concepto |>Debe |>Haber |>Saldo"
         MouseIcon       =   "FrmTfondos.frx":00F4
         _NumberOfBands  =   1
         _Band(0).Cols   =   6
         _Band(0).TextStyleBand=   0
         _Band(0).TextStyleHeader=   0
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid gridFondo 
         Height          =   4815
         Index           =   1
         Left            =   -74775
         TabIndex        =   30
         Tag             =   "250|1000|1000|1000|1450|1450"
         Top             =   2205
         Width           =   6855
         _ExtentX        =   12091
         _ExtentY        =   8493
         _Version        =   393216
         ForeColor       =   -2147483646
         Cols            =   6
         FixedCols       =   0
         BackColorFixed  =   -2147483646
         ForeColorFixed  =   -2147483643
         BackColorSel    =   65280
         ForeColorSel    =   0
         BackColorBkg    =   -2147483636
         GridColor       =   -2147483633
         GridLinesFixed  =   1
         AllowUserResizing=   1
         BorderStyle     =   0
         MousePointer    =   99
         FormatString    =   "|^Fecha |^Apto |^Período|>Monto |>Saldo"
         MouseIcon       =   "FrmTfondos.frx":0256
         _NumberOfBands  =   1
         _Band(0).Cols   =   6
         _Band(0).TextStyleBand=   0
         _Band(0).TextStyleHeader=   0
      End
      Begin VB.Frame fraFondo 
         Caption         =   "Agregar Movimiento:"
         Height          =   1095
         Index           =   2
         Left            =   -74790
         TabIndex        =   13
         Top             =   2175
         Width           =   9405
         Begin VB.TextBox txtFondo 
            Alignment       =   1  'Right Justify
            Height          =   315
            Index           =   17
            Left            =   7920
            TabIndex        =   16
            Text            =   "0,00"
            Top             =   585
            Width           =   1335
         End
         Begin VB.TextBox txtFondo 
            Height          =   315
            Index           =   16
            Left            =   2220
            TabIndex        =   15
            Top             =   585
            Width           =   5700
         End
         Begin VB.ComboBox cmb 
            Height          =   315
            ItemData        =   "FrmTfondos.frx":03B8
            Left            =   1320
            List            =   "FrmTfondos.frx":03C8
            Style           =   2  'Dropdown List
            TabIndex        =   14
            Top             =   585
            Width           =   900
         End
         Begin MSMask.MaskEdBox msk 
            Bindings        =   "FrmTfondos.frx":03DC
            Height          =   315
            Left            =   150
            TabIndex        =   17
            TabStop         =   0   'False
            Top             =   585
            Width           =   1155
            _ExtentX        =   2037
            _ExtentY        =   556
            _Version        =   393216
            MaxLength       =   12
            Format          =   "dd/MM/yyyy"
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin VB.Label label1 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Monto"
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   22
            Left            =   7935
            TabIndex        =   21
            Top             =   330
            Width           =   1320
         End
         Begin VB.Label label1 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Concepto"
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   20
            Left            =   2220
            TabIndex        =   20
            Top             =   330
            Width           =   5715
         End
         Begin VB.Label label1 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Tipo"
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   19
            Left            =   1320
            TabIndex        =   19
            Top             =   330
            Width           =   900
         End
         Begin VB.Label label1 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Fecha:"
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   18
            Left            =   150
            TabIndex        =   18
            Top             =   330
            Width           =   1170
         End
      End
   End
   Begin MSAdodcLib.Adodc AdoTfondos 
      Height          =   330
      Left            =   1290
      Top             =   7140
      Visible         =   0   'False
      Width           =   2205
      _ExtentX        =   3889
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
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
      Caption         =   "AdoTgastos"
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
   Begin VB.Image img 
      Height          =   135
      Index           =   0
      Left            =   915
      Picture         =   "FrmTfondos.frx":03FE
      Top             =   10
      Width           =   135
   End
   Begin VB.Image img 
      Height          =   135
      Index           =   1
      Left            =   1200
      Picture         =   "FrmTfondos.frx":045F
      Top             =   10
      Width           =   135
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
            Picture         =   "FrmTfondos.frx":04BD
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmTfondos.frx":063F
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmTfondos.frx":07C1
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmTfondos.frx":0943
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmTfondos.frx":0AC5
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmTfondos.frx":0C47
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmTfondos.frx":0DC9
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmTfondos.frx":0F4B
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmTfondos.frx":10CD
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmTfondos.frx":124F
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmTfondos.frx":13D1
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmTfondos.frx":1553
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "FrmTfondos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
    'variables publicas a nivel de módulo
    Dim Titulo As String
    Dim strDet(1) As String
    
    Private Sub AdoTfondos_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, _
    ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, _
    ByVal pRecordset As ADODB.Recordset)
    On Error Resume Next
    Titulo = AdoTfondos.Recordset!codGasto & " - " & AdoTfondos.Recordset!Titulo
    End Sub

    Private Sub chkFiltro_Click()
    'variables locales
    Dim rstDet As ADODB.Recordset
    Dim curSaldo As Currency
    '
    If chkFiltro.Value = vbChecked Then
    
        Set rstDet = New ADODB.Recordset
        rstDet.CursorLocation = adUseClient
        rstDet.Open strDet(0), AdoTfondos.ConnectionString, adOpenKeyset, adLockOptimistic, _
        adCmdText
        rstDet.Sort = "Periodo DESC"
        rstDet.Filter = "CodProp='" & txtFondo(21) & "'"
        
        gridFondo(1).Rows = 2
        Call rtnLimpiar_Grid(gridFondo(1))
        If Not rstDet.BOF And Not rstDet.EOF Then
            rstDet.MoveFirst
            gridFondo(1).Rows = rstDet.RecordCount + 1
            Do
                
                I = I + 1
                gridFondo(1).TextMatrix(I, 1) = rstDet("Freg")
                gridFondo(1).TextMatrix(I, 2) = rstDet("CodProp")
                gridFondo(1).TextMatrix(I, 3) = Format(rstDet("Periodo"), "MM/YYYY")
                gridFondo(1).TextMatrix(I, 4) = _
                Format(rstDet("Total"), "#,##0.00")
                curSaldo = curSaldo + rstDet("Total")
                gridFondo(1).TextMatrix(I, 5) = Format(curSaldo, "#,##0.00")
                rstDet.MoveNext
                
            Loop Until rstDet.EOF
            
        End If
        
    ElseIf chkFiltro.Value = vbUnchecked Then
    
        Call EdoCta_Fondo(AdoTfondos.Recordset("CodGasto"))
        
    End If
    '
    End Sub

    Private Sub cmbFondo_Click(Index As Integer)
    If optFondo(0) Then Call EdoCta_Fondo(AdoTfondos.Recordset("CodGasto"))
    End Sub

    '---------------------------------------------------------------------------------------------
    Private Sub dtcFondo_Click(Index As Integer, Area As Integer) '
    '---------------------------------------------------------------------------------------------
    If Area = 2 Then
    '
        With AdoTfondos.Recordset
            '
            If Not .EOF Or Not .BOF Then
                .MoveFirst
                .Find "CodGasto='" & dtcFondo(IIf(Index = 0, 0, 1)) & "'"
                If Not .EOF Then
                    'Titulo = .Fields("CodGasto") & " - " & .Fields("Titulo")
                    Call EdoCta_Fondo(!codGasto)
                    gridFondo(0).Redraw = True
                    gridFondo(1).Redraw = True
                End If
                '
            End If
            '
        End With
        '
    End If
    '
    End Sub


    Private Sub dtpfondo_Change(Index As Integer)
    dtpfondo(0).Refresh: dtpfondo(1).Refresh
    If dtpfondo(0).Value <= dtpfondo(1).Value Then
        Call EdoCta_Fondo(AdoTfondos.Recordset!codGasto)
        gridFondo(0).Redraw = True
        gridFondo(1).Redraw = True
    End If
    End Sub

    '---------------------------------------------------------------------------------------------
    Private Sub dtpfondo_CloseUp(Index As Integer)  '
    '---------------------------------------------------------------------------------------------
    '
    dtpfondo(0).Refresh: dtpfondo(1).Refresh
    '
    If dtpfondo(0).Value <= dtpfondo(1).Value And optFondo(1) Then
        Call EdoCta_Fondo(AdoTfondos.Recordset!codGasto)
        gridFondo(0).Redraw = True
        gridFondo(1).Redraw = True
    End If
    '
    End Sub

    '---------------------------------------------------------------------------------------------
    '   Carga el Formulario
    '---------------------------------------------------------------------------------------------
    Private Sub Form_Load()
    '
    On Error GoTo Cerrar
    '
    If mcDatos <> "" Then
    '
        AdoTfondos.ConnectionString = cnnOLEDB + mcDatos
        AdoTfondos.RecordSource = "SELECT * FROM tgastos WHERE fondo = true"
        AdoTfondos.Refresh
        Set DataGrid1.DataSource = AdoTfondos
        Set txtFondo(12).DataSource = AdoTfondos
        'configura la presentación del FlexGrid
        Set dtcFondo(I).RowSource = AdoTfondos
        cmbFondo(0).Text = cmbFondo(I).List(Month(Date) - 1)
        For I = 0 To 1
        '
            With gridFondo(I)
                Set .FontFixed = LetraTitulo(LoadResString(527), 7.5, , True)
                Set .Font = LetraTitulo(LoadResString(528), 8)
                Call centra_titulo(gridFondo(I), True)
                Set DataGrid1.Font = .Font
                '
            End With
            '
        Next
        Set DataGrid1.HeadFont = LetraTitulo(LoadResString(527), 7.5, True)
    End If
    Call RtnEstado(6, Toolbar1)
    Toolbar1.Buttons("New").Enabled = False
Cerrar:
    If Err.Number <> 0 Then Unload Me
    End Sub
    
    
    
    Private Sub Form_Resize()
    '
    'provedimiento para reorganizar los controles en el formulario
    '
    On Error Resume Next
    
    With sstFondo
    
        .Left = (Me.ScaleWidth - .Width) / 2
        .Height = Me.ScaleHeight - .Top - 200
        fraFondo(6).Left = .Left + (6680 - 675)
        dtcFondo(0).Left = .Left + (2820 - 675)
        fraFondo(0).Height = .Height - fraFondo(0).Top - 200
        DataGrid1.Height = fraFondo(0).Height - fraFondo(8).Height - 500
        fraFondo(8).Top = DataGrid1.Top + DataGrid1.Height
        gridFondo(0).Height = .Height - gridFondo(0).Top - 200
        'gridFondo(1).Left = gridFondo(0).Left
        'gridFondo(1).Top = gridFondo(0).Top
        gridFondo(1).Height = gridFondo(0).Height
       '
    End With
    
    End Sub

    Private Sub gridFondo_MouseDown(Index%, Button%, Shift%, X As Single, Y As Single)
    'varibles locales
    Dim rstDet As ADODB.Recordset
    Dim Linea&, INI&
    Dim strF2 As String
    Dim curSA As Currency
    '
    If Index = 0 And Button = 2 Then PopupMenu FrmAdmin.Mante
    If Index = 1 And gridFondo(1).Col = 0 And gridFondo(1).RowSel >= 2 Then
    'amplia el grid
        With gridFondo(1)
            .Col = 0
            .Row = .RowSel
            If .CellPicture = img(0) Then   'plus
                curSA = IIf(.TextMatrix(.Row - 1, 5) = "", 0, .TextMatrix(.Row - 1, 5))
                'coloca en negrita la columna seleccionada
                .ColSel = .Cols - 1
                .FillStyle = flexFillRepeat
                .CellFontBold = True
                .CellFontItalic = True
                .FillStyle = flexFillSingle
                'cambia la imagen
                Set .CellPicture = img(1)
                .CellPictureAlignment = flexAlignCenterCenter
                Set rstDet = New ADODB.Recordset
                rstDet.Open strDet(0), AdoTfondos.ConnectionString, adOpenKeyset, _
                adLockOptimistic, adCmdText
                strF2 = DateAdd("d", -1, DateAdd("m", 1, CDate(.TextMatrix(.Row, 1))))
                rstDet.Filter = "Freg >=#" & .TextMatrix(.Row, 1) & "# AND Freg <=#" & strF2 & "#"
                If Not rstDet.EOF And Not rstDet.BOF Then
                    rstDet.MoveFirst
                    Linea = .Row + 1
                    INI = Linea
                    Do
                        
                        .AddItem "", Linea
                        .RowHeight(Linea) = 200
                        '.TextMatrix(Linea, 1) = rstDet("Freg")
                        .TextMatrix(Linea, 2) = rstDet("CodProp")
                        .TextMatrix(Linea, 3) = Format(rstDet("Periodo"), "MM/YYYY")
                        .TextMatrix(Linea, 4) = Format(rstDet("Total"), "#,##0.00")
                        curSA = rstDet("Total") + curSA
                        .TextMatrix(Linea, 5) = Format(curSA, "#,##0.00")
                        rstDet.MoveNext
                        Linea = Linea + 1
                    Loop Until rstDet.EOF
                    rstDet.Close
                    Set rstDet = Nothing
                    .Col = 0
                    .Row = INI
                    .ColSel = .Cols - 1
                    .RowSel = Linea - 1
                    .FillStyle = flexFillRepeat
                    .CellFontSize = 7.5
                    .FillStyle = flexFillSingle
                End If
                
            ElseIf .CellPicture = img(1) Then   'minus
                'quita la negrita
                .ColSel = .Cols - 1
                .FillStyle = flexFillRepeat
                .CellFontBold = False
                .CellFontItalic = False
                .FillStyle = flexFillSingle
                '
                'cambia la imagen
                Set .CellPicture = img(0)
                .CellPictureAlignment = flexAlignCenterCenter
                '
                'elimina las celdas del detalle
                Linea = .Row + 1
                .Row = Linea
                Do Until .CellPicture = img(0)
                    .Row = Linea
                    If .CellPicture = 0 Then .RemoveItem (Linea)
                Loop
                .Row = Linea - 1
            End If
            '
        End With
    End If
    '
    End Sub

    Private Sub optavan_Click(Index As Integer)
    If Index = 0 Then
        Label1(23) = "Recaudado:"
    ElseIf Index = 1 Then
        Label1(23) = "Por Recaudar:"
    End If
    Call EdoCta_Fondo(AdoTfondos.Recordset("CodGasto"))
    End Sub

    Private Sub optFondo_Click(Index As Integer)
    Call EdoCta_Fondo(AdoTfondos.Recordset("CodGasto"))
    End Sub

    '---------------------------------------------------------------------------------------------
    '   Maneja los procedimientos que responden al hacer click sobre una
    '   ficha de este control
    '---------------------------------------------------------------------------------------------
    Private Sub sstFondo_Click(PreviousTab As Integer)
    '
    Dim I As Integer
    With AdoTfondos.Recordset
    '
        Select Case sstFondo.Tab
        '
            'Cuentas de Fondos
            Case 0
            Toolbar1.Buttons("New").Enabled = False
            chkReport.Visible = False
            For I = 1 To 4: Toolbar1.Buttons(I).Enabled = True
            Next
            fraFondo(6).Visible = False
            dtcFondo(0).Visible = False
            '-------------------------
            
            Case 1, 2 'Estado de cuenta
            '-------------------------
                chkReport.Visible = True
                For I = 1 To 4: Toolbar1.Buttons(I).Enabled = False
                Next
                If Not .EOF And Not .BOF Then
                    dtcFondo(0).Text = !codGasto
                    Call EdoCta_Fondo(!codGasto)
                    gridFondo(0).Redraw = True
                    gridFondo(1).Redraw = True
                End If
                Toolbar1.Buttons("New").Enabled = True
                fraFondo(6).Visible = True
                dtcFondo(0).Visible = True
                '
        End Select
        '
    End With
    '
    End Sub


    '---------------------------------------------------------------------------------------------
    '   Maneja los procedimientos que responden al evento click de la barra de
    '   herramientas, utilizando la propiedad Key de cada Boton
    '---------------------------------------------------------------------------------------------
    Private Sub Toolbar1_ButtonClick(ByVal Button As ComctlLib.Button)
    'variables locales
    Dim lngDistance As Long
    
    With AdoTfondos.Recordset
    '
        Select Case UCase(Button.Key)
        '
            Case "FIRST"    'Primer Registro
            '----------------
                If Not .EOF Or .BOF Then .MoveFirst
                '
            Case "PREVIOUS" 'Registro Previo
            '----------------
                If Not .EOF Or .BOF Then .MovePrevious
                If .BOF Then .MoveLast
                '
            Case "NEXT"     'Siguiente Registro
            '----------------
                If Not .EOF Or .BOF Then .MoveNext
                If .EOF Then .MoveFirst
                '
            Case "END" 'Último Registro
            '----------------
                If Not .EOF Or .BOF Then .MoveLast
            
            Case "NEW"  'agregar movimiento al fondo
            '---------------
                Call RtnEstado(Button.Index, Toolbar1)
                lngDistance = fraFondo(2).Height + 200
                gridFondo(0).Top = gridFondo(0).Top + lngDistance
                gridFondo(0).Height = gridFondo(0).Height - lngDistance
                'limpiar los controles de edición
                msk.PromptInclude = False
                msk = ""
                msk.PromptInclude = True
                cmb.ListIndex = -1
                txtFondo(16) = ""
                txtFondo(17) = "0,00"
                msk.SetFocus
                Call rtnBitacora("Agregar movimiento cta. " & dtcFondo(0))
                '--------------
            Case "UNDO" 'cancelar
            '---------------
                Call RtnEstado(Button.Index, Toolbar1)
                lngDistance = fraFondo(2).Height + 200
                gridFondo(0).Top = gridFondo(0).Top - lngDistance
                gridFondo(0).Height = gridFondo(0).Height + lngDistance
                Call rtnBitacora("Transacción cancelada por el usuario...")
                
            Case "SAVE" 'guardar movimiento
            '---------------
                If agregar Then
                    
                    lngDistance = fraFondo(2).Height + 200
                    gridFondo(0).Top = gridFondo(0).Top - lngDistance
                    gridFondo(0).Height = gridFondo(0).Height + lngDistance
                    Call RtnEstado(Button.Index, Toolbar1)
                    Call EdoCta_Fondo(!codGasto)
                    
                End If
                
                '
            Case "CLOSE"    'Cerrar
            '----------------
                Unload Me
                Set FrmTfondos = Nothing
                '
            Case "PRINT"    'Imprimir
            '----------------
                Dim datFecha1 As Date
                Dim datFecha2 As Date
                Dim errLocal As Long
                Dim img As String
                Dim rpReporte As ctlReport
                '
                MousePointer = vbHourglass
                If sstFondo.Tab = 0 Then
                    mcTitulo = "Catálogo de Cuentas de Fondos"
                    mcReport = "ListaGas.Rpt"
                    mcOrdCod = "+{Tgastos.CodGasto}"
                    mcOrdAlfa = "+{Tgastos.Titulo}"
                    mcCrit = "{Tgastos.Fondo}"
                    FrmReport.Show
                
                ElseIf sstFondo.Tab = 1 Then
                    
                    If chkReport.Value = vbChecked And optavan(0) Then
                        optavan(1).Value = True
                        Call EdoCta_Fondo(dtcFondo(0))
                    End If
                    If gridFondo(0).Rows > 1 Then
                    
                        Dim cnnFondo As New ADODB.Connection
                        cnnFondo.Open AdoTfondos.ConnectionString
                        cnnFondo.Execute "DELETE FROM EdoCtaF;"
                        Dim I As Integer
                        Dim strFecha As String
                        I = 1
                        
                        With gridFondo(0)
                            
                            strFecha = IIf(IsDate(.TextMatrix(2, 0)) = False, _
                            DateAdd("yyyy", -5, Date), DateAdd("d", -1, .TextMatrix(2, 0)))
                            Do Until I = .Rows
                                'Debug.Print i
                                cnnFondo.Execute "INSERT INTO EdoCtaF(Fecha,Tp,Descripcion,Debe" _
                                & ",Haber,Saldo) VALUES('" & IIf(.TextMatrix(I, 0) = "", strFecha, .TextMatrix(I, 0)) & "','" _
                                & .TextMatrix(I, 1) & "','" & .TextMatrix(I, 2) & "','" _
                                & .TextMatrix(I, 3) & "','" & .TextMatrix(I, 4) & "','" _
                                & .TextMatrix(I, 5) & "');"
                                I = I + 1
                                If I = 1 Then datFecha1 = IIf(.TextMatrix(I, 0) = "", strFecha, .TextMatrix(I, 0))
                                If I = .Rows - 1 Then datFecha2 = .TextMatrix(I, 0)
                            Loop
                            cnnFondo.Execute "INSERT INTO EdoCtaF(Fecha,Tp,Descripcion,Debe" _
                            & ",Haber,Saldo) VALUES(DATE(),'','',0,0,0)"
                            If optFondo(0) Then
                                datFecha1 = "01/" & cmbFondo(0) & "/" & Year(Date)
                                datFecha2 = DateAdd("m", 1, datFecha1)
                                datFecha2 = DateAdd("d", -1, datFecha2)
                            ElseIf optFondo(1) Then
                                datFecha1 = dtpfondo(0)
                                datFecha2 = dtpfondo(1)
                            
                            End If
                            'Call clear_Crystal(FrmAdmin.rptReporte)
                            Set rpReporte = New ctlReport
                            '
                            With rpReporte
                            '
                                .Reporte = gcReport + IIf(chkReport.Value = vbChecked, "edoctaf1.rpt", "edoctaf.rpt")
                                .OrigenDatos(0) = mcDatos
                                .Formulas(0) = "Condominio='" & gcCodInm & "-" & gcNomInm & "'"
                                .Formulas(1) = "Titulo='" & Titulo & "'"
                                .Formulas(2) = "Desde='" & datFecha1 & "'"
                                .Formulas(3) = "Hasta='" & datFecha2 & "'"
                                .Formulas(4) = "Debitos='" & txtFondo(15) & "'"
                                .Formulas(5) = "Creditos='" & txtFondo(14) & "'"
                                .Formulas(6) = "saldoP='" & txtFondo(13) & "'"
                                .Formulas(7) = "xRecaudar=" & CLng(txtFondo(19))
                                .TituloVentana = "Consulta de Movimientos"
                                'errlocal = .PrintReport
                                .Imprimir
'                                If errlocal = 0 Then
'                                    Call rtnBitacora("Impresión Edo.Cta.Fondo " _
'                                    & .Formulas(1) & " Inmueble: " & gcCodInm)
'                                Else
'                                    Call rtnBitacora(.LastErrorString & " Inm:" & gcCodInm)
'                                    MsgBox .LastErrorString, vbCritical, "Error: " & .LastErrorNumber
'                                End If
                            End With
                            '
                            Set rpReporte = Nothing
                            cnnFondo.Close
                            Set cnnFondo = Nothing
                        End With
                    End If
                ElseIf sstFondo.Tab = 2 Then
                    With gridFondo(1)
                        cnnConexion.Execute "DELETE * FROM Mov_Cuotas"
                        For I = 1 To .Rows - 1
                            If I = 1 Then
                                cnnConexion.Execute "INSERT INTO Mov_Cuotas(Fecha,Apto,Monto,Sa" _
                                & "ldo) VALUES ('" & .TextMatrix(I, 1) & "','" & .TextMatrix(I, 2) _
                                & "','" & .TextMatrix(I, 4) & "','" & .TextMatrix(I, 5) & "')"
                                
                            Else
                                .Col = 0
                                .Row = I
                                If I = 2 Then datFecha1 = .TextMatrix(I, 1)
                                img = IIf(.CellPicture = Me.img(0), "+", IIf(.CellPicture = Me.img(1), "-", ""))
                                cnnConexion.Execute "INSERT INTO Mov_Cuotas(Signo,Fecha,Apto,Pe" _
                                & "riodo,Monto,Saldo) VALUES ('" & img & "','" & _
                                .TextMatrix(I, 1) & "','" & .TextMatrix(I, 2) & "','01-" & _
                                .TextMatrix(I, 3) & "','" & .TextMatrix(I, 4) & "','" _
                                & .TextMatrix(I, 5) & "')"
                                datFecha2 = IIf(IsDate(.TextMatrix(I, 1)), .TextMatrix(I, 1), datFecha2)
                            End If
                        Next
                    End With
                    If optFondo(0) Then
                        datFecha1 = "01/" & cmbFondo(0) & "/" & Year(Date)
                        datFecha2 = DateAdd("m", 1, datFecha1)
                        datFecha2 = DateAdd("d", -1, datFecha2)
                    ElseIf optFondo(1) Then
                        datFecha1 = dtpfondo(0)
                        datFecha2 = dtpfondo(1)
                    End If
                    'Call clear_Crystal(FrmAdmin.rptReporte)
                    '
                    Set rpReporte = New ctlReport
                    With rpReporte
                    
                        .Reporte = gcReport & "mov_cuotas.rpt"
                        .OrigenDatos(0) = gcPath & "\sac.mdb"
                        .Formulas(0) = "subtitulo='" & Titulo & "'"
                        .Formulas(1) = "opcion='" & IIf(optavan(0), optavan(0).Caption, optavan(1).Caption) & "'"
                        .Formulas(2) = "Desde='" & datFecha1 & "'"
                        .Formulas(3) = "Hasta='" & datFecha2 & "'"
                        .Formulas(4) = "etiqueta='" & Label1(23).Caption & "'"
                        .Formulas(5) = "cantidad='" & txtFondo(19) & "'"
                        .Formulas(6) = "pagado='" & txtFondo(18) & "'"
                        .Formulas(7) = "Inmueble='" & gcCodInm & " - " & gcNomInm & "'"
                        .TituloVentana = "Consulta de Movimientos Cuotas Especiales"
                        'errlocal = .PrintReport
                        .Imprimir
'                        If errlocal = 0 Then
'                            Call rtnBitacora("Impresión Edo.Cta.Cuotas " _
'                            & .Formulas(1) & " Inmueble: " & gcCodInm)
'                        Else
'                            Call rtnBitacora(.LastErrorString & " Inm:" & gcCodInm)
'                            MsgBox .LastErrorString, vbCritical, "Error: " & .LastErrorNumber
'                        End If
                    End With
                    Set rpReporte = Nothing
                End If
                MousePointer = vbDefault
                '
            Case "DELETE"
                MsgBox "Para eliminar un movimiento seleccionelo de la lista y haga click con " _
                & "el segundo botón del mouse", vbInformation, App.ProductName
            
        End Select
        '
     End With
    '
    End Sub

    
    '---------------------------------------------------------------------------------------------
    '   Devuelve los registros solicitados según parámetros enviados por el user.
    Private Sub EdoCta_Fondo(CG As String)  '
    '---------------------------------------------------------------------------------------------
    'variables locales
    Dim rstMovFondo As New ADODB.Recordset
    Dim rstSA As New ADODB.Recordset  'variables locales
    Dim rstCuotas As New ADODB.Recordset
    Dim rstSAC As New ADODB.Recordset
    Dim curSaldo As Currency
    Dim curDebe As Currency
    Dim curHaber As Currency
    Dim curSA(1) As Currency
    Dim Fecha1, Fecha2
    Dim strSQL As String, strSQLsa As String
    Dim strSQL1 As String, strSQLsa1 As String
    '
    txtFondo(13) = Format(ftnSaldo(CG), "#,##0.00")
    '
    If optFondo(0) Then 'mes seleccionado
        '
        Fecha1 = CDate("01/" & cmbFondo(0) & "/" & Year(Date))
        Fecha2 = DateAdd("m", 1, Fecha1)
        Fecha2 = DateAdd("d", -1, Fecha2)
        Fecha2 = Format(Fecha2, "mm/dd/yy")
        Fecha1 = Format(Fecha1, "mm/dd/yy")
        'selección movimiento del fondo
        strSQL = "SELECT * FROM MovFondo WHERE CodGasto='" & CG & "' And Fecha>=#" & _
        Fecha1 & "# and Fecha<=#" & Fecha2 & "# AND Not Del ORDER BY Fecha,Debe,Concepto;"
        'saldo vienen...
        strSQLsa = "SELECT SUM(Haber) - SUM(Debe) FROM MovFondo WHERE CodGasto='" & _
        CG & "' And Fecha <#" & Fecha1 & "# and Not Del GROUP BY CodGasto;"
        '
        'selección movimiento cuotas especiales
        If optavan(0) Then  'cobrado
            strSQL1 = "SELECT Cdate('01/' & Format(F.FReg,'mm/yyyy')) as Periodo,Sum(DF.Monto) " _
            & "AS Total FROM DetFact as DF INNER JOIN Factura as F ON DF.Fact = F.FACT WHERE F." _
            & "Saldo=0 AND DF.CodGasto='" & CG & "' AND (F.Freg Between #" & Fecha1 & "# and " _
            & "#" & Fecha2 & "#) GROUP BY Cdate('01/' & Format(F.Freg,'mm/yyyy'));"
            '
            strDet(0) = "SELECT F.Freg,F.CodProp,F.Periodo, DF.Monto AS Total FROM DetFact as D" _
            & "F INNER JOIN Factura as F ON DF.Fact = F.FACT WHERE F.Saldo=0 AND DF.CodGasto='" _
            & CG & "' AND (F.Freg Between #" & Fecha1 & "# and #" & Fecha2 & "#);"
            '
            strSQLsa1 = "SELECT Sum(DF.Monto) FROM DetFact as DF INNER JOIN Factura as F ON DF." _
            & "Fact = F.FACT WHERE F.Saldo=0 AND DF.CodGasto='" & CG & "' AND F.freg<#" _
            & Fecha1 & "#;"

        Else    'por cobrar
        
            strSQL1 = "SELECT Cdate('01/' & Format(F.FReg,'mm/yyyy')) as Periodo,Sum(DF.Monto) " _
            & "AS Total FROM DetFact as DF INNER JOIN Factura as F ON DF.Fact = F.FACT WHERE F." _
            & "Saldo>0 AND DF.CodGasto='" & CG & "' AND (F.Freg Between #" & Fecha1 & "# and " _
            & "#" & Fecha2 & "#) GROUP BY Cdate('01/' & Format(F.Freg,'mm/yyyy'));"
            '
            strDet(0) = "SELECT F.Periodo as Freg,F.CodProp,F.Periodo, DF.Monto AS Total FROM DetFact as D" _
            & "F INNER JOIN Factura as F ON DF.Fact = F.FACT WHERE F.Saldo>0 AND DF.CodGasto='" _
            & CG & "' AND (F.Freg Between #" & Fecha1 & "# and #" & Fecha2 & "#);"
            '
            strSQLsa1 = "SELECT Sum(DF.Monto) FROM DetFact as DF INNER JOIN Factura as F ON DF." _
            & "Fact = F.FACT WHERE F.Saldo>0 AND DF.CodGasto='" & CG & "' AND F.freg<#" _
            & Fecha1 & "#;"
            
        End If
        
    ElseIf optFondo(1) Then 'rango de fechas
        '
        Fecha1 = Format(dtpfondo(0), "mm/dd/yy")
        Fecha2 = Format(dtpfondo(1), "mm/dd/yy")
        'movimiento del fondo
        strSQL = "SELECT * FROM MovFondo WHERE CodGasto='" & CG & "' And Fecha>=#" & _
        Fecha1 & "# and Fecha<=#" & Fecha2 & "# AND Not Del ORDER BY Fecha,Debe,Concepto;"
        'monto vienen...
        strSQLsa = "SELECT Sum(HABER) - Sum(Debe) FROM MovFondo WHERE CodGasto='" & CG & _
        "' And Fecha <#" & Fecha1 & "# AND Not Del;"
        '
        'selección movimiento cuotas especiales
        If optavan(0) Then  'cobrado
        
            strSQL1 = "SELECT Cdate('01/' & format(F.Freg,'mm/yyyy')) as Periodo,Sum(DF.Monto) " _
            & "AS Total FROM DetFact as DF INNER JOIN Factura as F ON DF.Fact = F.FACT WHERE F." _
            & "Saldo=0 AND DF.CodGasto='" & CG & "' AND (F.Freg Between #" & Fecha1 & "# and #" _
            & Fecha2 & "#) GROUP BY Cdate('01/' & format(F.Freg,'mm/yyyy'));"
            '
            strDet(0) = "SELECT F.Freg,F.CodProp,F.Periodo, DF.Monto AS Total FROM DetFact as D" _
            & "F INNER JOIN Factura as F ON DF.Fact = F.FACT WHERE F.Saldo=0 AND DF.CodGasto='" _
            & CG & "' AND (F.Freg Between #" & Fecha1 & "# and #" & Fecha2 & "#);"
            '
            strSQLsa1 = "SELECT Sum(DF.Monto) FROM DetFact as DF INNER JOIN Factura as F ON DF." _
            & "Fact = F.FACT WHERE F.Saldo=0 AND DF.CodGasto='" & CG & "' AND F.freg<#" _
            & Fecha1 & "#;"
            
        Else    'por cobrar
            strSQL1 = "SELECT Cdate('01/' & format(F.Freg,'mm/yyyy')) as Periodo,Sum(DF.Monto) " _
            & "AS Total FROM DetFact as DF INNER JOIN Factura as F ON DF.Fact = F.FACT WHERE F." _
            & "Saldo>0 AND DF.CodGasto='" & CG & "' AND (F.Freg Between #" & Fecha1 & "# and #" _
            & Fecha2 & "#) GROUP BY Cdate('01/' & format(F.Freg,'mm/yyyy'));"
            '
            strDet(0) = "SELECT F.Periodo as Freg,F.CodProp,F.Periodo, DF.Monto AS Total FROM DetFact as D" _
            & "F INNER JOIN Factura as F ON DF.Fact = F.FACT WHERE F.Saldo>0 AND DF.CodGasto='" _
            & CG & "' AND (F.Freg Between #" & Fecha1 & "# and #" & Fecha2 & "#);"
            '
            strSQLsa1 = "SELECT Sum(DF.Monto) FROM DetFact as DF INNER JOIN Factura as F ON DF." _
            & "Fact = F.FACT WHERE F.Saldo>0 AND DF.CodGasto='" & CG & "' AND F.freg<#" _
            & Fecha1 & "#;"
            '
        End If
    Else    'todos los movimientos
    
        strSQL = "SELECT * FROM MovFondo WHERE CodGasto='" & CG & "' AND Not Del ORDER BY " _
        & "Fecha,Debe,Concepto;"
        strSQLsa = ""
        strSQLsa1 = ""
        ''selección movimiento cuotas especiales
        If optavan(0) Then  'cobrado
            strSQL1 = "SELECT Cdate('01/' & Format(F.Freg,'mm/yyyy')) as Periodo,Sum(DF.Monto)" _
            & " AS Total FROM DetFact as DF INNER JOIN Factura as F ON DF.Fact = F.FACT WHERE " _
            & "F.Saldo=0 AND DF.CodGasto='" & CG & "' GROUP BY Cdate('01/' & Format(F.Freg,'mm/" _
            & "yyyy'));"
            '
            strDet(0) = "SELECT F.Freg,F.CodProp,F.Periodo, DF.Monto AS Total FROM DetFact as D" _
            & "F INNER JOIN Factura as F ON DF.Fact = F.FACT WHERE F.Saldo=0 AND DF.CodGasto='" _
            & CG & "';"
            '
            strSQLsa1 = "SELECT Monto FROM DetFact WHERE CodGasto ='" & CG & "' AND Codigo ='U" & gcCodInm & "'"
            
        Else    'por cobrar
        
            strSQL1 = "SELECT F.Periodo,Sum(DF.Monto)" _
            & " AS Total FROM DetFact as DF INNER JOIN Factura as F ON DF.Fact = F.FACT WHERE " _
            & "F.Saldo>0 AND DF.CodGasto='" & CG & "' GROUP BY F.Periodo;" ' _
            & ";"
            '
            strDet(0) = "SELECT F.Periodo as Freg,F.CodProp,F.Periodo, DF.Monto AS Total FROM DetFact as D" _
            & "F INNER JOIN Factura as F ON DF.Fact = F.FACT WHERE F.Saldo>0 AND DF.CodGasto='" _
            & CG & "';"
            '
        End If
        '
    End If
    '
    'abre los ADODB.Recordsets
    rstMovFondo.Open strSQL, AdoTfondos.ConnectionString, adOpenStatic, adLockReadOnly, adCmdText
    rstCuotas.Open strSQL1, AdoTfondos.ConnectionString, adOpenKeyset, adLockOptimistic, adCmdText
    '
    If Not strSQLsa = "" Then
        '
        rstSA.Open strSQLsa, AdoTfondos.ConnectionString, adOpenStatic, adLockReadOnly, adCmdText
        curSA(0) = IIf(IsNull(rstSA.Fields(0)), 0, rstSA.Fields(0))
        rstSA.Close
        Set rstSA = Nothing
        '
    Else
        curSA(0) = 0
        'curSA(1) = 0
    End If
    If Not strSQLsa1 = "" Then
        'saldo cuotas especiales
        rstSAC.Open strSQLsa1, AdoTfondos.ConnectionString, adOpenKeyset, adLockOptimistic, adCmdText
        If Not rstSAC.EOF And Not rstSAC.BOF Then
            curSA(1) = IIf(IsNull(rstSAC.Fields(0)), 0, rstSAC.Fields(0))
        End If
        rstSAC.Close
        Set rstSAC = Nothing
    Else
        curSA(1) = 0
    End If
    '
    gridFondo(0).Rows = 2
    Call rtnLimpiar_Grid(gridFondo(0))
    gridFondo(0).TextMatrix(1, 0) = Format(Fecha1, "MM/DD/YY")
    gridFondo(0).TextMatrix(1, 1) = "SA"
    gridFondo(0).TextMatrix(1, 2) = "VIENEN..........."
    gridFondo(0).TextMatrix(1, 3) = Format(0, "#,##0.00 ")
    gridFondo(0).TextMatrix(1, 4) = Format(curSA(0), "#,##0.00 ")
    gridFondo(0).TextMatrix(1, 5) = Format(curSA(0), "#,##0.00 ")
    
    If Not rstMovFondo.EOF And Not rstMovFondo.BOF Then
        gridFondo(0).Rows = rstMovFondo.RecordCount + 2
        gridFondo(0).ColAlignment(5) = flexAlignRightCenter
        'gridfondo.ColAlignment(1)=flexal
        With rstMovFondo
            .MoveFirst: I = 1
            'curHaber = curSA(0)
            '
            Do
                I = I + 1
                gridFondo(0).TextMatrix(I, 0) = Format(!fecha, "dd/mm/yy")
                gridFondo(0).TextMatrix(I, 1) = !Tipo
                gridFondo(0).TextMatrix(I, 2) = !Concepto
                gridFondo(0).TextMatrix(I, 3) = Format(IIf(IsNull(!Debe) Or !Debe = 0, 0, !Debe), "#,##0.00 ")
                gridFondo(0).TextMatrix(I, 4) = Format(IIf(IsNull(!Haber) Or !Haber = 0, 0, !Haber), "#,##0.00 ")
                curDebe = curDebe + IIf(IsNull(!Debe), 0, !Debe)
                curHaber = curHaber + IIf(IsNull(!Haber), 0, !Haber)
                If I = 1 Then
                    curSaldo = curHaber - curDebe
                Else
                    If gridFondo(0).TextMatrix(I, 3) = "" Then
                        curSaldo = gridFondo(0).TextMatrix(I - 1, 5)
                    Else
                        curSaldo = CCur(gridFondo(0).TextMatrix(I - 1, 5)) - _
                        CCur(gridFondo(0).TextMatrix(I, 3)) + CCur(gridFondo(0).TextMatrix(I, 4))
                    End If
                End If
                gridFondo(0).TextMatrix(I, 5) = Format(curSaldo, "#,##0.00 ; (#,##0.00) ")
                .MoveNext
            Loop Until .EOF
            txtFondo(14) = Format(curHaber, "#,##0.00")
            txtFondo(15) = Format(curDebe, "#,##0.00")
            gridFondo(0).Col = 0
            gridFondo(0).Row = 1
            '
        End With
    End If
    rstMovFondo.Close
    Set rstMovFondo = Nothing
    'agrega el conjunto de registros al grid de cuotas especiales
    gridFondo(1).Rows = 2
    Call rtnLimpiar_Grid(gridFondo(1))
    curHaber = 0
    curSaldo = 0
    gridFondo(1).TextMatrix(1, 1) = Format(Fecha1, "dd/mm/yy")
    gridFondo(1).TextMatrix(1, 2) = "Vienen..."
    gridFondo(1).TextMatrix(1, 3) = "Vienen..."
    gridFondo(1).TextMatrix(1, 4) = "Vienen..."
    gridFondo(1).TextMatrix(1, 4) = Format(curSA(1), "#,##0.00")
    gridFondo(1).TextMatrix(1, 5) = Format(curSA(1), "#,##0.00")

    
    If Not rstCuotas.EOF And Not rstCuotas.EOF Then
    
        gridFondo(1).Rows = rstCuotas.RecordCount + 2
        With rstCuotas
            .MoveFirst: I = 1
                        
            curSaldo = curSA(1)
            '
            Do
                I = I + 1
                gridFondo(1).Col = 0
                gridFondo(1).Row = I
                Set gridFondo(1).CellPicture = img(0)
                gridFondo(1).CellPictureAlignment = flexAlignCenterCenter
                gridFondo(1).TextMatrix(I, 1) = .Fields("Periodo")
                gridFondo(1).TextMatrix(I, 2) = "Sub-Total"
                gridFondo(1).TextMatrix(I, 3) = Format(.Fields("Periodo"), "MM/YYYY")
                gridFondo(1).TextMatrix(I, 4) = Format(.Fields("Total"), "#,##0.00")
                'curDebe = curDebe + IIf(IsNull(!Debe), 0, !Debe)
                curHaber = .Fields("Total") + curHaber
                gridFondo(1).TextMatrix(I, 5) = Format(curSaldo + curHaber, "#,##0.00 ; (#,##0.00) ")
                .MoveNext
            Loop Until .EOF
            
            '
        End With
        '
    End If
    curHaber = curSA(1) + curHaber
    txtFondo(19) = Format(curHaber, "#,##0.00")
    txtFondo(18) = Format(curDebe, "#,##0.00")
    txtFondo(20) = Format(curHaber - curDebe, "#,##0.00")
    '
    End Sub

    '---------------------------------------------------------------------------------------------
    '   Funcion:    Saldo
    '   Devuelve un valor moneda que representa el saldo de la cuenta a una
    '   fecha, ambos parámetros determinados por el usuario
    '---------------------------------------------------------------------------------------------
    Private Function ftnSaldo(Gasto As String) As Currency
    'variables locales
    Dim datFecha1, datFecha2    'Variables locales
    Dim rstSaldo As New ADODB.Recordset
    '
    If optFondo(0).Value Then 'MES SEÑALADO
        Label1(15) = "Saldo a " & cmbFondo(0) & ":"
        Label1(24) = "Disponible al mes de " & cmbFondo(0) & ":"
        datFecha1 = "01-" & cmbFondo(0) & "-" & Year(Date)
        datFecha2 = DateAdd("m", 1, datFecha1)
        datFecha2 = DateAdd("d", -1, datFecha2)
        datFecha2 = Format(datFecha2, "mm/dd/yyyy")
        datFecha1 = Format(datFecha1, "mm/dd/yyyy")
    ElseIf optFondo(1).Value Then 'ENTRE RANDO DE FECHAS
        Label1(15) = "Saldo al " & dtpfondo(1) & ":"
        Label1(24) = "Disponible al " & dtpfondo(1) & ":"
        datFecha1 = Format(dtpfondo(0), "mm/dd/yyyy")
        datFecha2 = Format(dtpfondo(1), "mm/dd/yyyy")
    ElseIf optFondo(2).Value Then 'Todos los movimientos
        Label1(15) = "Saldo al " & Date & ":"
        Label1(24) = "Disponible al " & Date & ":"
        datFecha1 = "01/01/1975"
        datFecha2 = Format(Date, "mm/dd/yyyy")
    End If
    '
    rstSaldo.Open "SELECT CodGasto, Sum(Debe) AS D, Sum(Haber) AS H, Sum(Haber)-Sum(Debe) A" _
    & "S S FROM MovFondo WHERE Fecha BETWEEN #" & datFecha1 & "# AND #" & datFecha2 & "# AND Del=False GROUP " _
    & "BY CodGasto HAVING CodGasto='" & Gasto & "';", AdoTfondos.ConnectionString, _
    adOpenStatic, adLockReadOnly
    '
    If Not rstSaldo.EOF Or Not rstSaldo.BOF Then
        ftnSaldo = rstSaldo!s
    Else: ftnSaldo = 0
    End If
    '
    rstSaldo.Close
    Set rstSaldo = Nothing
    '
    End Function

    '---------------------------------------------------------------------------------------------
    '   Rutina:     Eliminar
    '
    '   Elimina un moviento del fondo
    '---------------------------------------------------------------------------------------------
    Public Sub Eliminar()
    'variables locaels
    Dim Z%, cnnT As ADODB.Connection
    Dim rstT As ADODB.Recordset, Cad As String
    Dim rstMF As ADODB.Recordset
    '
    Set cnnT = New ADODB.Connection
    Set rstT = New ADODB.Recordset
    '
    cnnT.Open cnnOLEDB + mcDatos
    '
    With gridFondo(0)
        '
        I = .RowSel
        
        rstT.Open "MovFondo", cnnT, adOpenKeyset, adLockOptimistic, adCmdTable
        '
        rstT.Filter = "CodGasto='" & dtcFondo(0) & "' And Tipo='" & .TextMatrix(I, 1) _
        & "' And dEBE=" & .TextMatrix(I, 3) & " AND hABER=" & .TextMatrix(I, 4) _
        & " AND Fecha='" & .TextMatrix(I, 0) & "' AND Concepto='" & .TextMatrix(I, 2) & "'"
        '
        If Not rstT.EOF And Not rstT.BOF Then
        
            Set rstMF = New ADODB.Recordset
            Set rstMF = cnnT.Execute("SELECT max(Periodo) FROM Factura WHERE Fact Not LIKE 'CH%'")
            '
            If rstMF.Fields(0) >= rstT.Fields("Periodo") Then
            
                MsgBox "Imposible eliminar este registro. Efectue un reintegro y/o ajuste", _
                vbInformation, "Eliminar Reg. " & Titulo
                Call rtnBitacora("Imposible eliminar Mov.Fondo " & dtcFondo(0) & "|" & _
                .TextMatrix(I, 1) & "|" & Left(.TextMatrix(I, 2), 10) & "|" & .TextMatrix(I, 3) _
                & "|" & .TextMatrix(I, 4))
            Else
                '
                Call rtnBitacora("Mov. Fondo Eliminado " & dtcFondo(0) & "|" & .TextMatrix(I, 1) & "|" & _
                Left(.TextMatrix(I, 2), 10) & "|" & .TextMatrix(I, 3) & "|" & .TextMatrix(I, 4))
                'actualzia el saldo actual de la cuenta de fondo
                If rstT("Debe") > 0 Then
                    cnnT.Execute "UPDATE TGastos SET SaldoActual=SaldoActual + '" & rstT("Debe") _
                    & "' WHERE CodGasto='" & rstT("CodGasto") & "'"
                Else
                    cnnT.Execute "UPDATE TGastos SET SaldoActual=SaldoActual - '" & rstT("Haber") _
                    & "' WHERE CodGasto='" & rstT("CodGasto") & "'"
                End If
                rstT.Update "Del", True
                '
                rstT.Requery
                Call EdoCta_Fondo(dtcFondo(0))
                MsgBox "Registro Eliminado", vbInformation, App.ProductName
                '
            End If
            '
            rstMF.Close
            Set rstMF = Nothing
            '
        End If
        '
        rstT.Close
        cnnT.Close
        Set rstT = Nothing
        Set cnnT = Nothing
        '
    End With
    '
    End Sub

    
    '------------------------------------------------------------------------------------
    '   Función:    Guardar
    '
    '   Si guarda el movimiento con éxito devuelve True de lo contrario
    '   retorna false. Además acutaliza el saldo de la cuenta
    '------------------------------------------------------------------------------------
    Public Function agregar() As Boolean
    On Error Resume Next
    'variables locales
    Dim cnnT As ADODB.Connection
    Dim rstUltMesFac As ADODB.Recordset
    Dim curDebe@, curHaber@, strBitacora$
    Dim Cargado_a As String
    '
    'valida la cantidad de datos necesarios mínimos para procesar la transacción
    If Not Validar_Guardar Then
        '
        'crea una instancia de los objetos
        Set cnnT = New ADODB.Connection
        Set rstUltMesFac = New ADODB.Recordset
        'abre la conexíon al orìgen de datos
        cnnT.Open cnnOLEDB + mcDatos
        'busca el último período facturado
        rstUltMesFac.Open "SELECT MAX(Periodo) as UMF FROM Factura WHERE FACT Not Like 'CH%'", _
        cnnT, adOpenKeyset, adLockOptimistic, adCmdText
        '
        If Not rstUltMesFac.EOF And Not rstUltMesFac.BOF And Not IsNull(rstUltMesFac("UMF")) Then
            Cargado_a = DateAdd("M", 1, rstUltMesFac("UMF"))
        Else
            Cargado_a = "01/" & Format(msk, "mm/yy")
        End If
        'cierra y descarga el objeto
        rstUltMesFac.Close
        Set rstUltMesFac = Nothing
        '
        cnnT.BeginTrans 'comienza una transacción
        If cmb.ListIndex <= 1 Then
            curHaber = txtFondo(17)
            curDebe = 0
        Else
            curDebe = txtFondo(17)
            curHaber = 0
        End If
        '
        cnnT.Execute "INSERT INTO MovFondo (CodGasto,Fecha,Tipo,Periodo,Concepto,Debe,Haber) VA" _
        & "LUES ('" & dtcFondo(0) & "','" & msk & "','" & cmb & "','" & Cargado_a & "','" & _
        txtFondo(16) & "','" & curDebe & "','" & curHaber & "')"
        '
        'actualzia el saldo actual de la cuenta de fondo
        cnnT.Execute "UPDATE TGastos SET SaldoActual=SaldoActual + '" & IIf(curDebe > 0, _
        curDebe * -1, curHaber) & "' WHERE CodGasto='" & dtcFondo(0) & "'"
        '
        If Err.Number <> 0 Then
            'si no ocurre ningun error
            MsgBox "Transacción cancelada..." & Err.Description, vbExclamation, App.ProductName
            cnnT.RollbackTrans  'procesa la toda la operación
            strBitacora = "Transacción cancelada por error " & Err.Number
        Else
            agregar = MsgBox("Transacción llevada a cabo con éxito", vbInformation, _
            App.ProductName)
            cnnT.CommitTrans    'reversa todas la operaciones
            strBitacora = "Movimieno guardado con éxito"
            AdoTfondos.Refresh
        End If
        Call rtnBitacora(strBitacora)
        cnnT.Close
        Set cnnT = Nothing
        '
    End If
    '
    End Function

    '------------------------------------------------------------------------------------
    '   Funcion:    validar_guardar
    '
    '   Si no tiene el mínimo de datos necesarios para procesar el movimiento
    '   envia un mensaje al usuario y devuelve un valor True
    '------------------------------------------------------------------------------------
    Private Function Validar_Guardar() As Boolean
    'Vairables locales
    Dim strTitulo$
    '
    strTitulo = App.ProductName
    If Not IsDate(msk) Then
        Validar_Guardar = MsgBox("Fecha No Válida..", vbExclamation, strTitulo)
    End If
    If cmb = "" Then
        Validar_Guardar = MsgBox("Falta Tipo de Movimiento..", vbExclamation, strTitulo)
    End If
    If txtFondo(16) = "" Then
        Validar_Guardar = MsgBox("Falta Descripción del Movimiento..", vbExclamation, strTitulo)
    End If
    If txtFondo(17) = "" Then
        Validar_Guardar = MsgBox("Falta Monto del Movimiento..", vbExclamation, strTitulo)
    End If
    '
    End Function


    Private Sub txtFondo_KeyPress(Index%, KeyAscii%)
    '
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If Index = 16 Then
        If KeyAscii = 13 Then txtFondo(17).SetFocus
    ElseIf Index = 17 Then
        If KeyAscii = 46 Then KeyAscii = 44
        Call Validacion(KeyAscii, "0123456798,")
        If KeyAscii = 13 Then txtFondo(17) = Format(txtFondo(17), "#,##0.00")
    End If
    '
    End Sub

