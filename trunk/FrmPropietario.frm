VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form FrmPropietario 
   AutoRedraw      =   -1  'True
   Caption         =   "Ficha de Propietarios"
   ClientHeight    =   6555
   ClientLeft      =   -615
   ClientTop       =   2190
   ClientWidth     =   8880
   ControlBox      =   0   'False
   Icon            =   "FrmPropietario.frx":0000
   LinkTopic       =   "Form3"
   MDIChild        =   -1  'True
   ScaleHeight     =   6555
   ScaleWidth      =   8880
   Tag             =   "1"
   WindowState     =   2  'Maximized
   Begin ComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   480
      Left            =   0
      TabIndex        =   0
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
            Key             =   "SAVE"
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
   Begin VB.Frame fraPropietario 
      Enabled         =   0   'False
      Height          =   2625
      Index           =   0
      Left            =   285
      TabIndex        =   1
      Top             =   1005
      Width           =   9045
      Begin VB.TextBox TxtPropietario 
         DataField       =   "Codigo"
         DataSource      =   "AdoProp"
         Height          =   315
         Index           =   0
         Left            =   1680
         TabIndex        =   54
         Top             =   420
         Width           =   960
      End
      Begin VB.TextBox TxtPropietario 
         DataField       =   "Nombre"
         DataSource      =   "AdoProp"
         Height          =   315
         Index           =   1
         Left            =   4470
         TabIndex        =   53
         Top             =   420
         Width           =   4440
      End
      Begin VB.TextBox TxtPropietario 
         DataField       =   "Contacto"
         DataSource      =   "AdoProp"
         Height          =   315
         Index           =   2
         Left            =   4455
         TabIndex        =   3
         Top             =   975
         Width           =   4440
      End
      Begin VB.TextBox TxtPropietario 
         DataField       =   "Cedula"
         DataSource      =   "AdoProp"
         Height          =   315
         Index           =   3
         Left            =   5730
         TabIndex        =   2
         Top             =   1530
         Width           =   3180
      End
      Begin MSDataListLib.DataCombo CmbCarJunta 
         DataField       =   "CarJunta"
         DataSource      =   "AdoProp"
         Height          =   315
         Left            =   5730
         TabIndex        =   4
         Top             =   2085
         Width           =   3180
         _ExtentX        =   5609
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   -2147483643
         ListField       =   "Nombre"
         BoundColumn     =   "Nombre"
         Text            =   ""
      End
      Begin MSMask.MaskEdBox MskFecha 
         DataField       =   "FecReg"
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "dd/MM/yyyy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   8202
            SubFormatType   =   3
         EndProperty
         DataSource      =   "AdoProp"
         Height          =   315
         Left            =   1680
         TabIndex        =   5
         Top             =   2085
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   556
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   10
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSDataListLib.DataCombo CmbTipoCon 
         DataField       =   "TipoCon"
         DataSource      =   "AdoProp"
         Height          =   315
         Left            =   1680
         TabIndex        =   6
         Top             =   1530
         Width           =   2100
         _ExtentX        =   3704
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   -2147483643
         ListField       =   "Nombre"
         BoundColumn     =   "Nombre"
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo CmbCivil 
         DataField       =   "Civil"
         DataSource      =   "AdoProp"
         Height          =   315
         Left            =   1680
         TabIndex        =   7
         Top             =   975
         Width           =   1620
         _ExtentX        =   2858
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   -2147483643
         ListField       =   "Nombre"
         BoundColumn     =   "Nombre"
         Text            =   ""
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Label1"
         Height          =   255
         Index           =   7
         Left            =   3105
         TabIndex        =   62
         Top             =   2115
         Width           =   2415
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Label1"
         Height          =   255
         Index           =   6
         Left            =   3855
         TabIndex        =   61
         Top             =   1560
         Width           =   1665
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Label1"
         Height          =   255
         Index           =   5
         Left            =   3300
         TabIndex        =   60
         Top             =   1005
         Width           =   1095
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Label1"
         Height          =   255
         Index           =   4
         Left            =   3300
         TabIndex        =   59
         Top             =   450
         Width           =   1095
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Label1"
         Height          =   255
         Index           =   3
         Left            =   150
         TabIndex        =   58
         Top             =   2115
         Width           =   1305
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Label1"
         Height          =   255
         Index           =   2
         Left            =   150
         TabIndex        =   57
         Top             =   1560
         Width           =   1305
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Label1"
         Height          =   255
         Index           =   1
         Left            =   150
         TabIndex        =   56
         Top             =   1005
         Width           =   1305
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Label1"
         Height          =   250
         Index           =   0
         Left            =   150
         TabIndex        =   55
         Top             =   452
         Width           =   1300
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   7440
      Left            =   0
      TabIndex        =   8
      Top             =   570
      Width           =   9630
      _ExtentX        =   16986
      _ExtentY        =   13123
      _Version        =   393216
      Tabs            =   5
      TabsPerRow      =   5
      TabHeight       =   520
      ShowFocusRect   =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Datos Generales"
      TabPicture(0)   =   "FrmPropietario.frx":000C
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "fraPropietario(1)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Datos Administrativos"
      TabPicture(1)   =   "FrmPropietario.frx":0028
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "fraPropietario(2)"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "SacOnLine"
      TabPicture(2)   =   "FrmPropietario.frx":0044
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "AdoProp"
      Tab(2).Control(1)=   "fraPropietario(3)"
      Tab(2).ControlCount=   2
      TabCaption(3)   =   "Lista"
      TabPicture(3)   =   "FrmPropietario.frx":0060
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "FrmBusca1"
      Tab(3).Control(1)=   "FrmBusca"
      Tab(3).Control(2)=   "DataGrid1"
      Tab(3).ControlCount=   3
      TabCaption(4)   =   "Cargar Deuda"
      TabPicture(4)   =   "FrmPropietario.frx":007C
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "fraPropietario(4)"
      Tab(4).ControlCount=   1
      Begin VB.Frame fraPropietario 
         Height          =   3825
         Index           =   4
         Left            =   -74730
         TabIndex        =   101
         Top             =   3330
         Width           =   9075
         Begin VB.CommandButton cmd 
            Height          =   420
            Left            =   5775
            Picture         =   "FrmPropietario.frx":0098
            Style           =   1  'Graphical
            TabIndex        =   107
            Top             =   1155
            Width           =   540
         End
         Begin VB.TextBox TxtPropietario 
            Alignment       =   1  'Right Justify
            Height          =   315
            Index           =   24
            Left            =   4425
            TabIndex        =   102
            Text            =   "0,00"
            Top             =   1200
            Width           =   1200
         End
         Begin MSMask.MaskEdBox mskPer 
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "dd/MM/yyyy"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   8202
               SubFormatType   =   3
            EndProperty
            Height          =   315
            Left            =   1620
            TabIndex        =   103
            Top             =   1215
            Width           =   1155
            _ExtentX        =   2037
            _ExtentY        =   556
            _Version        =   393216
            MaxLength       =   10
            Format          =   "dd/mm/yyyy"
            Mask            =   "01/##/####"
            PromptChar      =   "_"
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Facturado:"
            Height          =   255
            Index           =   44
            Left            =   3210
            TabIndex        =   106
            Top             =   1245
            Width           =   1095
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Período:"
            Height          =   255
            Index           =   43
            Left            =   570
            TabIndex        =   105
            Top             =   1245
            Width           =   885
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Utilice esta seccón para cargar la deuda al ingresar un nuevo propitario:"
            Height          =   255
            Index           =   42
            Left            =   435
            TabIndex        =   104
            Top             =   435
            Width           =   8295
         End
      End
      Begin MSDataGridLib.DataGrid DataGrid1 
         Bindings        =   "FrmPropietario.frx":04DA
         Height          =   4830
         Left            =   -74910
         TabIndex        =   52
         Top             =   885
         Width           =   9405
         _ExtentX        =   16589
         _ExtentY        =   8520
         _Version        =   393216
         AllowUpdate     =   0   'False
         BackColor       =   16777215
         BorderStyle     =   0
         HeadLines       =   1
         RowHeight       =   19
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
            Name            =   "Comic Sans MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   4
         BeginProperty Column00 
            DataField       =   "Codigo"
            Caption         =   "Código"
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
         BeginProperty Column01 
            DataField       =   "Nombre"
            Caption         =   "Propietario"
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
            DataField       =   "Contacto"
            Caption         =   "Contacto"
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
         BeginProperty Column03 
            DataField       =   "TelfHab"
            Caption         =   "Teléfono Hab."
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   "(0###)-###-##-##"
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
               ColumnWidth     =   629,858
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   3795,024
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   2805,166
            EndProperty
            BeginProperty Column03 
            EndProperty
         EndProperty
      End
      Begin VB.Frame fraPropietario 
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
         Height          =   3930
         Index           =   2
         Left            =   -74715
         TabIndex        =   51
         Top             =   3435
         Width           =   9045
         Begin VB.TextBox TxtPropietario 
            Alignment       =   1  'Right Justify
            DataField       =   "Deuda"
            BeginProperty DataFormat 
               Type            =   1
               Format          =   " #,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   8202
               SubFormatType   =   0
            EndProperty
            DataSource      =   "AdoProp"
            Height          =   285
            Index           =   23
            Left            =   7665
            Locked          =   -1  'True
            TabIndex        =   99
            Top             =   2475
            Width           =   1215
         End
         Begin VB.TextBox TxtPropietario 
            Alignment       =   1  'Right Justify
            DataField       =   "Recibos"
            BeginProperty DataFormat 
               Type            =   1
               Format          =   " #,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   8202
               SubFormatType   =   0
            EndProperty
            DataSource      =   "AdoProp"
            Height          =   285
            Index           =   22
            Left            =   2085
            Locked          =   -1  'True
            TabIndex        =   97
            Top             =   2475
            Width           =   1215
         End
         Begin VB.CheckBox Check1 
            Caption         =   "No enviar Aviso de Cobro al Edificio"
            DataField       =   "AC"
            DataSource      =   "AdoProp"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   435
            Left            =   3960
            TabIndex        =   95
            Top             =   1095
            Width           =   2115
         End
         Begin VB.CheckBox chkConvenio 
            Caption         =   "Convenio de Pago"
            DataField       =   "Convenio"
            DataSource      =   "AdoProp"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   285
            Left            =   3945
            TabIndex        =   93
            Top             =   675
            Width           =   2115
         End
         Begin VB.CheckBox chkDemanda 
            Caption         =   "Demandado"
            DataField       =   "Demanda"
            DataSource      =   "AdoProp"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   285
            Left            =   3945
            TabIndex        =   92
            Top             =   270
            Width           =   1800
         End
         Begin VB.TextBox TxtPropietario 
            Alignment       =   1  'Right Justify
            DataField       =   "GtoAdm"
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "#,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   8202
               SubFormatType   =   1
            EndProperty
            DataSource      =   "AdoProp"
            Height          =   285
            Index           =   14
            Left            =   7680
            Locked          =   -1  'True
            TabIndex        =   19
            Top             =   270
            Width           =   1215
         End
         Begin VB.TextBox TxtPropietario 
            Alignment       =   1  'Right Justify
            DataField       =   "Alicuota"
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "#,##0.000000"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   8202
               SubFormatType   =   1
            EndProperty
            DataSource      =   "AdoProp"
            Height          =   285
            Index           =   9
            Left            =   2070
            TabIndex        =   13
            Top             =   270
            Width           =   1215
         End
         Begin VB.TextBox TxtPropietario 
            Alignment       =   1  'Right Justify
            DataField       =   "Deuda"
            BeginProperty DataFormat 
               Type            =   1
               Format          =   " #,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   8202
               SubFormatType   =   0
            EndProperty
            DataSource      =   "AdoProp"
            Height          =   285
            Index           =   18
            Left            =   7665
            Locked          =   -1  'True
            TabIndex        =   23
            Top             =   2010
            Width           =   1215
         End
         Begin VB.TextBox TxtPropietario 
            Alignment       =   1  'Right Justify
            DataField       =   "FrePago"
            BeginProperty DataFormat 
               Type            =   0
               Format          =   "#.##0,00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   0
            EndProperty
            DataSource      =   "AdoProp"
            Height          =   285
            Index           =   12
            Left            =   2070
            TabIndex        =   16
            Top             =   1560
            Width           =   1215
         End
         Begin VB.TextBox TxtPropietario 
            Alignment       =   1  'Right Justify
            DataField       =   "Horario"
            BeginProperty DataFormat 
               Type            =   0
               Format          =   "#.##0,00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   0
            EndProperty
            DataSource      =   "AdoProp"
            Height          =   285
            Index           =   11
            Left            =   2070
            TabIndex        =   15
            Top             =   1140
            Width           =   1215
         End
         Begin VB.TextBox TxtPropietario 
            DataField       =   "Notas"
            DataSource      =   "AdoProp"
            Height          =   855
            Index           =   13
            Left            =   1350
            MultiLine       =   -1  'True
            TabIndex        =   18
            Top             =   2910
            Width           =   7530
         End
         Begin VB.TextBox TxtPropietario 
            Alignment       =   1  'Right Justify
            DataField       =   "UltPago"
            BeginProperty DataFormat 
               Type            =   1
               Format          =   " #,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   8202
               SubFormatType   =   0
            EndProperty
            DataSource      =   "AdoProp"
            Height          =   285
            Index           =   17
            Left            =   7665
            Locked          =   -1  'True
            TabIndex        =   22
            Top             =   1560
            Width           =   1215
         End
         Begin VB.TextBox TxtPropietario 
            Alignment       =   1  'Right Justify
            DataField       =   "FecUltPag"
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "dd/MM/yyyy"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   8202
               SubFormatType   =   3
            EndProperty
            DataSource      =   "AdoProp"
            Height          =   285
            Index           =   16
            Left            =   7665
            Locked          =   -1  'True
            TabIndex        =   21
            Top             =   1140
            Width           =   1215
         End
         Begin VB.TextBox TxtPropietario 
            Alignment       =   1  'Right Justify
            DataField       =   "Promesa"
            BeginProperty DataFormat 
               Type            =   0
               Format          =   "#,###.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   8202
               SubFormatType   =   0
            EndProperty
            DataSource      =   "AdoProp"
            Height          =   285
            Index           =   15
            Left            =   7680
            Locked          =   -1  'True
            TabIndex        =   20
            Top             =   705
            Width           =   1215
         End
         Begin VB.TextBox TxtPropietario 
            Alignment       =   1  'Right Justify
            DataField       =   "DiaCobro"
            BeginProperty DataFormat 
               Type            =   0
               Format          =   "#.##0,00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   0
            EndProperty
            DataSource      =   "AdoProp"
            Height          =   285
            Index           =   10
            Left            =   2070
            TabIndex        =   14
            Top             =   705
            Width           =   1215
         End
         Begin MSDataListLib.DataCombo CmbCobrador 
            DataField       =   "FormaCob"
            DataSource      =   "AdoProp"
            Height          =   315
            Left            =   2070
            TabIndex        =   17
            Top             =   1995
            Width           =   2025
            _ExtentX        =   3572
            _ExtentY        =   556
            _Version        =   393216
            BackColor       =   -2147483643
            ListField       =   "Cobrador"
            BoundColumn     =   "Nombre"
            Text            =   ""
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Deuda:"
            Height          =   255
            Index           =   41
            Left            =   5745
            TabIndex        =   100
            Top             =   2490
            Width           =   1500
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Nº Rec. Pendientes:"
            Height          =   255
            Index           =   38
            Left            =   165
            TabIndex        =   98
            Top             =   2460
            Width           =   1500
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Label1"
            Height          =   255
            Index           =   30
            Left            =   6015
            TabIndex        =   85
            Top             =   720
            Width           =   1500
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Label1"
            Height          =   255
            Index           =   29
            Left            =   6015
            TabIndex        =   84
            Top             =   1155
            Width           =   1500
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Label1"
            Height          =   255
            Index           =   28
            Left            =   6015
            TabIndex        =   83
            Top             =   1575
            Width           =   1500
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Label1"
            Height          =   255
            Index           =   27
            Left            =   6015
            TabIndex        =   82
            Top             =   2025
            Width           =   1500
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Label1"
            Height          =   255
            Index           =   26
            Left            =   6015
            TabIndex        =   81
            Top             =   285
            Width           =   1500
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Label1"
            Height          =   255
            Index           =   25
            Left            =   420
            TabIndex        =   80
            Top             =   2895
            Width           =   900
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Label1"
            Height          =   255
            Index           =   24
            Left            =   165
            TabIndex        =   79
            Top             =   2025
            Width           =   1500
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Label1"
            Height          =   255
            Index           =   23
            Left            =   165
            TabIndex        =   78
            Top             =   1575
            Width           =   1500
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Label1"
            Height          =   255
            Index           =   22
            Left            =   165
            TabIndex        =   77
            Top             =   1155
            Width           =   1500
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Label1"
            Height          =   255
            Index           =   21
            Left            =   165
            TabIndex        =   76
            Top             =   720
            Width           =   1500
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Label1"
            Height          =   255
            Index           =   20
            Left            =   165
            TabIndex        =   75
            Top             =   285
            Width           =   1500
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
         Height          =   1320
         Left            =   -74640
         TabIndex        =   37
         Top             =   6030
         Width           =   3210
         Begin VB.OptionButton OptBusca 
            Caption         =   "Cargo"
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
            Left            =   1740
            TabIndex        =   41
            Tag             =   "CarJunta"
            Top             =   825
            Width           =   1320
         End
         Begin VB.OptionButton OptBusca 
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
            Left            =   1740
            TabIndex        =   39
            Tag             =   "Nombre"
            Top             =   360
            Width           =   1335
         End
         Begin VB.OptionButton OptBusca 
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
            Left            =   135
            TabIndex        =   40
            Tag             =   "Contacto"
            Top             =   810
            Width           =   1455
         End
         Begin VB.OptionButton OptBusca 
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
            Tag             =   "Codigo"
            Top             =   390
            Value           =   -1  'True
            Width           =   1335
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
         Height          =   1335
         Left            =   -71175
         TabIndex        =   42
         Top             =   6015
         Width           =   5520
         Begin VB.TextBox TxtPropietario 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   28
            Left            =   1800
            Locked          =   -1  'True
            TabIndex        =   46
            Top             =   810
            Width           =   690
         End
         Begin VB.TextBox TxtPropietario 
            BackColor       =   &H00FFFFFF&
            Height          =   315
            Index           =   27
            Left            =   1020
            TabIndex        =   44
            Top             =   285
            Width           =   4335
         End
         Begin VB.CommandButton BotBusca 
            Height          =   330
            Left            =   4935
            Picture         =   "FrmPropietario.frx":04F0
            Style           =   1  'Graphical
            TabIndex        =   47
            ToolTipText     =   "Buscar"
            Top             =   750
            Width           =   375
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Label1"
            Height          =   255
            Index           =   40
            Left            =   60
            TabIndex        =   45
            Top             =   825
            Width           =   1680
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Label1"
            Height          =   255
            Index           =   39
            Left            =   75
            TabIndex        =   43
            Top             =   330
            Width           =   900
         End
      End
      Begin VB.Frame fraPropietario 
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
         Height          =   3960
         Index           =   1
         Left            =   270
         TabIndex        =   10
         Top             =   3435
         Width           =   9075
         Begin VB.CheckBox Check2 
            Caption         =   "No enviar Aviso de Cobro al Edificio"
            DataField       =   "AC"
            DataSource      =   "AdoProp"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   525
            Left            =   1425
            TabIndex        =   96
            Top             =   3120
            Width           =   3810
         End
         Begin VB.TextBox TxtPropietario 
            DataField       =   "Direccion"
            DataSource      =   "AdoProp"
            ForeColor       =   &H00404040&
            Height          =   945
            Index           =   5
            Left            =   1050
            MultiLine       =   -1  'True
            TabIndex        =   89
            Top             =   1230
            Width           =   3495
         End
         Begin VB.TextBox TxtPropietario 
            DataField       =   "Empresa"
            DataSource      =   "AdoProp"
            Height          =   315
            Index           =   4
            Left            =   1035
            TabIndex        =   87
            Top             =   360
            Width           =   3500
         End
         Begin VB.TextBox TxtPropietario 
            DataField       =   "Postal"
            DataSource      =   "AdoProp"
            Height          =   315
            Index           =   6
            Left            =   5790
            TabIndex        =   24
            Top             =   360
            Width           =   1200
         End
         Begin VB.TextBox TxtPropietario 
            DataField       =   "Email"
            DataSource      =   "AdoProp"
            Height          =   315
            Index           =   8
            Left            =   5775
            MaxLength       =   120
            TabIndex        =   12
            Top             =   2685
            Width           =   2800
         End
         Begin VB.TextBox TxtPropietario 
            DataField       =   "ExtOfc"
            DataSource      =   "AdoProp"
            Height          =   315
            Index           =   7
            Left            =   7860
            TabIndex        =   11
            Top             =   1731
            Width           =   1080
         End
         Begin MSDataListLib.DataCombo CmbCiudad 
            DataField       =   "Ciudad"
            DataSource      =   "AdoProp"
            Height          =   315
            Left            =   5775
            TabIndex        =   35
            Top             =   795
            Width           =   2805
            _ExtentX        =   4948
            _ExtentY        =   556
            _Version        =   393216
            BackColor       =   -2147483643
            ListField       =   "Nombre"
            BoundColumn     =   "Nombre"
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo CmbEstado 
            DataField       =   "Estado"
            DataSource      =   "AdoProp"
            Height          =   315
            Left            =   5775
            TabIndex        =   48
            Top             =   1230
            Width           =   2805
            _ExtentX        =   4948
            _ExtentY        =   556
            _Version        =   393216
            BackColor       =   -2147483643
            ListField       =   "Nombre"
            BoundColumn     =   "Nombre"
            Text            =   ""
         End
         Begin MSMask.MaskEdBox MskTelefono 
            DataField       =   "Telefonos"
            BeginProperty DataFormat 
               Type            =   0
               Format          =   "(0000)-###-##-##"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   8202
               SubFormatType   =   0
            EndProperty
            DataSource      =   "AdoProp"
            Height          =   315
            Index           =   2
            Left            =   5775
            TabIndex        =   49
            Top             =   1731
            Width           =   1605
            _ExtentX        =   2831
            _ExtentY        =   556
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   16
            Format          =   "(####)-###-##-##"
            Mask            =   "(####)-###-##-##"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MskTelefono 
            DataField       =   "Celular"
            BeginProperty DataFormat 
               Type            =   0
               Format          =   "(0000)-###-##-##"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   8202
               SubFormatType   =   0
            EndProperty
            DataSource      =   "AdoProp"
            Height          =   315
            Index           =   3
            Left            =   5775
            TabIndex        =   50
            Top             =   2208
            Width           =   1605
            _ExtentX        =   2831
            _ExtentY        =   556
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   16
            Format          =   "(####)-###-##-##"
            Mask            =   "(####)-###-##-##"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MskTelefono 
            DataField       =   "Fax"
            BeginProperty DataFormat 
               Type            =   0
               Format          =   "(0000)-###-##-##"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   8202
               SubFormatType   =   0
            EndProperty
            DataSource      =   "AdoProp"
            Height          =   315
            Index           =   1
            Left            =   1425
            TabIndex        =   88
            Top             =   2730
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   556
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   16
            Format          =   "(####)-###-##-##"
            Mask            =   "(####)-###-##-##"
            PromptChar      =   "_"
         End
         Begin MSDataListLib.DataCombo CmbCargo 
            DataField       =   "Cargo"
            DataSource      =   "AdoProp"
            Height          =   315
            Left            =   1065
            TabIndex        =   90
            Top             =   795
            Width           =   3495
            _ExtentX        =   6165
            _ExtentY        =   556
            _Version        =   393216
            BackColor       =   -2147483643
            ListField       =   "Nombre"
            BoundColumn     =   "Nombre"
            Text            =   ""
         End
         Begin MSMask.MaskEdBox MskTelefono 
            DataField       =   "TelfHab"
            BeginProperty DataFormat 
               Type            =   0
               Format          =   "(0000)-###-##-##"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   8202
               SubFormatType   =   0
            EndProperty
            DataSource      =   "AdoProp"
            Height          =   315
            Index           =   0
            Left            =   1425
            TabIndex        =   91
            Top             =   2295
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   556
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   16
            Format          =   "(0###)-###-##-##"
            Mask            =   "(####)-###-##-##"
            PromptChar      =   "_"
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Label1"
            Height          =   255
            Index           =   19
            Left            =   4740
            TabIndex        =   74
            Top             =   2700
            Width           =   900
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Label1"
            Height          =   255
            Index           =   18
            Left            =   4590
            TabIndex        =   73
            Top             =   2220
            Width           =   1050
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Label1"
            Height          =   255
            Index           =   17
            Left            =   60
            TabIndex        =   72
            Top             =   2760
            Width           =   1230
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "N° Tel.Ofc."
            Height          =   255
            Index           =   16
            Left            =   4740
            TabIndex        =   71
            Top             =   1746
            Width           =   900
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Label1"
            Height          =   255
            Index           =   15
            Left            =   4740
            TabIndex        =   70
            Top             =   1260
            Width           =   900
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Label1"
            Height          =   255
            Index           =   14
            Left            =   4740
            TabIndex        =   69
            Top             =   825
            Width           =   900
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Label1"
            Height          =   255
            Index           =   13
            Left            =   4740
            TabIndex        =   68
            Top             =   390
            Width           =   900
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Ext.:"
            Height          =   195
            Index           =   12
            Left            =   7500
            TabIndex        =   67
            Top             =   1776
            Width           =   315
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Label1"
            Height          =   255
            Index           =   11
            Left            =   60
            TabIndex        =   66
            Top             =   2325
            Width           =   1230
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Label1"
            Height          =   255
            Index           =   10
            Left            =   105
            TabIndex        =   65
            Top             =   1200
            Width           =   900
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Label1"
            Height          =   255
            Index           =   9
            Left            =   105
            TabIndex        =   64
            Top             =   825
            Width           =   900
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Label1"
            Height          =   255
            Index           =   8
            Left            =   105
            TabIndex        =   63
            Top             =   390
            Width           =   900
         End
      End
      Begin VB.Frame fraPropietario 
         Height          =   3420
         Index           =   3
         Left            =   -74700
         TabIndex        =   9
         Top             =   3435
         Width           =   9015
         Begin ComctlLib.Toolbar bHerramienta 
            Height          =   450
            Left            =   0
            TabIndex        =   94
            Top             =   0
            Width           =   9000
            _ExtentX        =   15875
            _ExtentY        =   794
            ButtonWidth     =   714
            ButtonHeight    =   688
            AllowCustomize  =   0   'False
            Appearance      =   1
            ImageList       =   "ImageList1"
            _Version        =   327682
            BeginProperty Buttons {0713E452-850A-101B-AFC0-4210102A8DA7} 
               NumButtons      =   12
               BeginProperty Button1 {0713F354-850A-101B-AFC0-4210102A8DA7} 
                  Enabled         =   0   'False
                  Key             =   "FIRST"
                  Object.ToolTipText     =   "Primer Registro"
                  Object.Tag             =   ""
                  ImageIndex      =   1
               EndProperty
               BeginProperty Button2 {0713F354-850A-101B-AFC0-4210102A8DA7} 
                  Enabled         =   0   'False
                  Key             =   "BACK"
                  Object.ToolTipText     =   "Registro Anterior"
                  Object.Tag             =   ""
                  ImageIndex      =   2
               EndProperty
               BeginProperty Button3 {0713F354-850A-101B-AFC0-4210102A8DA7} 
                  Enabled         =   0   'False
                  Key             =   "NEXT"
                  Object.ToolTipText     =   "Siguiente Registro"
                  Object.Tag             =   ""
                  ImageIndex      =   3
               EndProperty
               BeginProperty Button4 {0713F354-850A-101B-AFC0-4210102A8DA7} 
                  Enabled         =   0   'False
                  Key             =   "LAST"
                  Object.ToolTipText     =   "Último Registro"
                  Object.Tag             =   ""
                  ImageIndex      =   4
               EndProperty
               BeginProperty Button5 {0713F354-850A-101B-AFC0-4210102A8DA7} 
                  Key             =   "NEW"
                  Object.ToolTipText     =   "Nuevo Registro"
                  Object.Tag             =   ""
                  ImageIndex      =   5
               EndProperty
               BeginProperty Button6 {0713F354-850A-101B-AFC0-4210102A8DA7} 
                  Enabled         =   0   'False
                  Key             =   "SAVE"
                  Object.ToolTipText     =   "Guardar Registro"
                  Object.Tag             =   ""
                  ImageIndex      =   6
               EndProperty
               BeginProperty Button7 {0713F354-850A-101B-AFC0-4210102A8DA7} 
                  Enabled         =   0   'False
                  Object.Visible         =   0   'False
                  Key             =   "FIND"
                  Object.ToolTipText     =   "Buscar Registro"
                  Object.Tag             =   ""
                  ImageIndex      =   7
               EndProperty
               BeginProperty Button8 {0713F354-850A-101B-AFC0-4210102A8DA7} 
                  Enabled         =   0   'False
                  Key             =   "UNDO"
                  Object.ToolTipText     =   "Cancelar Registro"
                  Object.Tag             =   ""
                  ImageIndex      =   8
               EndProperty
               BeginProperty Button9 {0713F354-850A-101B-AFC0-4210102A8DA7} 
                  Enabled         =   0   'False
                  Key             =   "DELETE"
                  Object.ToolTipText     =   "Eliminar Registro"
                  Object.Tag             =   ""
                  ImageIndex      =   9
               EndProperty
               BeginProperty Button10 {0713F354-850A-101B-AFC0-4210102A8DA7} 
                  Enabled         =   0   'False
                  Key             =   "EDIT"
                  Object.ToolTipText     =   "Editar Registro"
                  Object.Tag             =   ""
                  ImageIndex      =   10
               EndProperty
               BeginProperty Button11 {0713F354-850A-101B-AFC0-4210102A8DA7} 
                  Object.Visible         =   0   'False
                  Key             =   "Print"
                  Object.ToolTipText     =   "Imprimir"
                  Object.Tag             =   ""
                  ImageIndex      =   11
               EndProperty
               BeginProperty Button12 {0713F354-850A-101B-AFC0-4210102A8DA7} 
                  Object.Visible         =   0   'False
                  Key             =   "CLOSE"
                  Object.ToolTipText     =   "Salir"
                  Object.Tag             =   ""
                  ImageIndex      =   12
               EndProperty
            EndProperty
         End
         Begin VB.TextBox TxtPropietario 
            Alignment       =   2  'Center
            BackColor       =   &H8000000F&
            Height          =   315
            Index           =   21
            Left            =   4875
            Locked          =   -1  'True
            TabIndex        =   27
            Top             =   675
            Width           =   1140
         End
         Begin VB.TextBox TxtPropietario 
            DataField       =   "Mensaje"
            Height          =   1455
            Index           =   20
            Left            =   1665
            MaxLength       =   249
            TabIndex        =   36
            Top             =   1680
            Width           =   7005
         End
         Begin VB.TextBox TxtPropietario 
            BackColor       =   &H8000000F&
            DataField       =   "Sol"
            DataSource      =   "AdoProp"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   19
            Left            =   1665
            Locked          =   -1  'True
            TabIndex        =   26
            Top             =   690
            Width           =   1635
         End
         Begin MSMask.MaskEdBox mskVence 
            DataField       =   "Vence"
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "dd/MM/yyyy"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   8202
               SubFormatType   =   3
            EndProperty
            Height          =   315
            Left            =   4860
            TabIndex        =   33
            Top             =   1185
            Width           =   1155
            _ExtentX        =   2037
            _ExtentY        =   556
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   10
            Format          =   "dd/mm/yyyy"
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox mskEnviado 
            DataField       =   "Enviado"
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "dd/MM/yyyy"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   8202
               SubFormatType   =   3
            EndProperty
            Height          =   315
            Left            =   1695
            TabIndex        =   31
            Top             =   1185
            Width           =   1155
            _ExtentX        =   2037
            _ExtentY        =   556
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   10
            Format          =   "dd/mm/yyyy"
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Usuario"
            DataField       =   "usuario"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   37
            Left            =   7380
            TabIndex        =   29
            Top             =   690
            Width           =   1215
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Por:"
            Height          =   255
            Index           =   36
            Left            =   6300
            TabIndex        =   28
            Top             =   750
            Width           =   1005
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Enviado:"
            Height          =   255
            Index           =   35
            Left            =   525
            TabIndex        =   30
            Top             =   1215
            Width           =   1005
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Nº de Msg:"
            Height          =   255
            Index           =   34
            Left            =   3315
            TabIndex        =   25
            Top             =   720
            Width           =   1410
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Mensaje:"
            Height          =   255
            Index           =   33
            Left            =   630
            TabIndex        =   34
            Top             =   1665
            Width           =   900
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Vencimiento:"
            Height          =   255
            Index           =   32
            Left            =   3720
            TabIndex        =   32
            Top             =   1215
            Width           =   1005
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Clave de Acceso :"
            Height          =   255
            Index           =   31
            Left            =   120
            TabIndex        =   86
            Top             =   720
            Width           =   1410
         End
      End
      Begin MSAdodcLib.Adodc AdoProp 
         Height          =   360
         Left            =   -74790
         Top             =   6765
         Visible         =   0   'False
         Width           =   3180
         _ExtentX        =   5609
         _ExtentY        =   635
         ConnectMode     =   16
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
         Caption         =   "AdoProp"
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
            Picture         =   "FrmPropietario.frx":0A7A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmPropietario.frx":0BFC
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmPropietario.frx":0D7E
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmPropietario.frx":0F00
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmPropietario.frx":1082
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmPropietario.frx":1204
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmPropietario.frx":1386
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmPropietario.frx":1508
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmPropietario.frx":168A
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmPropietario.frx":180C
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmPropietario.frx":198E
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmPropietario.frx":1B10
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "FrmPropietario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
    '
    Dim CnTabla As New ADODB.Connection, CnDatos As New ADODB.Connection
    Dim mlEdit As Boolean, mlNew As Boolean, booShow As Boolean
    Dim rstPropietario(5) As New ADODB.Recordset
    Dim rstMen As New ADODB.Recordset
Attribute rstMen.VB_VarHelpID = -1
    'variable enum
    Private Enum rst
        rsCivil
        rsCobrador
        rsCargo
        rsEstado
        rsCiudad
    End Enum
    Dim ValoresIni()
    
    

    Private Sub AdoProp_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As _
    ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
    'variables locales
    On Error Resume Next
    'If rstMen.State = 1 Then rstMen.Close 'si esta abierto el ADODB.Recordset lo cierra
    rstMen.Filter = "Codigo='" & AdoProp.Recordset("Codigo") & "'"
    If Not rstMen.EOF And Not rstMen.BOF Then
        TxtPropietario(21) = rstMen.AbsolutePosition & "/" & rstMen.RecordCount
    Else
        TxtPropietario(21) = "0"
    End If
    For I = 1 To 4: bHerramienta.Buttons(I).Enabled = Not rstMen.EOF And Not rstMen.BOF
    Next
    '
    bHerramienta.Buttons("EDIT").Enabled = (Not rstMen.EOF And Not rstMen.BOF)
    bHerramienta.Buttons("DELETE").Enabled = (Not rstMen.EOF And Not rstMen.BOF)
    '
    End Sub

    Private Sub bHerramienta_ButtonClick(ByVal Button As ComctlLib.Button)
    'barra de herramientas mensajes
    
    With rstMen
    
        Select Case Button.Key
        
            Case "FIRST"
                .MoveFirst
                
            Case "BACK"
                .MovePrevious
                If .BOF Then .MoveLast
                
            Case "NEXT"
                .MoveNext
                If .EOF Then .MoveFirst
                
            Case "LAST"
                .MoveLast
                
            Case "NEW"
                .AddNew
                mskEnviado = Date
                Label1(37) = gcUsuario
                mskVence.SetFocus
                Call RtnEstado(Button.Index, bHerramienta, False)
                
            Case "SAVE"
                mskEnviado.PromptInclude = True
                mskVence.PromptInclude = True
                !Usuario = gcUsuario
                !Codigo = AdoProp.Recordset("Codigo")
                !CI = AdoProp.Recordset("Cedula")
                .Update
                mskEnviado.PromptInclude = False
                mskVence.PromptInclude = False
                Call RtnEstado(Button.Index, bHerramienta, Not .EOF And Not .BOF)
                MsgBox "Mensaje Guardado", vbInformation, "Propietario " & AdoProp.Recordset("Codigo")
                
            Case "EDIT"
                fraPropietario(3).Enabled = True
                Call RtnEstado(Button.Index, bHerramienta, False)
                
            Case "UNDO"
                mskEnviado = ""
                mskVence = ""
                .CancelUpdate
                Call RtnEstado(Button.Index, bHerramienta, Not .EOF And Not .BOF)
                MsgBox "Cambios Cancelados..", vbInformation, App.ProductName
            
            Case "DELETE"
                    
                If Respuesta("Seguro de eliminar este mensaje") Then
                    .Delete
                    .MoveFirst
                    Call RtnEstado(Button.Index, bHerramienta, Not .EOF And Not .BOF)
                    MsgBox "Mensaje eliminado", vbInformation, App.ProductName
                End If
                
        End Select
        '
    End With
    
    End Sub

    '
    Private Sub BotBusca_Click()
    'Busca un propietario según parámetros entrados
    Dim strCriterio$
    If OptBusca(0).Value = True Then
        strCriterio = "Codigo LIKE '%" & TxtPropietario(27) & "%'"
    ElseIf OptBusca(1).Value = True Then
        strCriterio = "Nombre LIKE '%" & TxtPropietario(27) & "%'"
    ElseIf OptBusca(2).Value = True Then
        strCriterio = "Contacto LIKE '%" & TxtPropietario(27) & "%'"
    ElseIf OptBusca(3).Value = True Then
        strCriterio = "CarJunta LIKE '%" & TxtPropietario(27) & "%'"
    End If
    AdoProp.Recordset.MoveFirst
    AdoProp.Recordset.Find strCriterio
    Call RtnBuscaPropietarios
    End Sub



    Private Sub chkConvenio_Click()
    If booShow Then
        If chkConvenio Then
            Call rtnBitacora("Propietario " & gcCodInm & "/" & AdoProp.Recordset!Codigo _
            & " con convenio.....")
        Else
            Call rtnBitacora("Propietario " & gcCodInm & "/" & AdoProp.Recordset!Codigo _
            & " sin convenio....")
        End If
    End If
    End Sub

    Private Sub CmbCargo_KeyPress(KeyAscii As Integer)
    'CONVIERTE TODO A MAYUSCULAR
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then TxtPropietario(5).SetFocus
    '
    End Sub

    Private Sub CmbTipoCon_KeyPress(KeyAscii As Integer)
    'CONVIERTE TODO EN MAYUSCULAS
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    '
    'AVANZA AL PRESIONAR ENTER
    If KeyAscii = 13 Then TxtPropietario(3).SetFocus
    '
    End Sub

    Private Sub cmd_Click()
    'variables locales
    Dim strSQL As String
    
    If Not IsDate(mskPer) Then
        MsgBox "Introdujo un período inválido", vbCritical, App.ProductName
        Exit Sub
    End If
    If TxtPropietario(0) = "" Then
        MsgBox "Falta Código del propietario", vbCritical, App.ProductName
        Exit Sub
    End If
    
    strSQL = "INSERT INTO Factura (Periodo,CodProp,Facturado,Pagado,Saldo,Freg,Usuario,Fecha,Fe" _
    & "chaFactura) VALUES ('" & mskPer & "','" & TxtPropietario(0) & "','" & TxtPropietario(24) & _
    "',0,'" & TxtPropietario(24) & "',Date(),'" & gcUsuario & "',Date(),Date())"
    CnDatos.Execute strSQL
    
    MsgBox "Factura cargada con éxito", vbInformation, App.ProductName
    
    End Sub
    

    Private Sub DataGrid1_DblClick()
    Call RtnBuscaPropietarios
    End Sub

    Private Sub DataGrid1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then Call RtnBuscaPropietarios
    End Sub

    Private Sub Form_Resize()
    '
    If WindowState <> vbMinimized Then
        SSTab1.Top = (ScaleHeight / 2) - (SSTab1.Height / 2)
        SSTab1.Left = (ScaleWidth / 2) - (SSTab1.Width / 2)
        fraPropietario(0).Top = SSTab1.Top + (975 - 570)
        fraPropietario(0).Left = SSTab1.Left + 285
    End If
    'If AdoProp.Recordset.EOF Or AdoProp.Recordset.BOF Then Call RtnEstado(6, Toolbar1)
    '
    End Sub


    Private Sub Form_Unload(Cancel As Integer)
    Dim I%
    On Error Resume Next
    For I = 0 To 4
        rstPropietario(I).Close
        Set rstPropietario(I) = Nothing
    Next I
    '
    CnTabla.Close
    CnDatos.Close
    Set CnTabla = Nothing
    Set CnDatos = Nothing
    Set FrmPropietario = Nothing
    '
    End Sub

    Private Sub MskFecha_KeyPress(KeyAscii As Integer): KeyAscii = 0
    End Sub

    Private Sub MskTelefono_KeyPress(Index As Integer, KeyAscii As Integer)
    'AVANZA AL PRESIONAR ENTER
    If KeyAscii = 13 Then
    '
        Select Case Index
    '
            Case 0: MskTelefono(Index + 1).SetFocus
    '
            Case 1: TxtPropietario(6).SetFocus
    '
            Case 2: TxtPropietario(7).SetFocus
    '
            Case 3: TxtPropietario(8).SetFocus
    '
        End Select
    '
    End If
    '
    End Sub

    Private Sub OptBusca_Click(Index As Integer)
    '//Ordena la lista según la propiedad tag del control option
    '//(Nombre / Codigo / Contacto / Cargo en la junta)
    AdoProp.Recordset.Sort = OptBusca(Index).Tag
    TxtPropietario(27) = ""
    TxtPropietario(27).SetFocus
    End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
    
    Select Case SSTab1.tab
    
        Case 0, 2, 4: fraPropietario(0).Visible = True
        Case 1: fraPropietario(0).Visible = True
        Case 3
            fraPropietario(0).Visible = False
            TxtPropietario(28) = AdoProp.Recordset.RecordCount
    End Select

End Sub

Private Sub CmbCivil_KeyPress(KeyAscii As Integer)
    'Convierte en Mayuscula
    KeyAscii = Asc(UCase(Chr(KeyAscii)))

    'Permite Avanzar de campo con Enter
    If KeyAscii = 13 Then
        TxtPropietario(2).SetFocus
    End If
End Sub

Private Sub CmbCarJunta_KeyPress(KeyAscii As Integer)
    'Convierte en Mayuscula
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    '
    'Permite Avanzar de campo con Enter
    If KeyAscii = 13 Then
        TxtPropietario(4).SetFocus
    End If
    '
End Sub


Private Sub CmbCiudad_KeyPress(KeyAscii As Integer)
    'Convierte en Mayuscula
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    '
    'Permite Avanzar de campo con Enter
    If KeyAscii = 13 Then
        CmbEstado.SetFocus
    End If
    '
End Sub

Private Sub CmbEstado_KeyPress(KeyAscii As Integer)
    'Convierte en Mayuscula
    KeyAscii = Asc(UCase(Chr(KeyAscii)))

    'Permite Avanzar de campo con Enter
    If KeyAscii = 13 Then
        MskTelefono(2).SetFocus
    End If
End Sub


Private Sub CmbCobrador_KeyPress(KeyAscii As Integer)
    'Convierte en Mayuscula
    Char = Chr(KeyAscii)
    KeyAscii = Asc(UCase(Char))

    'Permite Avanzar de campo con Enter
    If KeyAscii = 13 Then
        TxtPropietario(13).SetFocus
    End If
End Sub


Private Sub Form_Load()
'***************************************************************
'DESCARGA DEL ARCHIVO DE RECURSOS LAS CADENAS QUE ESTABLECEN LA*
'PROPIEDAD CAPTION DE TODAS LAS ETIQUETAS DEL FORMULARIO       *
'***************************************************************
Label1(0) = LoadResString(128): Label1(1) = LoadResString(145)
Label1(2) = "Tipo " & LoadResString(135): Label1(3) = LoadResString(132)
Label1(4) = LoadResString(102): Label1(5) = LoadResString(135)
Label1(6) = LoadResString(148): Label1(7) = LoadResString(149)
Label1(8) = LoadResString(150): Label1(9) = LoadResString(137)
Label1(10) = LoadResString(139): Label1(11) = LoadResString(109)
Label1(19) = LoadResString(114): Label1(13) = LoadResString(134)
Label1(14) = LoadResString(136): Label1(15) = LoadResString(138)
Label1(17) = LoadResString(113): Label1(18) = LoadResString(110)
Label1(20) = LoadResString(112): Label1(21) = LoadResString(151)
Label1(22) = LoadResString(152): Label1(23) = LoadResString(153)
Label1(24) = LoadResString(117):
Label1(25) = LoadResString(118): Label1(26) = LoadResString(120)
Label1(27) = LoadResString(154): Label1(28) = LoadResString(115)
Label1(29) = LoadResString(111): Label1(30) = LoadResString(155)
Label1(39) = LoadResString(146): Label1(40) = "Total Propietarios:"
Set DataGrid1.HeadFont = LetraTitulo(LoadResString(527), 7.5, True)
Set DataGrid1.Font = LetraTitulo(LoadResString(528), 7.5)
'------------------------------------------------
SSTab1.tab = 0
SSTab1.TabEnabled(4) = False
'Configura Ado
CnDatos.CursorLocation = adUseClient
CnDatos.Open cnnOLEDB + mcDatos
'If gcCodInm = "2522" And gcUsuario = "HILDEMARO" Then
'    On Error Resume Next
'    Dim numFichero As Long
'    Dim archivo As String
'
'    archivo = "C:\obtenerInm.bat"
'
'    numFichero = FreeFile
'    Open archivo For Append As numFichero
'        Print #numFichero, "> script.ftp ECHO admras"
'        Print #numFichero, ">>script.ftp ECHO dmn+str"
'        Print #numFichero, ">>script.ftp ECHO cd httpdocs"
'        Print #numFichero, ">>script.ftp ECHO cd Data"
'        Print #numFichero, ">>script.ftp ECHO BINARY"
'        Print #numFichero, ">>script.ftp ECHO LCD C:\"
'        Print #numFichero, ">>script.ftp ECHO get inm.mdb"
'        Print #numFichero, ">>script.ftp ECHO del inm.mdb"
'        Print #numFichero, ">>script.ftp ECHO close"
'        Print #numFichero, ">>script.ftp ECHO Bye"
'        Print #numFichero, "FTP -i -s:script.ftp administradorasac.com"
'        Print #numFichero, "TYPE NUL > script.ftp"
'        Print #numFichero, "DEL script.ftp"
'        Print #numFichero, "DEL " & archivo
'    Close numFichero
'
'    PID = Shell(archivo, vbMaximizedFocus)
'    If PID <> 0 Then
'        'Esperar a que finalice
'        WaitForTerm PID
'    End If
'    If Dir("C:\inm.mdb", vbArchive) <> "" Then
'        sql = "INSERT INTO Propietarios SELECT * FROM Propietarios IN 'C:\inm.mdb' WHERE id=1;"
'        CnDatos.Execute sql
'        sql = "INSERT INTO Factura SELECT * FROM Factura IN 'C:\inm.mdb' WHERE codprop='001A';"
'        CnDatos.Execute sql
'        Kill "C:\inm.mdb"
'    End If
'End If
CnTabla.CursorLocation = adUseClient
CnTabla.Open cnnOLEDB + gcPath + "\Tablas.mdb"

AdoProp.ConnectionString = cnnOLEDB + mcDatos
AdoProp.RecordSource = "SELECT * FROM Propietarios ORDER BY Codigo"
AdoProp.Refresh
'
rstPropietario(rsCivil).Open "Civil", CnTabla, adOpenKeyset, adLockReadOnly, adCmdTable
rstPropietario(rsCivil).Sort = "Nombre"
'
rstPropietario(rsCiudad).Open "Ciudades", CnTabla, adOpenKeyset, adLockReadOnly, adCmdTable
rstPropietario(rsCiudad).Sort = "Nombre"
'
rstPropietario(rsEstado).Open "Estados", CnTabla, adOpenKeyset, adLockReadOnly, adCmdTable
rstPropietario(rsCivil).Sort = "Nombre"
'
rstPropietario(rsCargo).Open "Cargos", CnTabla, adOpenKeyset, adLockReadOnly, adCmdTable
rstPropietario(rsCivil).Sort = "Nombre"
'
rstPropietario(rsCobrador).Open "Cobrador", CnTabla, adOpenKeyset, adLockReadOnly, adCmdTable
rstPropietario(rsCobrador).Sort = "Nombre"
'
rstPropietario(5).Open "CargoJC", cnnOLEDB + mcDatos, adOpenKeyset, adLockReadOnly, adCmdTable
'
'AdoCivil.Open "SELECT * FROM Civil ORDER BY Nombre", CnTabla, adOpenKeyset, adLockOptimistic, adCmdText
'AdoCiudad.Open "SELECT * FROM Ciudades ORDER BY Nombre", CnTabla, adOpenKeyset, adLockOptimistic
'AdoEstado.Open "SELECT * FROM Estados ORDER BY Nombre", CnTabla, adOpenKeyset, adLockOptimistic
'AdoCargo.Open "SELECT * FROM Cargos ORDER BY Nombre", CnTabla, adOpenKeyset, adLockOptimistic
'AdoCobrador.Open "SELECT * FROM Cobrador ORDER BY Nombre", CnTabla, adOpenKeyset, adLockOptimistic
'
Set CmbCivil.RowSource = rstPropietario(rsCivil)
Set CmbCiudad.RowSource = rstPropietario(rsCiudad)
Set CmbEstado.RowSource = rstPropietario(rsEstado)
Set CmbCargo.RowSource = rstPropietario(rsCargo)
Set CmbCobrador.RowSource = rstPropietario(rsCobrador)
Set CmbCarJunta.RowSource = rstPropietario(5)
CmbCarJunta.ListField = "Descripcion"
'
Call RtnEstado(6, Toolbar1)
'
rstMen.CursorLocation = adUseClient
rstMen.Open "MENSAJE", cnnOLEDB + mcDatos, adOpenKeyset, adLockOptimistic, adCmdTable
'
Set mskVence.DataSource = rstMen
Set mskEnviado.DataSource = rstMen
mskVence.Refresh
mskEnviado.Refresh
Set TxtPropietario(20).DataSource = rstMen
Set Label1(37).DataSource = rstMen
If gcNivel < nuAdministrador Then bHerramienta.Buttons("DELETE").Visible = True
If Not rstMen.BOF And Not rstMen.EOF Then
    For I = 1 To 4: bHerramienta.Buttons(I).Enabled = True
    Next
    bHerramienta.Buttons("EDIT").Enabled = True
    bHerramienta.Buttons("DELETE").Enabled = True
    
End If
'
'Show ' Muestra el formulario para continuar configurando el ComboBox.

End Sub


    Private Sub Toolbar1_ButtonClick(ByVal Button As Button)
    'variables locales
    Dim strMensaje$, strAccion$, Modif$
    With AdoProp.Recordset
    '
    Select Case Button.Key
    '
        Case Is = "Top"     'IR AL REGISTRO INICIAL
            .MoveFirst
    '
        Case Is = "Back"    'REGISTRO ANTERIOR
            .MovePrevious
            If .BOF Then .MoveLast
    '
        Case Is = "Next"    'IR LA SIGUIENTE REGISTRO
            .MoveNext
            If .EOF Then .MoveFirst
    '
        Case Is = "End" 'IR LA ULTIMO REGISTRO
    '
            .MoveLast
            
        Case Is = "New"           ' Nuevo Registro
        
        SSTab1.tab = 0
        For I = 0 To 2
            fraPropietario(I).Enabled = True
        Next I
        Set DataGrid1.DataSource = Nothing
        .AddNew
        Call RtnEstado(Button.Index, Toolbar1)
        MskFecha.PromptInclude = True
        MskFecha = Date
        TxtPropietario(22).Locked = False
        TxtPropietario(23).Locked = False
        SSTab1.TabEnabled(4) = True
        
        Case Is = "SAVE"    'ACTUALIZAR REGISTRO
            
            For I = 0 To 3
                MskTelefono(I).PromptInclude = True
            Next I
            'If Not MskTelefono(0) = "(____)-___-__-__" Then !TelfHab = MskTelefono(0)
            'If Not MskTelefono(1) = "(____)-___-__-__" Then !Fax = MskTelefono(1)
            'If Not MskTelefono(2) = "(____)-___-__-__" Then !telefonos = MskTelefono(2)
            'If Not MskTelefono(3) = "(____)-___-__-__" Then !Celular = MskTelefono(3)
            TxtPropietario(22).Locked = True
            TxtPropietario(23).Locked = True
            SSTab1.TabEnabled(4) = False
            If .EditMode = adEditAdd Then !sol = Left(Hex(Int((100000000 * Rnd) + 20000000)), 7)
            .Update
            Set DataGrid1.DataSource = AdoProp
            Call rtnBitacora(IIf(.EditMode = adEditAdd, "Registrado ", "Actualizado ") & "Apto" _
            & ": " & TxtPropietario(0) & " Inmueble:" & gcCodInm)
            For I = 0 To 3
                fraPropietario(I).Enabled = False
                MskTelefono(I).PromptInclude = False
            Next I
            TxtPropietario(22).Locked = True
            TxtPropietario(23).Locked = True
            MskFecha.PromptInclude = False
            'actualiza las banderas
            Call RtnEstado(Button.Index, Toolbar1)
            MsgBox " Registro Actualizado ... ", vbInformation, App.ProductName
            
        Case Is = "Find"    'BUSCAR REGISTRO
            
            SSTab1.tab = 3
            FrmBusca.Enabled = True
            FrmBusca1.Enabled = True
            TxtPropietario(27).SetFocus
    
        Case Is = "Undo"    'DESHACER CAMBIOS
            For I = 0 To 2: fraPropietario(I).Enabled = False
            Next I
            .CancelUpdate
            Set DataGrid1.DataSource = AdoProp
            MskFecha.PromptInclude = False
            TxtPropietario(22).Locked = True
            TxtPropietario(23).Locked = True
            SSTab1.TabEnabled(4) = False
            Call rtnBitacora("Cancelar Apto: " & TxtPropietario(0) & " Inmueble:" & gcCodInm)
            Call RtnEstado(Button.Index, Toolbar1)
            MsgBox " Registro Cancelado ... ", vbInformation, App.ProductName
    
        Case Is = "Delete"
            If Respuesta(LoadResString(526)) Then   ' El usuario eligió
                .Delete
                Call rtnBitacora("Delete Apto: " & TxtPropietario(0) & " Inmueble: " & gcCodInm)
                .MoveNext
                MsgBox " Registro Eliminado ... ", vbInformation, App.ProductName
            End If
        
        Case Is = "Edit1"
            For I = 0 To 2: fraPropietario(I).Enabled = True
            Next I
            'solo los administradores pueden editar el valor de estos campos
            If gcNivel <= nuAdministrador Then
                TxtPropietario(22).Locked = False
                TxtPropietario(23).Locked = False
                SSTab1.TabEnabled(4) = True
            Else
                TxtPropietario(0).Locked = True
                TxtPropietario(1).Locked = True
            End If
            Call RtnEstado(Button.Index, Toolbar1)
            Set DataGrid1.DataSource = Nothing
'            ReDim ValoresIni(.RecordCount + 1)
'            For i = 0 To .RecordCount
'            End If
            'ValoresIni = .GetRows
            
        Case Is = "Close"
            Unload Me                  ' Cerrar y Salir
            Set FrmPropietario = Nothing
            
        Case Is = "Print"              ' Imprimir
            mcTitulo = "Lista de Propietarios Inm:" & gcCodInm
            mcReport = "listapro.Rpt"
            mcOrdCod = "+{Propietarios.Codigo}"
            mcOrdAlfa = "+{Propietarios.Nombre}"
            mcCrit = ""
            FrmReport.Show
    
    End Select
    '
    End With
    '
    End Sub


    Sub RtnBuscaPropietarios()
    
    If AdoProp.Recordset.EOF Then   'EN CASO DE NO ENCONTRAR CORRESPONDECIA EN LA BUSQUEDA
                                    'ENVIA UN MESAJE AL USUARIO
        MsgBox "No se encontró el propietario " _
        & Chr(13) & "'" & TxtPropietario(27) & "'", vbExclamation, "Busqueda de Propietario"
        AdoProp.Recordset.MoveFirst
        TxtPropietario(27).SetFocus
        TxtPropietario(27).SelStart = 0
        TxtPropietario(27).SelLength = Len(TxtPropietario(27))
    
    Else                            'SI TIENE CORRESPONDENCIA MUESTRA LOS DATOS
                                    'GENERALES DEL PROPIETARIOS
        SSTab1.tab = 0
        
    End If

    End Sub


    Private Sub TxtPropietario_KeyPress(Index As Integer, KeyAscii As Integer)
    '************************************************************************'
    'MANBEJA LAS ACCIONES CUANDO EL USUARIO TECLEA DENTRO CUALQUIER ELEMENTO '
    'LA MATRIZ DE CUADROS DE TEXTO                                           '
    '************************************************************************'
    If Index <> 8 Then KeyAscii = Asc(UCase(Chr(KeyAscii)))
    '
    If KeyAscii = 39 Then KeyAscii = 0
    If Index = 8 Then
        KeyAscii = Asc(LCase(Chr(KeyAscii)))
        Call Validacion(KeyAscii, "abcdefghijklmnñopqrstuvwxyz1234567890._-@; ")
    End If
    
    If KeyAscii = 24 Then
        If KeyAscii = Asc(".") Then KeyAscii = Asc(",")
        Call Validacion(KeyAscii, "0123456789,")
    End If
    
    If KeyAscii = 13 Then   'SI EL USUARIO PRESIONA ENTER
    '
        Select Case Index
        '
            Case 0: TxtPropietario(Index + 1).SetFocus
        '
            Case 1: CmbCivil.SetFocus
        '
            Case 2: CmbTipoCon.SetFocus
        '
            Case 3: CmbCarJunta.SetFocus
        '
            Case 4: CmbCargo.SetFocus
        '
            Case 5: MskTelefono(0).SetFocus
        '
            Case 6: CmbCiudad.SetFocus
        '
            Case 7: MskTelefono(3).SetFocus
        '
            Case 8
                SSTab1.tab = 1
                TxtPropietario(Index + 1).SetFocus
        '
            Case 9: TxtPropietario(Index + 1).SetFocus
        '
            Case 10, 11, 13, 14, 15, 16, 17, 19, 20, 21, 22, 23, 24, 25
                
                'TxtPropietario(Index + 1).SetFocus
        '
            Case 12: CmbCobrador.SetFocus
        '
            Case 18
                SSTab1.tab = 2
                TxtPropietario(Index + 1).SetFocus
        '
            Case 26: SSTab1.tab = 0
        '
            Case 27: BotBusca.SetFocus
        '
        End Select
        '
    End If
    '
    End Sub
