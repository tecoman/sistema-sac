VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Begin VB.Form FrmParamFact 
   Caption         =   "Parámetros de Facturación"
   ClientHeight    =   7290
   ClientLeft      =   1260
   ClientTop       =   1560
   ClientWidth     =   9660
   ControlBox      =   0   'False
   Icon            =   "FrmParamFact.frx":0000
   LinkTopic       =   "Form3"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   7290
   ScaleWidth      =   9660
   Tag             =   "1"
   WindowState     =   2  'Maximized
   Begin ComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   480
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   9660
      _ExtentX        =   17039
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
            Object.Visible         =   0   'False
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
            Object.Visible         =   0   'False
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
   Begin TabDlg.SSTab SSTab1 
      Height          =   7680
      Left            =   1740
      TabIndex        =   0
      Top             =   795
      Width           =   8340
      _ExtentX        =   14711
      _ExtentY        =   13547
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
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
      TabCaption(0)   =   "Parámetros"
      TabPicture(0)   =   "FrmParamFact.frx":000C
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Lista de Inmuebles"
      TabPicture(1)   =   "FrmParamFact.frx":0028
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "DataGrid1"
      Tab(1).Control(1)=   "FrmBusca1"
      Tab(1).Control(2)=   "FrmBusca"
      Tab(1).ControlCount=   3
      Begin MSDataGridLib.DataGrid DataGrid1 
         Height          =   5520
         Left            =   -74835
         TabIndex        =   10
         Top             =   450
         Width           =   8010
         _ExtentX        =   14129
         _ExtentY        =   9737
         _Version        =   393216
         AllowUpdate     =   0   'False
         BackColor       =   16777215
         HeadLines       =   1
         RowHeight       =   15
         RowDividerStyle =   3
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
         Caption         =   "PARAMETROS DE FACURACION"
         ColumnCount     =   2
         BeginProperty Column00 
            DataField       =   "CodInm"
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
         SplitCount      =   1
         BeginProperty Split0 
            AllowRowSizing  =   0   'False
            AllowSizing     =   0   'False
            BeginProperty Column00 
               Alignment       =   2
               ColumnWidth     =   915,024
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   5760
            EndProperty
         EndProperty
      End
      Begin VB.Frame Frame1 
         Enabled         =   0   'False
         Height          =   7095
         Left            =   165
         TabIndex        =   11
         Top             =   375
         Width           =   8070
         Begin VB.CheckBox Check1 
            Caption         =   "Detalle fondos en el recibo de pago"
            DataField       =   "FactFV"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   4
            Left            =   4245
            TabIndex        =   58
            Top             =   3615
            Width           =   3765
         End
         Begin VB.TextBox txtParam 
            Alignment       =   1  'Right Justify
            DataField       =   "CodFondoE"
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   8202
               SubFormatType   =   1
            EndProperty
            Height          =   315
            Index           =   17
            Left            =   7155
            TabIndex        =   53
            Top             =   3090
            Width           =   700
         End
         Begin VB.TextBox txtParam 
            DataField       =   "CodInm"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   360
            Index           =   13
            Left            =   1335
            Locked          =   -1  'True
            TabIndex        =   50
            Top             =   353
            Width           =   930
         End
         Begin VB.TextBox txtParam 
            DataField       =   "Nombre"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   360
            Index           =   14
            Left            =   2325
            Locked          =   -1  'True
            TabIndex        =   49
            Top             =   360
            Width           =   5520
         End
         Begin VB.TextBox txtParam 
            Alignment       =   1  'Right Justify
            DataField       =   "Honorarios"
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "#,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   8202
               SubFormatType   =   0
            EndProperty
            Height          =   315
            Index           =   0
            Left            =   3045
            TabIndex        =   38
            Text            =   "0,00"
            Top             =   923
            Width           =   1215
         End
         Begin VB.TextBox txtNota 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   945
            Left            =   225
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   36
            Top             =   5925
            Width           =   7635
         End
         Begin VB.TextBox txtParam 
            Alignment       =   1  'Right Justify
            DataField       =   "Facturacion"
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "#,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   8202
               SubFormatType   =   0
            EndProperty
            Height          =   315
            Index           =   6
            Left            =   2775
            TabIndex        =   35
            Top             =   3083
            Width           =   1485
         End
         Begin VB.TextBox txtParam 
            Alignment       =   1  'Right Justify
            DataField       =   "DiasPenal"
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "#,##0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   8202
               SubFormatType   =   0
            EndProperty
            Height          =   315
            Index           =   12
            Left            =   7155
            TabIndex        =   30
            Top             =   2723
            Width           =   700
         End
         Begin VB.TextBox txtParam 
            Alignment       =   1  'Right Justify
            DataField       =   "ClausulaPenal"
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "#,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   8202
               SubFormatType   =   0
            EndProperty
            Height          =   315
            Index           =   5
            Left            =   3375
            Locked          =   -1  'True
            TabIndex        =   29
            Top             =   2723
            Width           =   885
         End
         Begin VB.TextBox txtParam 
            Alignment       =   1  'Right Justify
            DataField       =   "honomorosidad"
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "#,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   8202
               SubFormatType   =   0
            EndProperty
            Height          =   315
            Index           =   11
            Left            =   7155
            TabIndex        =   28
            Top             =   2363
            Width           =   700
         End
         Begin VB.TextBox txtParam 
            Alignment       =   1  'Right Justify
            DataField       =   "CostTelegrama"
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "#,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   8202
               SubFormatType   =   0
            EndProperty
            Height          =   315
            Index           =   4
            Left            =   3375
            TabIndex        =   15
            Top             =   2363
            Width           =   885
         End
         Begin VB.TextBox txtParam 
            Alignment       =   1  'Right Justify
            DataField       =   "CostCarta"
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "#,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   8202
               SubFormatType   =   0
            EndProperty
            Height          =   315
            Index           =   3
            Left            =   3375
            TabIndex        =   14
            Top             =   2003
            Width           =   885
         End
         Begin VB.TextBox txtParam 
            Alignment       =   1  'Right Justify
            DataField       =   "PorcFondo"
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "#,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   8202
               SubFormatType   =   0
            EndProperty
            Height          =   315
            Index           =   7
            Left            =   7155
            TabIndex        =   16
            Top             =   923
            Width           =   700
         End
         Begin VB.TextBox txtParam 
            Alignment       =   1  'Right Justify
            DataField       =   "Interes"
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "#,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   8202
               SubFormatType   =   0
            EndProperty
            Height          =   315
            Index           =   8
            Left            =   7155
            TabIndex        =   17
            Top             =   1283
            Width           =   700
         End
         Begin VB.TextBox txtParam 
            Alignment       =   1  'Right Justify
            DataField       =   "MesesMora"
            BeginProperty DataFormat 
               Type            =   0
               Format          =   "#.##0,00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   8202
               SubFormatType   =   0
            EndProperty
            Height          =   315
            Index           =   2
            Left            =   3780
            TabIndex        =   13
            Top             =   1643
            Width           =   480
         End
         Begin VB.TextBox txtParam 
            Alignment       =   1  'Right Justify
            DataField       =   "MesesInt"
            BeginProperty DataFormat 
               Type            =   0
               Format          =   "#.##0,00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   8202
               SubFormatType   =   0
            EndProperty
            Height          =   315
            Index           =   1
            Left            =   3795
            TabIndex        =   12
            Top             =   1283
            Width           =   465
         End
         Begin VB.TextBox txtParam 
            Alignment       =   1  'Right Justify
            DataField       =   "Gestion"
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "#,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   8202
               SubFormatType   =   0
            EndProperty
            Height          =   315
            Index           =   9
            Left            =   7155
            TabIndex        =   18
            Top             =   1643
            Width           =   700
         End
         Begin VB.TextBox txtParam 
            Alignment       =   1  'Right Justify
            DataField       =   "CheqDev"
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "#,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   8202
               SubFormatType   =   0
            EndProperty
            Height          =   315
            Index           =   10
            Left            =   7155
            TabIndex        =   19
            Top             =   2003
            Width           =   700
         End
         Begin VB.CheckBox Check1 
            Caption         =   "Telegráma para Propietario con 10 Recibos Vencidos"
            DataField       =   "EmiteT2"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   8
            Left            =   1305
            TabIndex        =   27
            Top             =   5820
            Visible         =   0   'False
            Width           =   6015
         End
         Begin VB.CheckBox Check1 
            Caption         =   "Telegráma para Propietario con 8 Recibos Vencidos"
            DataField       =   "EmiteT1"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   7
            Left            =   1305
            TabIndex        =   26
            Top             =   5475
            Visible         =   0   'False
            Width           =   5850
         End
         Begin VB.CheckBox Check1 
            Caption         =   "Cheques Devueltos"
            DataField       =   "FactCH"
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
            Index           =   3
            Left            =   240
            TabIndex        =   23
            Top             =   4665
            Width           =   2415
         End
         Begin VB.CheckBox Check1 
            Caption         =   "Gestión de Cobranza"
            DataField       =   "FactGC"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   2
            Left            =   240
            TabIndex        =   22
            Top             =   4305
            Width           =   2415
         End
         Begin VB.CheckBox Check1 
            Caption         =   "Intereses de Mora"
            DataField       =   "FactIM"
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
            Index           =   1
            Left            =   240
            TabIndex        =   21
            Top             =   3975
            Width           =   2415
         End
         Begin VB.CheckBox Check1 
            Caption         =   "Gastos Administrativos"
            DataField       =   "FactGA"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   0
            Left            =   240
            TabIndex        =   20
            Top             =   3615
            Width           =   2415
         End
         Begin MSDataListLib.DataList DataList1 
            Height          =   1035
            Index           =   0
            Left            =   4245
            TabIndex        =   54
            Top             =   4230
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   1826
            _Version        =   393216
            Enabled         =   0   'False
            ListField       =   "CodGasto"
         End
         Begin MSDataListLib.DataList DataList1 
            Height          =   1035
            Index           =   1
            Left            =   6270
            TabIndex        =   55
            Top             =   4230
            Width           =   1545
            _ExtentX        =   2725
            _ExtentY        =   1826
            _Version        =   393216
            Enabled         =   0   'False
            ListField       =   "CodGasto"
         End
         Begin VB.CheckBox Check1 
            Caption         =   "Imprimir Nº Cuenta - Banco"
            DataField       =   "FactCB"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   5
            Left            =   240
            TabIndex        =   24
            Top             =   5010
            Width           =   5385
         End
         Begin VB.CheckBox Check1 
            Caption         =   "Telegráma para Propietario con 6 Recibos Vencidos"
            DataField       =   "EmiteC3"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   6
            Left            =   1305
            TabIndex        =   25
            Top             =   5130
            Visible         =   0   'False
            Width           =   5385
         End
         Begin VB.Label lblParam 
            Caption         =   "Seleccionados:"
            Height          =   240
            Index           =   18
            Left            =   6285
            TabIndex        =   57
            Top             =   3990
            Width           =   1125
         End
         Begin VB.Label lblParam 
            Caption         =   "Fondos:"
            Height          =   240
            Index           =   17
            Left            =   4260
            TabIndex        =   56
            Top             =   4005
            Width           =   750
         End
         Begin VB.Label lblParam 
            Caption         =   "Cód. Fondo Esp. Emergencia:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   16
            Left            =   4335
            TabIndex        =   52
            Top             =   3135
            Width           =   2940
         End
         Begin VB.Label lblParam 
            Caption         =   "Inmueble:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   0
            Left            =   345
            TabIndex        =   51
            Top             =   405
            Width           =   900
         End
         Begin VB.Label lblParam 
            Caption         =   "Días Vencidos Clausula Penal:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   15
            Left            =   4335
            TabIndex        =   48
            Top             =   2760
            Width           =   2940
         End
         Begin VB.Label lblParam 
            Caption         =   "% Honorarios de Abogado:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   14
            Left            =   4335
            TabIndex        =   47
            Top             =   2400
            Width           =   2565
         End
         Begin VB.Label lblParam 
            Caption         =   "% Por Cheque Devuelto:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   13
            Left            =   4335
            TabIndex        =   46
            Top             =   2040
            Width           =   2565
         End
         Begin VB.Label lblParam 
            Caption         =   "% Gestión de Cobranza:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   12
            Left            =   4335
            TabIndex        =   45
            Top             =   1680
            Width           =   2565
         End
         Begin VB.Label lblParam 
            Caption         =   "% Intereses por Mora:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   11
            Left            =   4335
            TabIndex        =   44
            Top             =   1320
            Width           =   2565
         End
         Begin VB.Label lblParam 
            Caption         =   "% Fondo de Reserva:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   10
            Left            =   4335
            TabIndex        =   43
            Top             =   960
            Width           =   2565
         End
         Begin VB.Label lblParam 
            Caption         =   "Valor Clausula Penal:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   9
            Left            =   255
            TabIndex        =   42
            Top             =   2760
            Width           =   3090
         End
         Begin VB.Label lblParam 
            Caption         =   "Costo Emi. Telegrama Morosidad:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   8
            Left            =   255
            TabIndex        =   41
            Top             =   2400
            Width           =   3690
         End
         Begin VB.Label lblParam 
            Caption         =   "Gastos Emisión Carta de Deuda:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   7
            Left            =   255
            TabIndex        =   40
            Top             =   2040
            Width           =   3090
         End
         Begin VB.Label lblParam 
            Caption         =   "Meses Vencidos para Legal:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   6
            Left            =   255
            TabIndex        =   39
            Top             =   1680
            Width           =   3090
         End
         Begin VB.Label lblParam 
            Caption         =   "Nota Aviso de Cobro:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   5
            Left            =   270
            TabIndex        =   37
            Top             =   5595
            Width           =   2475
         End
         Begin VB.Label lblParam 
            Caption         =   "Límite Pre-Recibo:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   4
            Left            =   255
            TabIndex        =   34
            Top             =   3120
            Width           =   2475
         End
         Begin VB.Label lblParam 
            Caption         =   "Meses Vencidos para Intereses:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   2
            Left            =   255
            TabIndex        =   32
            Top             =   1320
            Width           =   3090
         End
         Begin VB.Label lblParam 
            Caption         =   "Honorarios por Inmueble:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   1
            Left            =   255
            TabIndex        =   31
            Top             =   960
            Width           =   2475
         End
         Begin VB.Line Line1 
            BorderColor     =   &H80000016&
            Index           =   1
            X1              =   30
            X2              =   7530
            Y1              =   795
            Y2              =   795
         End
         Begin VB.Line Line1 
            BorderColor     =   &H80000016&
            Index           =   0
            X1              =   60
            X2              =   7560
            Y1              =   3165
            Y2              =   3165
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
         Height          =   1305
         Left            =   -72525
         TabIndex        =   5
         Top             =   6135
         Width           =   5700
         Begin VB.CommandButton BotBusca 
            Height          =   330
            Left            =   4935
            Picture         =   "FrmParamFact.frx":0044
            Style           =   1  'Graphical
            TabIndex        =   8
            ToolTipText     =   "Buscar"
            Top             =   750
            Width           =   375
         End
         Begin VB.TextBox txtParam 
            BackColor       =   &H00FFFFFF&
            Height          =   315
            Index           =   15
            Left            =   1005
            TabIndex        =   7
            Top             =   285
            Width           =   4335
         End
         Begin VB.TextBox txtParam 
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
            Height          =   315
            Index           =   16
            Left            =   1785
            Locked          =   -1  'True
            TabIndex        =   6
            Top             =   795
            Width           =   690
         End
         Begin VB.Label lblParam 
            Caption         =   "Total Inmuebles"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   3
            Left            =   150
            TabIndex        =   33
            Top             =   817
            Width           =   1545
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
            Left            =   165
            TabIndex        =   9
            Top             =   315
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
         Height          =   1320
         Left            =   -74835
         TabIndex        =   2
         Top             =   6120
         Width           =   2205
         Begin VB.OptionButton OptBusca 
            Caption         =   "Por Código"
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
            Left            =   240
            TabIndex        =   4
            Top             =   465
            Value           =   -1  'True
            Width           =   1335
         End
         Begin VB.OptionButton OptBusca 
            Caption         =   "Por Nombre"
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
            Left            =   255
            TabIndex        =   3
            Top             =   855
            Width           =   1335
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
            Picture         =   "FrmParamFact.frx":05CE
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmParamFact.frx":0750
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmParamFact.frx":08D2
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmParamFact.frx":0A54
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmParamFact.frx":0BD6
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmParamFact.frx":0D58
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmParamFact.frx":0EDA
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmParamFact.frx":105C
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmParamFact.frx":11DE
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmParamFact.frx":1360
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmParamFact.frx":14E2
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmParamFact.frx":1664
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "FrmParamFact"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
    Dim rstFondos As New ADODB.Recordset
    Dim rstFact_Fondo As New ADODB.Recordset
    Dim ValoresIniciales()
    
    '---------------------------------------------------------------------------------------------
    Private Sub BotBusca_Click()    '
    '---------------------------------------------------------------        ------------------------------
    '
    If OptBusca(0) Then
        Call Busqueda("CodInm='" & txtParam(15) & "'")
    ElseIf OptBusca(1) Then
        Call Busqueda("Nombre LIKE *" & txtParam(15) & "*")
    End If
    '
    End Sub

    

Private Sub Check1_Click(Index As Integer)
If Index = 4 Then
    DataList1(0).Enabled = Check1(4)
    DataList1(1).Enabled = Check1(4)
End If
End Sub

    '---------------------------------------------------------------------------------------------
    Private Sub DataGrid1_DblClick()
    '---------------------------------------------------------------------------------------------
    '
    SSTab1.tab = 0
    txtNota = Obtener_Nota("\" & txtParam(13) & "\")
    Call Busca_Fondos
    '
    End Sub

Private Sub DataList1_DblClick(Index As Integer)
On Error Resume Next
Select Case Index
    Case 0  'agregar cuenta de fondo
        rstFact_Fondo.AddNew
        rstFact_Fondo!CodInm = txtParam(13)
        rstFact_Fondo!codGasto = DataList1(0).Text
        rstFact_Fondo.Update
        If Err.Number = 0 Then
            rstFact_Fondo.Requery
        Else
            rstFact_Fondo.CancelUpdate
        End If
    Case 1  'eliminar cuenta de fondo
        rstFact_Fondo.Delete
        rstFact_Fondo.Requery
End Select
End Sub

'    Private Sub Form_Resize()
'    'configura la presentación dela ficha
'    If WindowState <> vbMinimized Then
'        SSTab1.Left = (ScaleWidth / 2) - (SSTab1.Width / 2)
'        SSTab1.Top = (ScaleHeight / 2) - (SSTab1.Height / 2) + (Toolbar1.Height / 2)
'    End If
'    End Sub

    '---------------------------------------------------------------------------------------------
    Private Sub Form_Load() '
    '---------------------------------------------------------------------------------------------
    For I = 0 To 14: Set txtParam(I).DataSource = FrmAdmin.objRst
    Next
    Set txtParam(17).DataSource = FrmAdmin.objRst
    For I = 0 To 5: Set Check1(I).DataSource = FrmAdmin.objRst
    Next
    Set DataGrid1.DataSource = FrmAdmin.objRst
    With FrmAdmin.objRst
        .Find "CodInm='" & gcCodInm & "'"
        If .EOF Then
            .MoveFirst
            .Find "CodInm='" & gcCodInm & "'"
            If .EOF Then .MoveFirst
        End If
        
        Call RtnEstado(6, Toolbar1)
        If Not .EOF Or Not .BOF Then
            .Find "CodInm ='" & gcCodInm & "'"
        Else
            For I = 1 To 4
                Toolbar1.Buttons(I).Enabled = False
            Next
        End If
        txtParam(16) = .RecordCount
    End With
    txtNota = Obtener_Nota(gcUbica)
    Call Busca_Fondos
    With DataGrid1
        .HeadFont = LetraTitulo(LoadResString(527), 7.5, True)
        .Font = LetraTitulo(LoadResString(528), 9)
    End With
    ReDim ValoresIniciales(FrmAdmin.objRst.Fields.count - 1)
    '
    End Sub


    Private Sub Form_Unload(Cancel As Integer)
    'On Error Resume Next
    rstFondos.Close
    Set rstFondos = Nothing
    End Sub

    '---------------------------------------------------------------------------------------------
    Private Sub Toolbar1_ButtonClick(ByVal Button As Button)    '
    '---------------------------------------------------------------------------------------------
    '
    With FrmAdmin.objRst
    '
        Me.MousePointer = vbHourglass
        Select Case Button.Key
        '
            Case "First"    ' Primer Registro
        '   ---------------------
                If Not .EOF Or Not .BOF Then .MoveFirst
                txtNota = Obtener_Nota("\" & txtParam(13) & "\")
                Call Busca_Fondos
                
            Case "Previous"    ' Registro Anterior
        '   ---------------------
                If Not .EOF Or Not .BOF Then .MovePrevious
                If .BOF Then .MoveLast
                txtNota = Obtener_Nota("\" & txtParam(13) & "\")
                Call Busca_Fondos
                
            Case "Next" ' Registro Siguiente
        '   ---------------------
                If Not .EOF Or Not .BOF Then .MoveNext
                If .EOF Then .MoveFirst
                txtNota = Obtener_Nota("\" & txtParam(13) & "\")
                Call Busca_Fondos
                
            Case "End"  ' Ultimo Registro
        '   ---------------------
                If Not .EOF Or Not .BOF Then .MoveLast
                txtNota = Obtener_Nota("\" & txtParam(13) & "\")
                Call Busca_Fondos
                
'            Case "New"  ' Nuevo FrmInmueble.
'        '   ---------------------
'                Frame1.Enabled = True
'                .AddNew
'                Call RtnEstado(Button.Index, Toolbar1)
            
            Case "Save"    ' Actualizar.
        '   ---------------------
                If txtParam(0) = "" Then txtParam(0) = 0
                Call rtnGuardar_Nota
                !Honorarios = CCur(txtParam(0))
                If .EditMode = adEditInProgress Then
                    For I = 0 To .Fields.count - 1
                        If ValoresIniciales(I) <> .Fields(I) Then
                            
                            Call rtnBitacora("Actulizado el campo " & .Fields(I).Name & " Valor" _
                            & " Inicial: " & ValoresIniciales(I) & " Valor Actual: " & .Fields(I))
                            If UCase(.Fields(I).Name) = "HONORARIOS" Then
                            
                                cnnConexion.Execute "INSERT INTO Movimiento_Honorarios(CodInm,H" & _
                                "onoAnterior,Usuario,Fecha,Hora) VALUES ('" & txtParam(13) & "','" & _
                                ValoresIniciales(I) & "','" & gcUsuario & "',Date(),Time())"
                                
                            End If
                        End If
                    Next
                    
                End If
                !Freg = Date
                .Update
                'llevar un control del aumento de los honorarios
                
                Set DataGrid1.DataSource = FrmAdmin.objRst
                SSTab1.TabEnabled(1) = Not SSTab1.TabEnabled(1)
                Frame1.Enabled = False
                Call RtnEstado(Button.Index, Toolbar1)
                MsgBox " Registro Actualizado ... ", vbInformation, App.ProductName
                Call rtnBitacora("Actualizados los Parámetros de Fact. Inm: " & txtParam(13))
    
            Case "Find"     ' Buscar.
        '   ---------------------
                SSTab1.tab = 1
                FrmBusca.Enabled = True
                FrmBusca1.Enabled = True
                txtParam(15).SetFocus
            
            Case "Undo" ' Buscar.
        '   ---------------------
                Frame1.Enabled = False
                .CancelUpdate
                Set DataGrid1.DataSource = FrmAdmin.objRst
                SSTab1.TabEnabled(1) = Not SSTab1.TabEnabled(1)
                Call RtnEstado(Button.Index, Toolbar1)
               MsgBox " Registro Cancelado ... ", vbInformation, App.ProductName
            
'            Case "Delete"   'eliminar
'        '   ---------------------
'                MsgBox "Ipción No Disponible..", vbInformation, App.ProductName
        
            Case "Edit1"    'Editar
        '   ---------------------
                Frame1.Enabled = True
                Set DataGrid1.DataSource = Nothing
                SSTab1.TabEnabled(1) = Not SSTab1.TabEnabled(1)
                For I = 0 To .Fields.count - 1: ValoresIniciales(I) = .Fields(I)
                Next

                Call RtnEstado(Button.Index, Toolbar1)
            
            Case "Close"    'Cerrar formulario
        '   ---------------------
                Unload Me
                Set FrmParamFact = Nothing
                Exit Sub
                
            Case "Print"    ' Imprimir
        '   ---------------------
            mcTitulo = "Parámetros de Facturación"
            mcReport = "LisParFact.Rpt"
            mcOrdCod = "+{Inmueble.CodInm}"
            mcOrdAlfa = "+{Inmueble.Nombre}"
            mcCrit = ""
            'FrmReport.Show
        End Select
        '
        Me.MousePointer = vbDefault
    End With
    
    '
    End Sub

    
    '---------------------------------------------------------------------------------------------
    '
    '   Rutina:     Guardar_Nota
    '
    '   Guarda en la carpeta del inmueble un archivo "notas.txt", con la información
    '   que aparecerá en el aviso de cobro de los propietarios
    '---------------------------------------------------------------------------------------------
    Private Sub rtnGuardar_Nota()
    'variables locales
    Dim numFichero As Integer
    Dim strArchivo As String
    '
    numFichero = FreeFile
    strArchivo = "notas.txt"
    Open gcPath & "\" & txtParam(13) & "\" & strArchivo For Output As numFichero
    Print #numFichero, txtNota
    Close numFichero
    '
    End Sub
    
    '---------------------------------------------------------------------------------------------
    Private Sub Txt(CuadroTxt As TextBox, OrigenTxt As TextBox, intTecla%, strFormato$)
    '---------------------------------------------------------------------------------------------
    If intTecla = 46 Then intTecla = 44 'convierte el punto en coma
    Call Validacion(intTecla, "0123456789,") 'Permite entrada solo de números
    '
    If intTecla = 13 Then   'Presionó {enter} avanza el foco
        OrigenTxt = Format(OrigenTxt, strFormato)
        With CuadroTxt
            .SelStart = 0
            .SelLength = Len(CuadroTxt)
            .SetFocus
        End With
    End If
    '
    End Sub

    
    '---------------------------------------------------------------------------------------------
    Private Sub txtParam_KeyPress(Index As Integer, KeyAscii As Integer)
    '---------------------------------------------------------------------------------------------
    Select Case Index
    '
        Case 0, 3, 4, 6, 7, 8, 9, 10, 11
        '---------------------
            Call Txt(txtParam(Index + 1), txtParam(Index), KeyAscii, "#,##0.00")
            
        Case 1, 2, 5 'Meses Intereses
        '---------------------
            Call Txt(txtParam(Index + 1), txtParam(Index), KeyAscii, "0")
            
       Case 15  'Cuadro de Busqueda
       '---------------------
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
            If KeyAscii = 13 Then
                If OptBusca(0) Then
                    Call Busqueda("CodInm='" & txtParam(15) & "'")
                ElseIf OptBusca(1) Then
                    Call Busqueda("Nombre LIKE '*" & txtParam(15) & "*'")
                End If
            End If
            '
    End Select
    '
    End Sub

    '---------------------------------------------------------------------------------------------
    '
    '   Rutina:     Busqueda
    '
    '   Entrada:    sctrCriterio, cadena de busqueda
    '---------------------------------------------------------------------------------------------
    Private Sub Busqueda(strCriterio As String)
    '
    Dim boo2da As Boolean   'Variable local
    With FrmAdmin.objRst
        If Not .EOF Then .MoveNext
        .Find strCriterio
        '
        If .EOF Then
            .MoveFirst
            boo2da = True
            .Find strCriterio
            If .EOF And boo2da Then MsgBox "Inmueble No Encontrado...", vbInformation, _
            App.ProductName
        End If
        '
    End With
    '
    End Sub

    Private Sub Busca_Fondos()
    'variables locales
    Dim strBD$
    Dim strSQL$
    'si el recodset está abierto lo cierra
    If rstFondos.State = 1 Then rstFondos.Close
    strBD = cnnOLEDB + gcPath + FrmAdmin.objRst("Ubica") + "inm.mdb"
    'strBD = cnnOLEDB + mcDatos
    strSQL = "SELECT * FROM Tgastos WHERE Fondo=True;"
    rstFondos.Open strSQL, strBD, adOpenStatic, adLockReadOnly, adCmdText
    Set DataList1(0).RowSource = rstFondos
    
    '--------------------------
    If rstFact_Fondo.State = 1 Then rstFact_Fondo.Close
    rstFact_Fondo.Open "SELECT * FROM Fact_FondoInm WHERE CodInm='" & txtParam(13) & "'", _
    cnnConexion, adOpenKeyset, adLockOptimistic, adCmdText
    Set DataList1(1).RowSource = rstFact_Fondo
    '--------------------------
    End Sub
