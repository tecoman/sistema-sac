VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form FrmCheqDevuelto 
   Caption         =   "Cheques Devueltos"
   ClientHeight    =   3135
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   4680
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   3135
   ScaleWidth      =   4680
   Tag             =   "1"
   WindowState     =   2  'Maximized
   Begin ComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   480
      Left            =   0
      TabIndex        =   13
      Top             =   0
      Width           =   4680
      _ExtentX        =   8255
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
   Begin TabDlg.SSTab SSTab1 
      Height          =   5490
      Left            =   465
      TabIndex        =   14
      Top             =   735
      Width           =   11145
      _ExtentX        =   19659
      _ExtentY        =   9684
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
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
      TabCaption(0)   =   "Registro"
      TabPicture(0)   =   "FrmCheqDevuelto.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "fraCDev(1)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "fraCDev(0)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "mntFecha"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "Lista"
      TabPicture(1)   =   "FrmCheqDevuelto.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "fraCDev(3)"
      Tab(1).Control(1)=   "fraCDev(2)"
      Tab(1).Control(2)=   "dtgCheques"
      Tab(1).ControlCount=   3
      Begin MSComCtl2.MonthView mntFecha 
         Height          =   2310
         Left            =   3525
         TabIndex        =   7
         Top             =   1365
         Visible         =   0   'False
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   4075
         _Version        =   393216
         ForeColor       =   -2147483630
         BackColor       =   -2147483633
         Appearance      =   0
         MousePointer    =   99
         MouseIcon       =   "FrmCheqDevuelto.frx":0038
         ShowToday       =   0   'False
         StartOfWeek     =   72155138
         TitleBackColor  =   -2147483646
         TitleForeColor  =   -2147483639
         CurrentDate     =   37522
      End
      Begin MSDataGridLib.DataGrid dtgCheques 
         Height          =   3615
         Left            =   -74790
         TabIndex        =   23
         Top             =   405
         Width           =   9450
         _ExtentX        =   16669
         _ExtentY        =   6376
         _Version        =   393216
         AllowUpdate     =   -1  'True
         BackColor       =   16777215
         BorderStyle     =   0
         HeadLines       =   1
         RowHeight       =   15
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
         Caption         =   "CHEQUES DEVUELTOS"
         ColumnCount     =   6
         BeginProperty Column00 
            DataField       =   "Codigo"
            Caption         =   "Apto."
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   """Bs"" #,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   8202
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   "Fecha"
            Caption         =   "Fecha"
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
            DataField       =   "Monto"
            Caption         =   "Monto"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   "#,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   8202
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column03 
            DataField       =   "NumCheque"
            Caption         =   "Num. Cheque"
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
            DataField       =   "Banco"
            Caption         =   "Banco"
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
            DataField       =   "Recuperado"
            Caption         =   "Rec."
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   5
               Format          =   "0%"
               HaveTrueFalseNull=   1
               TrueValue       =   "X"
               FalseValue      =   ""
               NullValue       =   ""
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   8202
               SubFormatType   =   0
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            MarqueeStyle    =   1
            AllowFocus      =   0   'False
            AllowRowSizing  =   0   'False
            AllowSizing     =   0   'False
            BeginProperty Column00 
               ColumnWidth     =   659,906
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   1019,906
            EndProperty
            BeginProperty Column02 
               Alignment       =   1
            EndProperty
            BeginProperty Column03 
            EndProperty
            BeginProperty Column04 
               ColumnWidth     =   2160
            EndProperty
            BeginProperty Column05 
               Alignment       =   2
               ColumnWidth     =   720
            EndProperty
         EndProperty
      End
      Begin VB.Frame fraCDev 
         Enabled         =   0   'False
         Height          =   2070
         Index           =   0
         Left            =   200
         TabIndex        =   25
         Top             =   400
         Width           =   10725
         Begin VB.CheckBox Chk 
            Alignment       =   1  'Right Justify
            Caption         =   "Cobro Comisión:"
            DataField       =   "Comision"
            DataSource      =   "AdoDevuelto"
            Height          =   315
            Left            =   3675
            TabIndex        =   30
            Top             =   1485
            Width           =   2175
         End
         Begin VB.TextBox txtCdev 
            BackColor       =   &H00FFFFFF&
            CausesValidation=   0   'False
            DataField       =   "Motivo"
            BeginProperty DataFormat 
               Type            =   0
               Format          =   """Bs"" #,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   8202
               SubFormatType   =   0
            EndProperty
            DataSource      =   "AdoDevuelto"
            Height          =   1200
            Index           =   4
            Left            =   6105
            MaxLength       =   249
            MultiLine       =   -1  'True
            TabIndex        =   28
            Text            =   "FrmCheqDevuelto.frx":019A
            Top             =   690
            Width           =   4350
         End
         Begin VB.CommandButton cmdGasto 
            Height          =   255
            Left            =   5580
            Picture         =   "FrmCheqDevuelto.frx":019C
            Style           =   1  'Graphical
            TabIndex        =   27
            Top             =   705
            Width           =   270
         End
         Begin MSMask.MaskEdBox mskCdev 
            DataField       =   "Fecha"
            DataSource      =   "AdoDevuelto"
            Height          =   315
            Left            =   4545
            TabIndex        =   6
            Top             =   675
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   556
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   10
            Format          =   "DD/MM/YYYY"
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin VB.TextBox txtCdev 
            BackColor       =   &H00FFFFFF&
            DataField       =   "NumCheque"
            BeginProperty DataFormat 
               Type            =   0
               Format          =   """Bs"" #,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   8202
               SubFormatType   =   0
            EndProperty
            DataSource      =   "AdoDevuelto"
            Height          =   315
            Index           =   0
            Left            =   1290
            MaxLength       =   6
            TabIndex        =   4
            Text            =   " "
            Top             =   675
            Width           =   1725
         End
         Begin VB.TextBox txtCdev 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
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
            DataSource      =   "AdoDevuelto"
            Height          =   315
            Index           =   1
            Left            =   1290
            TabIndex        =   11
            Text            =   " "
            Top             =   1485
            Width           =   1725
         End
         Begin MSDataListLib.DataCombo cmbCDev 
            Bindings        =   "FrmCheqDevuelto.frx":02E6
            DataField       =   "Nombre"
            Height          =   315
            Index           =   1
            Left            =   2385
            TabIndex        =   2
            Top             =   285
            Width           =   3495
            _ExtentX        =   6165
            _ExtentY        =   556
            _Version        =   393216
            ListField       =   "Nombre"
            BoundColumn     =   "Nombre"
            Text            =   " "
         End
         Begin MSDataListLib.DataCombo cmbCDev 
            Bindings        =   "FrmCheqDevuelto.frx":0303
            DataField       =   "Banco"
            DataSource      =   "AdoDevuelto"
            Height          =   315
            Index           =   2
            Left            =   1290
            TabIndex        =   9
            Top             =   1080
            Width           =   4590
            _ExtentX        =   8096
            _ExtentY        =   556
            _Version        =   393216
            ListField       =   "NombreBanco"
            BoundColumn     =   ""
            Text            =   " "
         End
         Begin MSDataListLib.DataCombo cmbCDev 
            Bindings        =   "FrmCheqDevuelto.frx":0320
            DataField       =   "Codigo"
            DataSource      =   "AdoDevuelto"
            Height          =   315
            Index           =   0
            Left            =   1290
            TabIndex        =   1
            Top             =   285
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   556
            _Version        =   393216
            ListField       =   "Codigo"
            BoundColumn     =   "Codigo"
            Text            =   " "
         End
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            Caption         =   "Motivo"
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
            Left            =   6090
            TabIndex        =   29
            Top             =   420
            Width           =   4320
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            Caption         =   "Propietario:"
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
            Left            =   165
            TabIndex        =   0
            Top             =   300
            Width           =   1005
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            Caption         =   "Cheque N°:"
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
            Left            =   165
            TabIndex        =   3
            Top             =   690
            Width           =   1005
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            Caption         =   "Banco:"
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
            Left            =   165
            TabIndex        =   8
            Top             =   1095
            Width           =   1005
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            Caption         =   "Monto:"
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
            Left            =   165
            TabIndex        =   10
            Top             =   1500
            Width           =   1005
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            Caption         =   "Fecha Cheque:"
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
            Left            =   3075
            TabIndex        =   5
            Top             =   690
            Width           =   1320
         End
      End
      Begin VB.Frame fraCDev 
         Caption         =   "Cheques Recibos:"
         ForeColor       =   &H00400000&
         Height          =   1650
         Index           =   1
         Left            =   200
         TabIndex        =   24
         Top             =   2670
         Width           =   7300
         Begin MSFlexGridLib.MSFlexGrid FlexCheques 
            Height          =   1335
            Left            =   105
            TabIndex        =   12
            Tag             =   "1100|1100|2000|1100|1500|9000"
            Top             =   225
            Width           =   7095
            _ExtentX        =   12515
            _ExtentY        =   2355
            _Version        =   393216
            Cols            =   6
            FixedCols       =   0
            BackColorFixed  =   -2147483646
            ForeColorFixed  =   -2147483639
            BackColorSel    =   -2147483639
            GridColorFixed  =   12632256
            GridLinesFixed  =   1
            MousePointer    =   99
            FormatString    =   "Fecha Mov. |Cheque Nº |Banco |Fecha Cheq. |Monto |Detalle Op."
            MouseIcon       =   "FrmCheqDevuelto.frx":033D
         End
      End
      Begin VB.Frame fraCDev 
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
         Height          =   1275
         Index           =   2
         Left            =   -74835
         TabIndex        =   20
         Top             =   4035
         Width           =   1710
         Begin VB.OptionButton OptBusca 
            Caption         =   "Propietario"
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
            Left            =   135
            TabIndex        =   22
            Top             =   855
            Width           =   1500
         End
         Begin VB.OptionButton OptBusca 
            Caption         =   "Apartamento"
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
            TabIndex        =   21
            Top             =   465
            Value           =   -1  'True
            Width           =   1500
         End
      End
      Begin VB.Frame fraCDev 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1275
         Index           =   3
         Left            =   -72810
         TabIndex        =   15
         Top             =   3990
         Width           =   5520
         Begin VB.TextBox txtCdev 
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
            Index           =   3
            Left            =   1515
            Locked          =   -1  'True
            TabIndex        =   18
            Top             =   810
            Width           =   960
         End
         Begin VB.TextBox txtCdev 
            BackColor       =   &H00FFFFFF&
            Height          =   285
            Index           =   2
            Left            =   1005
            TabIndex        =   17
            Top             =   285
            Width           =   4335
         End
         Begin VB.CommandButton BotBusca 
            BackColor       =   &H80000004&
            Height          =   400
            Left            =   4695
            MaskColor       =   &H80000004&
            Style           =   1  'Graphical
            TabIndex        =   16
            ToolTipText     =   "Buscar"
            Top             =   700
            Width           =   600
         End
         Begin VB.Label Label3 
            Caption         =   "Total Registros:"
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
            Index           =   6
            Left            =   150
            TabIndex        =   26
            Top             =   832
            Width           =   1335
         End
         Begin VB.Label Label3 
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
            Index           =   1
            Left            =   165
            TabIndex        =   19
            Top             =   315
            Width           =   630
         End
      End
   End
   Begin MSAdodcLib.Adodc AdoDevuelto 
      Height          =   330
      Left            =   585
      Top             =   6315
      Visible         =   0   'False
      Width           =   2325
      _ExtentX        =   4101
      _ExtentY        =   582
      ConnectMode     =   16
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
      BackColor       =   0
      ForeColor       =   -2147483639
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
      Caption         =   "AdoDevuelto"
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
   Begin MSAdodcLib.Adodc AdoPropietario 
      Height          =   330
      Left            =   2850
      Top             =   6315
      Visible         =   0   'False
      Width           =   2430
      _ExtentX        =   4286
      _ExtentY        =   582
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
      BackColor       =   0
      ForeColor       =   -2147483624
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
      Caption         =   "AdoPropietario"
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
            Picture         =   "FrmCheqDevuelto.frx":049F
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmCheqDevuelto.frx":0621
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmCheqDevuelto.frx":07A3
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmCheqDevuelto.frx":0925
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmCheqDevuelto.frx":0AA7
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmCheqDevuelto.frx":0C29
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmCheqDevuelto.frx":0DAB
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmCheqDevuelto.frx":0F2D
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmCheqDevuelto.frx":10AF
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmCheqDevuelto.frx":1231
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmCheqDevuelto.frx":13B3
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmCheqDevuelto.frx":1535
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "FrmCheqDevuelto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
    'SAC.Modulo cheques Devueltos. Consulta y Edición de Cheques devueltos xcondominio
    Dim adoGrid As ADODB.Recordset 'Recordset de los registros del grid
    Dim AdoBancos As ADODB.Recordset
    Dim strCheque$
    '---------------------------------------------------------------------------------------------
    
    
    Private Sub BotBusca_Click()
    '
    With adoGrid
        .MoveFirst
        If Not .EOF Then
    '       Opciones de Busqueda
            If OptBusca(0).Value = True Then    'Por Codigo de Apto
                .Find "Codigo LIKE '*" & txtCdev(2) & "*'"
            ElseIf OptBusca(1).Value = True Then    'Por Nombre
                With AdoPropietario.Recordset
                    .MoveFirst
                    .Find "Nombre Like '*" & txtCdev(2) & "*'"
                    If .EOF Then MsgBox "Propietario No Registrado de este condominio", _
                            vbInformation, App.ProductName
                End With
                .Find "Codigo ='" & AdoPropietario.Recordset!Codigo & "'"
            End If
            If .EOF Then
                MsgBox "Propietario '" & txtCdev(2) & "'" & vbCrLf _
                & "No Tiene Cheque Devuelto", vbInformation, App.ProductName
            Else
                With AdoDevuelto.Recordset
                    .MoveFirst
                    .Find "Codigo = '" & adoGrid!Codigo & "'"
                    Call rtnAsignaCampo("codigo", !Codigo, 0)
                    Call RtnCheqMovCaja(!Codigo)
                End With
            End If
        End If
    End With
    End Sub

    '---------------------------------------------------------------------------------------------
    Private Sub cmbCDev_Click(Index As Integer, Area As Integer)    '
    '---------------------------------------------------------------------------------------------
    '
    If Area = 2 Then
    '
        Select Case Index
            Case 0  'código de propietario
    '       ---------------------------
                Call rtnAsignaCampo("Codigo", cmbCDev(0), 1)
                Call RtnCheqMovCaja(cmbCDev(0))
    
            Case 1  'nombre de propietario
    '       ---------------------------
                Call rtnAsignaCampo("Nombre", cmbCDev(1), 1)
                Call RtnCheqMovCaja(cmbCDev(0))
    '
        End Select
    '
    End If
    '
    End Sub

    '---------------------------------------------------------------------------------------------
    Private Sub cmbCDev_KeyPress(Index As Integer, KeyAscii As Integer) '
    '---------------------------------------------------------------------------------------------
    '
    KeyAscii = Asc(UCase(Chr(KeyAscii)))    'Convierte todo a mayúsculas
    '
    If KeyAscii = 13 Then
    '
        Select Case Index
            Case 0  'código de propietario
    '       ---------------------------
                
                If cmbCDev(0) <> "" And Len(cmbCDev(0)) >= 4 Then _
                    Call rtnAsignaCampo("Codigo", cmbCDev(0), 1): Call RtnCheqMovCaja(cmbCDev(0))
                If cmbCDev(0) = "" Then cmbCDev(1).SetFocus
                
            Case 1  'nombre de propietario
    '       ---------------------------
                If cmbCDev(1) <> "" Then
                    Call rtnAsignaCampo("Nombre", cmbCDev(1), 1)
                    Call RtnCheqMovCaja(cmbCDev(0))
                End If
                If cmbCDev(1) = "" Then cmbCDev(0).SetFocus
    '
            Case 2  'nombre de banco
    '       ---------------------------
                txtCdev(1).SetFocus
                
        End Select
    '
    End If
    '
    End Sub
    
    '---------------------------------------------------------------------------------------------
    Private Sub cmdGasto_Click()    '
    '---------------------------------------------------------------------------------------------
    '
    With mskCdev
        .SelStart = 0
        .SelLength = Len(.Text) + 2
        .SetFocus
    End With
    mntFecha.Visible = Not mntFecha.Visible
    '
    End Sub

    '---------------------------------------------------------------------------------------------
    Private Sub FlexCheques_DblClick()  '
    '---------------------------------------------------------------------------------------------
    '
    'El cheue seleccionado se agrega a los controles correpondientes para ser guardado
    '
    With FlexCheques
        txtCdev(0) = Right(.TextMatrix(.RowSel, 1), 6)
        cmbCDev(2) = .TextMatrix(.RowSel, 2)
         txtCdev(1) = .TextMatrix(.RowSel, 4)
         mskCdev = .TextMatrix(.RowSel, 3)
    End With
    '
    End Sub

    '---------------------------------------------------------------------------------------------
    Private Sub Form_Load() '
    '---------------------------------------------------------------------------------------------
    BotBusca.Picture = LoadResPicture("Buscar", vbResIcon)
    Set adoGrid = New ADODB.Recordset
    Set AdoBancos = New ADODB.Recordset
    '---------------------------------------------------------------------------------------------
    With AdoPropietario
        .ConnectionString = cnnOLEDB + mcDatos
        .RecordSource = "SELECT * FROM Propietarios WHERE Codigo<>'U*'"
        .Refresh
    End With
    '---------------------------------------------------------------------------------------------
    With AdoDevuelto
        .ConnectionString = cnnOLEDB + mcDatos
        .CursorLocation = adUseClient
        .CursorType = adOpenKeyset
        .CommandType = adCmdText
        .LockType = adLockOptimistic
        .RecordSource = "SELECT * FROM ChequeDevuelto ORDER BY Fecha,Hora"
        .Refresh
    End With
'    Set txtCdev(4).DataSource = AdoDevuelto
'    txtCdev(4).DataField = "Motivo"
    
    '
    'Call RtnConfigGrid
    Call centra_titulo(FlexCheques, True)
    If Not AdoDevuelto.Recordset.EOF Then
        Call rtnAsignaCampo("Codigo", cmbCDev(0), 0)
        Call RtnCheqMovCaja(cmbCDev(0))
        txtCdev(3) = AdoDevuelto.Recordset.RecordCount
    End If
    '---------------------------------------------------------------------------------------------
    adoGrid.CursorLocation = adUseClient
    adoGrid.Open "SELECT * FROM ChequeDevuelto ORDER BY Fecha,Hora", _
    cnnOLEDB + mcDatos, adOpenKeyset, adLockOptimistic, adCmdText
    Set dtgCheques.DataSource = adoGrid
    dtgCheques.Refresh
    '---------------------------------------------------------------------------------------------
    AdoBancos.Open "SELECT * FROM BANCOS ORDER BY NombreBanco", _
        cnnConexion, adOpenKeyset, adLockOptimistic
    Set cmbCDev(2).RowSource = AdoBancos
    mntFecha.Value = Date
    Call RtnEstado(6, Toolbar1)
    '
    End Sub

    '04/09/2002-----------------------------------------------------------------------------------
    Private Sub Form_Resize()   '
    '---------------------------------------------------------------------------------------------
    'Configura la presentación de todos los controles en pantalla
    On Error Resume Next
    Dim intEntero%
    
    If FrmCheqDevuelto.WindowState <> vbMinimized Then
    '   -----------------------------------------------
        With SSTab1     'Configura el SStab
            .Top = Toolbar1.Height + 200
            .Width = FrmCheqDevuelto.Width - 800
            .Height = FrmCheqDevuelto.Height - (.Top * 2)
        End With
    '   ---------------------------------------------------
        intEntero = IIf(SSTab1.tab = 0, 0, 1)
    '
        SSTab1.tab = 0
        With fraCDev(0) 'Marco Principal
            .Width = SSTab1.Width - (.Left * 2)
        End With
    '   ---------------------------------------------------
        With fraCDev(1) 'Marco Secundario
            .Top = fraCDev(0).Top + fraCDev(0).Height + 200
            .Height = SSTab1.Height - .Top - 100
            .Width = fraCDev(0).Width
        End With
    '   ---------------------------------------------------
        With FlexCheques    'FlexCheques
            .Height = fraCDev(1).Height - .Top - 100
            .Width = fraCDev(1).Width - (.Left * 2)
            If FrmCheqDevuelto.WindowState = vbNormal Then
                .ColWidth(5) = 0
            Else
                .ColWidth(5) = .Width - (.ColWidth(0) + .ColWidth(1) + .ColWidth(2) + .ColWidth(3) + .ColWidth(4)) - 300
            End If
        End With
    '   ----------------------------------------------------
        SSTab1.tab = 1
        dtgCheques.Width = SSTab1.Width - (dtgCheques.Left * 2)
        dtgCheques.Height = SSTab1.Height - fraCDev(2).Height - dtgCheques.Top - 200
        fraCDev(2).Top = dtgCheques.Height + dtgCheques.Top
        fraCDev(3).Top = fraCDev(2).Top
        SSTab1.tab = intEntero
    '
    End If
    End Sub
    
    '---------------------------------------------------------------------------------------------
    Private Sub mntFecha_DateClick(ByVal DateClicked As Date)   '
    '---------------------------------------------------------------------------------------------
    '
    mskCdev = DateClicked
    mntFecha.Visible = False
    cmbCDev(2).SetFocus
    '
    End Sub

    '---------------------------------------------------------------------------------------------
    Private Sub mntFecha_KeyPress(KeyAscii As Integer)  '
    '---------------------------------------------------------------------------------------------
    If KeyAscii = 13 Then
        mskCdev = mntFecha.Value
        mntFecha.Visible = False
        cmbCDev(2).SetFocus
    End If
    End Sub

    '---------------------------------------------------------------------------------------------
    Private Sub mntFecha_KeyUp(KeyCode As Integer, Shift As Integer)    '
    '---------------------------------------------------------------------------------------------
    '
    If KeyCode = 27 Then
        mntFecha.Visible = False
        With mskCdev
            .SelStart = 0
            .SelLength = 20
            .SetFocus
        End With
    End If
    '
    End Sub

    '---------------------------------------------------------------------------------------------
    Private Sub mskCdev_KeyDown(KeyCode As Integer, Shift As Integer)   '
    '---------------------------------------------------------------------------------------------
    '
    If Shift = 4 And KeyCode = 40 Then
        mntFecha.Visible = True
        mntFecha.SetFocus
    End If
    '
    End Sub

    '---------------------------------------------------------------------------------------------
    Private Sub mskCdev_KeyPress(KeyAscii As Integer)   '
    '---------------------------------------------------------------------------------------------
    '
    If KeyAscii = 13 Then
        If mskCdev = "" Then
            mskCdev = Date
            mskCdev.SetFocus
        Else
            cmbCDev(2).SetFocus
        End If
    End If
    '
    End Sub

    '---------------------------------------------------------------------------------------------
    Private Sub Toolbar1_ButtonClick(ByVal Button As ComctlLib.Button)
    '---------------------------------------------------------------------------------------------
    '
    With AdoDevuelto.Recordset
    '
        Select Case Button.Index
    '
            Case 1  'Primer Registro
    '       ---------------------------
                If Not .EOF Then
                    .MoveFirst
                    Call rtnAsignaCampo("codigo", !Codigo, 0)
                    Call RtnCheqMovCaja(!Codigo)
                    
                Else
                    Exit Sub
                End If
    '
            Case 2  'Registro Previo
    '       ---------------------------
                If Not .EOF Then
                    .MovePrevious
                    If .BOF Then
                        .MoveFirst
                    End If
                    Call rtnAsignaCampo("codigo", !Codigo, 0)
                    Call RtnCheqMovCaja(!Codigo)
                Else
                    Exit Sub
                End If
    '
            Case 3  'Siguiente Registro
    '       ---------------------------
                If Not .EOF Then
                    .MoveNext
                    If .EOF Then
                        .MoveFirst
                    End If
                    Call rtnAsignaCampo("codigo", !Codigo, 0)
                    Call RtnCheqMovCaja(!Codigo)
                Else
                    Exit Sub
                End If
    '
            Case 4  'Último Registro
    '       ---------------------------
                If Not .EOF Then
                    .MoveLast
                    Call rtnAsignaCampo("codigo", !Codigo, 0)
                    Call RtnCheqMovCaja(!Codigo)
                Else
                    Exit Sub
                End If
    '
            Case 5  'Nuevo Registro
    '       ---------------------------
                .AddNew
                fraCDev(0).Enabled = True
                cmbCDev(1) = ""
                cmbCDev(0).SetFocus
                For I = 0 To FlexCheques.Cols - 1
                    FlexCheques.TextArray(I + 6) = ""
                Next
                FlexCheques.Rows = 2
                chk.Value = vbUnchecked
                txtCdev(4) = ""
                Call RtnEstado(5, Toolbar1)
    '
            Case 6  'Actualizar Registro
    '        --------------------------
                Dim datMax As Date, strMsg$
                Dim curCdev@, mlNew As Boolean
                
                MousePointer = vbHourglass
                If .EditMode = adEditAdd Then
                    mlNew = True
                    strMsg = "Registrar "
                Else
                    strMsg = "Actualizar "
                End If
                strMsg = strMsg & "Cheque Devuelto Inm:" & gcCodInm & "/" & cmbCDev(0) _
                & "/" & txtCdev(0)
                Call RtnEstado(6, Toolbar1)
                fraCDev(0).Enabled = False
                mskCdev.PromptInclude = True
                !Usuario = gcUsuario
                !Hora = Time()
                !Freg = Date
                If chk.Value = vbUnchecked Then
                
                    If Respuesta("Este cheque devuelto no se le cobrará comisión." & _
                    vbCrLf & "¿Desea que se le cobre la comisión a este cliente?") Then
                    
                    chk.Value = vbChecked
                
                    End If
                End If
                .Update
                
                If mlNew Then   'si está adicionando un nuevo registro
                    'genera el cheque devuelto en sac.mdb
                    cnnConexion.Execute "INSERT INTO ChequeDevuelto (CodInm,Apto,Banco,Numero," _
                    & "Fecha,Monto,Motivo,Comision,Usuario,Freg) SELECT " & gcCodInm & ",'" & !Codigo & _
                    "','" & !Banco & "'," & !NumCheque & ",'" & !fecha & "','" & CCur(!Monto) & _
                    "','" & !Motivo & "'," & chk.Value & ",'" & gcUsuario & "',DATE()"
                    '----------------------------
                    'genera una factura pendiente en la deuda del propietario
                    curCdev = CCur(txtCdev(1))
                    txtCdev(3) = .RecordCount
                    datMax = MaxPer(cmbCDev(0))
                    cnnConexion.Execute "INSERT INTO Factura (Fact, Periodo, CodProp,Facturado," _
                    & "Pagado, Saldo, Freg, Usuario, Fecha,FechaFactura) IN'" & mcDatos & _
                    "' VALUES ('CHD." & txtCdev(0) & "','" & datMax & "','" & cmbCDev(0) & "','" _
                    & curCdev & "',0,'" & curCdev & "', Date()" & ",'" & gcUsuario & _
                    "',Left(Time(),10),Date())"
                    '----------------------------
                    'Aumenta la deuda del propietario
                    cnnConexion.Execute "UPDATE Propietarios IN '" & mcDatos & "' SET Deuda = d" _
                    & "euda +'" & curCdev & "' WHERE codigo ='" & cmbCDev(0) & "'"
                    '----------------------------
                    'Aumeta la deuda del inmuelbe
                    cnnConexion.Execute "UPDATE Inmueble SET Deuda = Deuda + '" & curCdev _
                    & "' WHERE CodInm = '" & gcCodInm & "'"
                    '
                Else
                    'Actualiza el registro en sac.mdb
                    cnnConexion.Execute "UPDATE ChequeDevuelto SET CodInm='" & gcCodInm & "',Apt" _
                    & "o='" & !Codigo & "',Banco='" & !Banco & "',Numero='" & !NumCheque & "',F" _
                    & "echa='" & !fecha & "',Monto='" & !Monto & "',Motivo='" & !Motivo & "',Usua" _
                    & "rio='" & gcUsuario & "' WHERE CodInm & Apto & Numero='" & strCheque & "';"
                End If
                mskCdev.PromptInclude = False
                MsgBox "Registro actualizado..."
                .Requery
                Call rtnBitacora(strMsg)
                MousePointer = vbDefault
                                        
            Case 7  'Buscar Registro
    '       ---------------------------
                'SSTab1.Tab = 1
                frmBuscaCheque.Show
            
            Case 8  'Cancelar Registro
    '       ---------------------------
                Call RtnEstado(6, Toolbar1)
                mskCdev.PromptInclude = True
                .CancelUpdate
                mskCdev.PromptInclude = False
                .Requery
                If Not .EOF Then
                    Call rtnAsignaCampo("codigo", !Codigo, 0)
                    Call RtnCheqMovCaja(!Codigo)
                Else
                    cmbCDev(1) = ""
                End If
                fraCDev(0).Enabled = False
                
            Case 9  'Eliminar Registro
    '       ---------------------------
                MsgBox "Estimado Usuario: " & gcUsuario & vbCrLf & "Usted no tiene los permisos" _
                & " necesarios para ejecutar esta opción," & vbCrLf & "Consulte con el administ" _
                & "rador  del sistema", vbExclamation + vbOKOnly
                Exit Sub
                Dim Confirma As Integer
                Confirma = MsgBox("Confirma eliminar el registro actual ?", vbOKCancel, "Eliminar Registro")
                If Confirma = vbOK Then
                    .Delete
                    .MoveNext
                    If .EOF Then
                        cmbCDev(1) = ""
                    Else
                        .MoveFirst
                    End If
                End If
    
            Case 10 'Editar
    '       ---------------------------
                fraCDev(0).Enabled = True
                strCheque = gcCodInm & cmbCDev(0) & txtCdev(0)
                Call RtnEstado(5, Toolbar1)
            
            Case 12 'Descargar Formulario
    '       ---------------------------
                Unload Me
                Set FrmCheqDevuelto = Nothing
                
            Case 11 'Imprimir Registro
    '       ---------------------------
                mcTitulo = "Listado de Cheques Devueltos"
                mcReport = "LisCheqDev.Rpt"
                mcOrdCod = "+{ChequeDevuelto.CodInm}"
                mcOrdAlfa = "+{ChequeDevuelto.Banco}"
                mcCrit = "{ChequeDevuelto.Freg}>={@Desde} AND {ChequeDevuelto.Freg}<={@Hasta}"
                FrmReport.Frame1.Visible = True
                FrmReport.Show
                                
        End Select
        '
     End With
    
    End Sub

    '---------------------------------------------------------------------------------------------
    Private Sub rtnAsignaCampo(campo$, valor$, bySi As Byte)
    '---------------------------------------------------------------------------------------------
    'Variable local
    Dim booSINO As Boolean
    '
    booSINO = BuscaProp(campo, valor, AdoPropietario)
    '
    If booSINO = True Then
        cmbCDev(0) = AdoPropietario.Recordset!Codigo
        cmbCDev(1) = AdoPropietario.Recordset!Nombre
        If bySi = 1 Then
            If txtCdev(0) = "" Then
                txtCdev(0).SetFocus
            ElseIf cmbCDev(2) = "" Then
                cmbCDev(2).SetFocus
            ElseIf txtCdev(1) = "" Then
                txtCdev(1).SetFocus
            End If
        End If
    Else
        MsgBox "No Tengo Registrado ese Propietario '" & valor & "'"
        AdoPropietario.Refresh
    End If
    '
    End Sub


    '---------------------------------------------------------------------------------------------
    Private Sub txtCdev_KeyPress(Index As Integer, KeyAscii As Integer)
    '---------------------------------------------------------------------------------------------
    '
    KeyAscii = Asc(UCase(Chr(KeyAscii)))    'Convierte todo a mayúsculas
    If Index = 1 Then
        If KeyAscii = 46 Then KeyAscii = 44 'CONVIERTE PUNTO(.) EN COMA(,)
    End If
    If KeyAscii = 13 Then
        Select Case Index
            Case 0  'N° del cheque
    '       ---------------------------
                mskCdev.SetFocus
                
            Case 1  'Monto
    '       ---------------------------
            If txtCdev(0) = "" Then txtCdev(0).SetFocus: Exit Sub
            If cmbCDev(2) = "" Then cmbCDev(2).SetFocus
            txtCdev(1) = Format(txtCdev(1), "#,##0.00")
            
            
        End Select
            
            
    '
    End If
    '
    End Sub

    '---------------------------------------------------------------------------------------------
    Private Sub RtnCheqMovCaja(StrApto As String)   '
    '---------------------------------------------------------------------------------------------
    '
    Dim strSQL As String    'Cadena SQL
    Dim adoCheqMovCaja As ADODB.Recordset   'Cheques recibidos en caja de un propietario
    Dim I, j As Integer
    Set adoCheqMovCaja = New ADODB.Recordset
    '---------------------------------------------------------------------------------------------
    strSQL = "SELECT FechaMovimientoCaja, NumDocumentoMovimientoCaja, " _
        & "BancoDocumentoMovimientoCaja, FechaChequeMovimientoCaja, MontoCheque, " _
        & "CuentaMovimientoCaja & ' ' & DescripcionMovimientoCaja FROM MovimientoCaja WHERE " _
        & "AptoMovimientoCaja='" & StrApto & "' AND InmuebleMovimientoCaja='" & gcCodInm _
        & "' AND Fpago='CHEQUE' UNION SELECT FechaMovimientoCaja, NumDocumentoMovimientoCaja1, " _
        & "BancoDocumentoMovimientoCaja1, FechaChequeMovimientoCaja1, MontoCheque1, " _
        & "CuentaMovimientoCaja & ' ' & DescripcionMovimientoCaja FROM MovimientoCaja WHERE " _
        & "AptoMovimientoCaja='" & StrApto & "' AND InmuebleMovimientoCaja='" & gcCodInm _
        & "'AND Fpago1='CHEQUE' UNION SELECT FechaMovimientoCaja, NumDocumentoMovimientoCaja2, " _
        & "BancoDocumentoMovimientoCaja2, FechaChequeMovimientoCaja2, MontoCheque2, " _
        & "CuentaMovimientoCaja & ' ' & DescripcionMovimientoCaja FROM MovimientoCaja WHERE " _
        & "AptoMovimientoCaja='" & StrApto & "' AND InmuebleMovimientoCaja='" & gcCodInm _
        & "'AND Fpago2='CHEQUE' ORDER BY FechaMovimientoCaja"
    '---------------------------------------------------------------------------------------------
    adoCheqMovCaja.Open strSQL, cnnConexion, adOpenStatic, adLockReadOnly
    If Not adoCheqMovCaja.EOF Then
    '
        adoCheqMovCaja.MoveFirst
        I = 0
        With FlexCheques
            
            .Rows = adoCheqMovCaja.RecordCount + 1
            Do Until adoCheqMovCaja.EOF
                I = I + 1
                For j = 0 To 5
                    .TextMatrix(I, j) = IIf(j = 4, Format(adoCheqMovCaja.Fields(j), "#,##0.00"), _
                        adoCheqMovCaja.Fields(j))
                Next
                adoCheqMovCaja.MoveNext
            Loop
        End With
    '
    Else
        
        For I = 0 To FlexCheques.Cols - 1
            FlexCheques.TextArray(I + 6) = ""
        Next
        FlexCheques.Rows = 2
    End If
    adoCheqMovCaja.Close: Set adoCheqMovCaja = Nothing
    End Sub

        
    'Funcion que calcula el ultimo perido facturado para determinado propietario------------------
    Private Function MaxPer(apto As String) '
    '---------------------------------------------------------------------------------------------
    '
    With AdoPropietario
        .RecordSource = "SELECT DateAdd('m',1,MAX(Periodo)) as V FROM Factura WHERE CodProp='" _
        & apto & "';"
        .Refresh
        MaxPer = IIf(IsNull(.Recordset.Fields(0)), DateAdd("d", -Day(Date) + 1, Date), _
        .Recordset.Fields(0))
        .RecordSource = "SELECT * FROM Propietarios WHERE Codigo<>'U*'"
        .Refresh
    End With
    '
    End Function
