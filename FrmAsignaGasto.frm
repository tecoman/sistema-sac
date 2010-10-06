VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form FrmAsignaGasto 
   AutoRedraw      =   -1  'True
   Caption         =   "Asignacion del Gasto"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   Tag             =   "1"
   WindowState     =   2  'Maximized
   Begin ComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   480
      Left            =   0
      TabIndex        =   22
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
            Enabled         =   0   'False
            Key             =   "First"
            Object.ToolTipText     =   "Primer Registro"
            Object.Tag             =   ""
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Enabled         =   0   'False
            Key             =   "Previous"
            Object.ToolTipText     =   "Registro Anterior"
            Object.Tag             =   ""
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Enabled         =   0   'False
            Key             =   "Next"
            Object.ToolTipText     =   "Siguiente Registro"
            Object.Tag             =   ""
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Enabled         =   0   'False
            Key             =   "End"
            Object.ToolTipText     =   "Último Registro"
            Object.Tag             =   ""
            ImageIndex      =   4
         EndProperty
         BeginProperty Button5 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Enabled         =   0   'False
            Key             =   "New"
            Object.ToolTipText     =   "Nuevo Registro"
            Object.Tag             =   ""
            ImageIndex      =   5
         EndProperty
         BeginProperty Button6 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Enabled         =   0   'False
            Key             =   "Save"
            Object.ToolTipText     =   "Guardar Registro"
            Object.Tag             =   ""
            ImageIndex      =   6
         EndProperty
         BeginProperty Button7 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Enabled         =   0   'False
            Key             =   "Find"
            Object.ToolTipText     =   "Buscar Registro"
            Object.Tag             =   ""
            ImageIndex      =   7
         EndProperty
         BeginProperty Button8 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Enabled         =   0   'False
            Key             =   "Undo"
            Object.ToolTipText     =   "Cancelar Registro"
            Object.Tag             =   ""
            ImageIndex      =   8
         EndProperty
         BeginProperty Button9 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Enabled         =   0   'False
            Key             =   "Delete"
            Object.ToolTipText     =   "Eliminar Registro"
            Object.Tag             =   ""
            ImageIndex      =   9
         EndProperty
         BeginProperty Button10 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Enabled         =   0   'False
            Key             =   "Edit1"
            Object.ToolTipText     =   "Editar Registro"
            Object.Tag             =   ""
            ImageIndex      =   10
         EndProperty
         BeginProperty Button11 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Enabled         =   0   'False
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
   Begin MSAdodcLib.Adodc AdoCuentas 
      Height          =   345
      Left            =   3195
      Top             =   7440
      Visible         =   0   'False
      Width           =   2670
      _ExtentX        =   4710
      _ExtentY        =   609
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
      Connect         =   "Provider=Microsoft.Jet.OLEDB.3.51;Persist Security Info=False;Data Source=f:\Sac\Datos\2503\Inm.mdb;Mode=ReadWrite"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.3.51;Persist Security Info=False;Data Source=f:\Sac\Datos\2503\Inm.mdb;Mode=ReadWrite"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "Tgastos"
      Caption         =   "AdoCuentas"
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
   Begin MSAdodcLib.Adodc AdoGastos 
      Height          =   345
      Left            =   420
      Top             =   7425
      Visible         =   0   'False
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   609
      ConnectMode     =   0
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
      Connect         =   "Provider=Microsoft.Jet.OLEDB.3.51;Persist Security Info=False;Data Source=f:\Sac\Datos\2503\Inm.mdb"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.3.51;Persist Security Info=False;Data Source=f:\Sac\Datos\2503\Inm.mdb"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "select * from tgastos "
      Caption         =   "AdoGastos"
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
      Height          =   7230
      Left            =   120
      TabIndex        =   15
      Top             =   600
      Width           =   11655
      _ExtentX        =   20558
      _ExtentY        =   12753
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   4
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
      TabPicture(0)   =   "FrmAsignaGasto.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "fraAgasto(0)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "&Lista"
      TabPicture(1)   =   "FrmAsignaGasto.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "DataGrid1"
      Tab(1).Control(1)=   "fraAgasto(6)"
      Tab(1).Control(2)=   "fraAgasto(5)"
      Tab(1).Control(3)=   "fraAgasto(3)"
      Tab(1).ControlCount=   4
      Begin MSDataGridLib.DataGrid DataGrid1 
         Height          =   4725
         Left            =   -74790
         TabIndex        =   17
         Top             =   585
         Width           =   11295
         _ExtentX        =   19923
         _ExtentY        =   8334
         _Version        =   393216
         AllowUpdate     =   0   'False
         AllowArrows     =   -1  'True
         HeadLines       =   2
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
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "LISTA DE FACTURAS PENDIENTES"
         ColumnCount     =   7
         BeginProperty Column00 
            DataField       =   "Tipo"
            Caption         =   "T/Doc."
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
            DataField       =   "NDOC"
            Caption         =   "Documento"
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
            DataField       =   "FACT"
            Caption         =   "Factura"
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
            DataField       =   "CODPROV"
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
         BeginProperty Column05 
            DataField       =   "Total"
            Caption         =   "Monto"
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
            DataField       =   "FRecep"
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
         SplitCount      =   1
         BeginProperty Split0 
            AllowRowSizing  =   0   'False
            ScrollGroup     =   0
            BeginProperty Column00 
               Alignment       =   2
               ColumnWidth     =   750,047
            EndProperty
            BeginProperty Column01 
               Alignment       =   2
               ColumnWidth     =   1049,953
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   734,74
            EndProperty
            BeginProperty Column03 
               Alignment       =   2
               ColumnWidth     =   870,236
            EndProperty
            BeginProperty Column04 
               ColumnWidth     =   4889,764
            EndProperty
            BeginProperty Column05 
               Alignment       =   1
            EndProperty
            BeginProperty Column06 
               Alignment       =   2
               ColumnWidth     =   1094,74
            EndProperty
         EndProperty
      End
      Begin VB.Frame fraAgasto 
         Caption         =   "Buscar:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1065
         Index           =   6
         Left            =   -69795
         TabIndex        =   35
         Top             =   5670
         Width           =   6195
         Begin VB.OptionButton OptTipoFactura 
            Caption         =   "Nº Cheque"
            Height          =   280
            Index           =   9
            Left            =   3195
            MaskColor       =   &H008080FF&
            TabIndex        =   63
            Tag             =   "Total"
            Top             =   705
            Visible         =   0   'False
            Width           =   1110
         End
         Begin VB.OptionButton OptTipoFactura 
            Caption         =   "Descripción"
            Height          =   280
            Index           =   6
            Left            =   1680
            TabIndex        =   41
            Tag             =   "Detalle"
            Top             =   705
            Width           =   1245
         End
         Begin VB.OptionButton OptTipoFactura 
            Caption         =   "N° Documento"
            Height          =   280
            Index           =   7
            Left            =   225
            TabIndex        =   40
            Tag             =   "Ndoc"
            Top             =   705
            Width           =   1410
         End
         Begin VB.TextBox TxtAGasto 
            Height          =   315
            Index           =   5
            Left            =   3630
            TabIndex        =   39
            Top             =   255
            Width           =   2445
         End
         Begin VB.CommandButton cmdHelp 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   350
            Index           =   3
            Left            =   4950
            Style           =   1  'Graphical
            TabIndex        =   38
            Top             =   650
            Width           =   1095
         End
         Begin VB.OptionButton OptTipoFactura 
            Caption         =   "Monto"
            Height          =   280
            Index           =   5
            Left            =   1695
            MaskColor       =   &H008080FF&
            TabIndex        =   37
            Tag             =   "Total"
            Top             =   272
            Width           =   855
         End
         Begin VB.OptionButton OptTipoFactura 
            Caption         =   "N° Factura"
            Height          =   280
            Index           =   4
            Left            =   225
            TabIndex        =   36
            Tag             =   "Fact"
            Top             =   272
            Value           =   -1  'True
            Width           =   1110
         End
      End
      Begin VB.Frame fraAgasto 
         Caption         =   "Ordenar Por"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1065
         Index           =   5
         Left            =   -74895
         TabIndex        =   32
         Top             =   5670
         Width           =   1980
         Begin VB.OptionButton OptTipoFactura 
            Caption         =   "Proveedor"
            Height          =   195
            Index           =   3
            Left            =   300
            TabIndex        =   34
            Tag             =   "CodProv"
            Top             =   360
            Value           =   -1  'True
            Width           =   1245
         End
         Begin VB.OptionButton OptTipoFactura 
            Caption         =   "Fecha Factura"
            Height          =   195
            Index           =   2
            Left            =   285
            TabIndex        =   33
            Tag             =   "Frecep"
            Top             =   675
            Width           =   1410
         End
      End
      Begin VB.Frame fraAgasto 
         Caption         =   "Ver Facturas"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1065
         Index           =   3
         Left            =   -72795
         TabIndex        =   18
         Top             =   5670
         Width           =   2865
         Begin VB.OptionButton OptTipoFactura 
            Caption         =   "Pa&gadas"
            Height          =   195
            Index           =   8
            Left            =   1605
            TabIndex        =   45
            Tag             =   "PAGADO"
            ToolTipText     =   "LISTA DE FACTURAS CANCELADAS"
            Top             =   675
            Width           =   1110
         End
         Begin VB.OptionButton OptTipoFactura 
            Caption         =   "&Asignadas"
            Height          =   195
            Index           =   1
            Left            =   285
            TabIndex        =   20
            Tag             =   "ASIGNADO"
            ToolTipText     =   "LISTA DE FACTURAS ASIGNADAS"
            Top             =   675
            Width           =   1110
         End
         Begin VB.OptionButton OptTipoFactura 
            Caption         =   "&Pendientes"
            Height          =   195
            Index           =   0
            Left            =   300
            TabIndex        =   19
            Tag             =   "PENDIENTE"
            ToolTipText     =   "LISTA DE FACTURAS PENDIENTES"
            Top             =   360
            Value           =   -1  'True
            Width           =   1245
         End
      End
      Begin VB.Frame fraAgasto 
         Height          =   6690
         Index           =   0
         Left            =   180
         TabIndex        =   16
         Top             =   390
         Visible         =   0   'False
         Width           =   11370
         Begin VB.Frame fraAgasto 
            BorderStyle     =   0  'None
            Height          =   1305
            Index           =   8
            Left            =   150
            TabIndex        =   46
            Top             =   300
            Width           =   11145
            Begin VB.TextBox TxtAGasto 
               Height          =   315
               Index           =   6
               Left            =   1200
               TabIndex        =   48
               Top             =   810
               Width           =   9885
            End
            Begin VB.TextBox TxtAGasto 
               Height          =   315
               Index           =   7
               Left            =   6930
               TabIndex        =   47
               Top             =   0
               Width           =   4125
            End
            Begin MSDataListLib.DataCombo DtcAGasto 
               Height          =   315
               Index           =   0
               Left            =   1200
               TabIndex        =   51
               Top             =   0
               Width           =   1275
               _ExtentX        =   2249
               _ExtentY        =   556
               _Version        =   393216
               Locked          =   -1  'True
               ListField       =   "NDOC"
               BoundColumn     =   ""
               Text            =   " "
            End
            Begin MSDataListLib.DataCombo DtcAGasto 
               Height          =   315
               Index           =   1
               Left            =   2475
               TabIndex        =   52
               Top             =   0
               Width           =   3525
               _ExtentX        =   6218
               _ExtentY        =   556
               _Version        =   393216
               Locked          =   -1  'True
               ListField       =   "Name"
               BoundColumn     =   ""
               Text            =   " "
            End
            Begin VB.Frame fraAgasto 
               BorderStyle     =   0  'None
               Enabled         =   0   'False
               Height          =   570
               Index           =   7
               Left            =   9390
               TabIndex        =   49
               Top             =   195
               Width           =   1680
               Begin MSComCtl2.DTPicker DtpAGasto 
                  Height          =   315
                  Left            =   195
                  TabIndex        =   50
                  Top             =   210
                  Width           =   1485
                  _ExtentX        =   2619
                  _ExtentY        =   556
                  _Version        =   393216
                  Format          =   50790401
                  CurrentDate     =   37466
               End
            End
            Begin VB.Label LblAgasto 
               Alignment       =   1  'Right Justify
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
               Height          =   315
               Index           =   4
               Left            =   1185
               TabIndex        =   62
               Top             =   405
               Width           =   1290
            End
            Begin VB.Label LblAgasto 
               Alignment       =   1  'Right Justify
               Caption         =   "&Documento :"
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
               Left            =   0
               TabIndex        =   61
               Top             =   105
               Width           =   1080
            End
            Begin VB.Label LblAgasto 
               Alignment       =   1  'Right Justify
               Caption         =   "Monto :"
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
               Left            =   2730
               TabIndex        =   59
               Top             =   495
               Width           =   870
            End
            Begin VB.Label LblAgasto 
               Alignment       =   1  'Right Justify
               Caption         =   "Factura Nº :"
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
               Left            =   0
               TabIndex        =   58
               Top             =   495
               Width           =   1080
            End
            Begin VB.Label LblAgasto 
               Alignment       =   1  'Right Justify
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
               Height          =   315
               Index           =   5
               Left            =   3735
               TabIndex        =   57
               Top             =   405
               Width           =   1620
            End
            Begin VB.Label LblAgasto 
               Alignment       =   1  'Right Justify
               Caption         =   "Diferencia:"
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
               Left            =   5940
               TabIndex        =   56
               Top             =   495
               Width           =   960
            End
            Begin VB.Label LblAgasto 
               Alignment       =   1  'Right Justify
               BackColor       =   &H80000005&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "0,00"
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
               Index           =   6
               Left            =   6930
               TabIndex        =   55
               Top             =   405
               Width           =   1620
            End
            Begin VB.Label LblAgasto 
               Alignment       =   1  'Right Justify
               Caption         =   "Descripción:"
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
               Left            =   0
               TabIndex        =   54
               Top             =   900
               Width           =   1080
            End
            Begin VB.Label LblAgasto 
               Alignment       =   1  'Right Justify
               Caption         =   "Benef.:"
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
               Left            =   5895
               TabIndex        =   53
               Top             =   75
               Width           =   960
            End
            Begin VB.Label LblAgasto 
               Alignment       =   1  'Right Justify
               Caption         =   "&Fecha :"
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
               Left            =   8385
               TabIndex        =   60
               Top             =   495
               Width           =   975
            End
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid GridGastos 
            Height          =   3735
            Index           =   1
            Left            =   390
            TabIndex        =   43
            Tag             =   "1470|4440"
            Top             =   2535
            Visible         =   0   'False
            Width           =   5940
            _ExtentX        =   10478
            _ExtentY        =   6588
            _Version        =   393216
            FixedCols       =   0
            BackColorFixed  =   -2147483646
            ForeColorFixed  =   -2147483639
            BackColorBkg    =   -2147483636
            WordWrap        =   -1  'True
            FocusRect       =   0
            HighLight       =   0
            GridLinesFixed  =   0
            GridLinesUnpopulated=   3
            ScrollBars      =   2
            PictureType     =   1
            BorderStyle     =   0
            FormatString    =   "CUENTA |DESCRIPCION"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            _NumberOfBands  =   1
            _Band(0).Cols   =   2
            _Band(0).GridLinesBand=   0
            _Band(0).TextStyleBand=   0
            _Band(0).TextStyleHeader=   0
         End
         Begin VB.Frame fraAgasto 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   3405
            Index           =   4
            Left            =   6330
            TabIndex        =   26
            Top             =   2520
            Visible         =   0   'False
            Width           =   3300
            Begin MSHierarchicalFlexGridLib.MSHFlexGrid GridGastos 
               Height          =   2760
               Index           =   2
               Left            =   90
               TabIndex        =   44
               Top             =   525
               Width           =   3135
               _ExtentX        =   5530
               _ExtentY        =   4868
               _Version        =   393216
               Cols            =   3
               BackColorFixed  =   -2147483646
               ForeColorFixed  =   -2147483639
               WordWrap        =   -1  'True
               _NumberOfBands  =   1
               _Band(0).Cols   =   3
            End
            Begin VB.ListBox LisAGasto 
               Height          =   2010
               Index           =   1
               Left            =   1710
               TabIndex        =   31
               Top             =   1215
               Width           =   1500
            End
            Begin VB.ListBox LisAGasto 
               BackColor       =   &H00FFFFFF&
               Height          =   2010
               Index           =   0
               ItemData        =   "FrmAsignaGasto.frx":0038
               Left            =   105
               List            =   "FrmAsignaGasto.frx":003A
               MultiSelect     =   2  'Extended
               TabIndex        =   30
               Top             =   1245
               Width           =   1500
            End
            Begin VB.CommandButton cmdHelp 
               BackColor       =   &H00FFFFFF&
               Caption         =   "&Aceptar"
               Height          =   315
               Index           =   2
               Left            =   2040
               Style           =   1  'Graphical
               TabIndex        =   29
               Top             =   120
               Width           =   1170
            End
            Begin VB.OptionButton Option1 
               BackColor       =   &H00FFFFFF&
               Caption         =   "&Propietarios Determinados"
               Height          =   270
               Index           =   1
               Left            =   100
               TabIndex        =   28
               Top             =   540
               Width           =   2190
            End
            Begin VB.OptionButton Option1 
               BackColor       =   &H00FFFFFF&
               Caption         =   "&Todos los Propietarios"
               Height          =   270
               Index           =   0
               Left            =   100
               TabIndex        =   27
               Top             =   165
               Value           =   -1  'True
               Width           =   2190
            End
         End
         Begin VB.Frame fraAgasto 
            Caption         =   "Distribución del Gasto:"
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
            ForeColor       =   &H80000008&
            Height          =   1215
            Index           =   2
            Left            =   195
            TabIndex        =   23
            Top             =   1590
            Width           =   11025
            Begin VB.CommandButton cmdHelp 
               Height          =   330
               Index           =   1
               Left            =   5820
               Picture         =   "FrmAsignaGasto.frx":003C
               Style           =   1  'Graphical
               TabIndex        =   12
               Tag             =   "Titulo LIKE"
               Top             =   585
               Width           =   255
            End
            Begin VB.CheckBox Check1 
               BackColor       =   &H80000005&
               Height          =   195
               Index           =   1
               Left            =   7185
               TabIndex        =   11
               Top             =   660
               Width           =   180
            End
            Begin VB.CheckBox Check1 
               BackColor       =   &H80000005&
               Height          =   195
               Index           =   0
               Left            =   6420
               TabIndex        =   9
               Top             =   660
               Width           =   195
            End
            Begin VB.ComboBox Combo1 
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   360
               Index           =   1
               ItemData        =   "FrmAsignaGasto.frx":0186
               Left            =   10260
               List            =   "FrmAsignaGasto.frx":0188
               Style           =   2  'Dropdown List
               TabIndex        =   7
               Top             =   570
               Width           =   585
            End
            Begin VB.CommandButton cmdHelp 
               Height          =   315
               Index           =   0
               Left            =   1380
               Picture         =   "FrmAsignaGasto.frx":018A
               Style           =   1  'Graphical
               TabIndex        =   13
               Tag             =   "CodGasto LIKE"
               Top             =   585
               Width           =   255
            End
            Begin VB.ComboBox Combo1 
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   360
               Index           =   0
               ItemData        =   "FrmAsignaGasto.frx":02D4
               Left            =   9390
               List            =   "FrmAsignaGasto.frx":02FE
               Style           =   2  'Dropdown List
               TabIndex        =   6
               Top             =   570
               Width           =   870
            End
            Begin VB.TextBox TxtAGasto 
               Alignment       =   1  'Right Justify
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   360
               Index           =   4
               Left            =   7650
               TabIndex        =   5
               Top             =   570
               Width           =   1740
            End
            Begin VB.TextBox TxtAGasto 
               Height          =   360
               Index           =   1
               Left            =   1665
               TabIndex        =   3
               Tag             =   "Titulo"
               Top             =   570
               Width           =   4155
            End
            Begin VB.TextBox TxtAGasto 
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
               Height          =   360
               Index           =   3
               Left            =   6900
               TabIndex        =   25
               Top             =   570
               Width           =   750
            End
            Begin VB.TextBox TxtAGasto 
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
               Height          =   360
               Index           =   2
               Left            =   6105
               TabIndex        =   24
               Top             =   570
               Width           =   795
            End
            Begin VB.TextBox TxtAGasto 
               Height          =   360
               Index           =   0
               Left            =   195
               MaxLength       =   6
               TabIndex        =   1
               Tag             =   "CodGasto"
               Top             =   570
               Width           =   1185
            End
            Begin VB.Label Label1 
               Alignment       =   2  'Center
               BackColor       =   &H80000002&
               Caption         =   "&ALICUOTA"
               ForeColor       =   &H80000009&
               Height          =   210
               Index           =   3
               Left            =   6900
               TabIndex        =   10
               Top             =   360
               Width           =   810
            End
            Begin VB.Label Label1 
               Alignment       =   2  'Center
               BackColor       =   &H80000002&
               Caption         =   "PERIODO"
               ForeColor       =   &H80000009&
               Height          =   210
               Index           =   5
               Left            =   9405
               TabIndex        =   14
               Top             =   360
               Width           =   1410
            End
            Begin VB.Label Label1 
               Alignment       =   2  'Center
               BackColor       =   &H80000002&
               Caption         =   "&MONTO"
               ForeColor       =   &H80000009&
               Height          =   210
               Index           =   4
               Left            =   7665
               TabIndex        =   4
               Top             =   360
               Width           =   1740
            End
            Begin VB.Label Label1 
               Alignment       =   2  'Center
               BackColor       =   &H80000002&
               Caption         =   "DE&SCRIPCION"
               ForeColor       =   &H80000009&
               Height          =   210
               Index           =   1
               Left            =   1665
               TabIndex        =   2
               Top             =   360
               Width           =   4440
            End
            Begin VB.Label Label1 
               Alignment       =   2  'Center
               BackColor       =   &H80000002&
               Caption         =   "&CUENTA"
               ForeColor       =   &H80000009&
               Height          =   210
               Index           =   0
               Left            =   210
               TabIndex        =   0
               Top             =   360
               Width           =   1455
            End
            Begin VB.Label Label1 
               Alignment       =   2  'Center
               BackColor       =   &H80000002&
               Caption         =   "COM&UN"
               ForeColor       =   &H80000009&
               Height          =   210
               Index           =   2
               Left            =   6105
               TabIndex        =   8
               Top             =   360
               Width           =   810
            End
         End
         Begin VB.Frame fraAgasto 
            Caption         =   "Detalle Asignacion de Gasto:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   3450
            Index           =   1
            Left            =   180
            TabIndex        =   21
            Top             =   2895
            Width           =   11070
            Begin MSHierarchicalFlexGridLib.MSHFlexGrid GridGastos 
               Height          =   2775
               Index           =   0
               Left            =   180
               TabIndex        =   42
               Top             =   480
               Width           =   10725
               _ExtentX        =   18918
               _ExtentY        =   4895
               _Version        =   393216
               Cols            =   6
               FixedCols       =   0
               BackColorFixed  =   -2147483646
               ForeColorFixed  =   -2147483639
               BackColorSel    =   65280
               BackColorBkg    =   -2147483636
               AllowBigSelection=   0   'False
               BorderStyle     =   0
               FormatString    =   "^CUENTA |DESCRIPCION |COMUN |ALICUOTA |>MONTO |PERIODO"
               BandDisplay     =   1
               _NumberOfBands  =   1
               _Band(0).Cols   =   6
            End
            Begin VB.Image ImgOK 
               Enabled         =   0   'False
               Height          =   240
               Index           =   1
               Left            =   0
               Top             =   720
               Visible         =   0   'False
               Width           =   240
            End
            Begin VB.Image ImgOK 
               Enabled         =   0   'False
               Height          =   240
               Index           =   0
               Left            =   30
               Top             =   240
               Visible         =   0   'False
               Width           =   240
            End
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
            Picture         =   "FrmAsignaGasto.frx":033E
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmAsignaGasto.frx":04C0
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmAsignaGasto.frx":0642
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmAsignaGasto.frx":07C4
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmAsignaGasto.frx":0946
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmAsignaGasto.frx":0AC8
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmAsignaGasto.frx":0C4A
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmAsignaGasto.frx":0DCC
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmAsignaGasto.frx":0F4E
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmAsignaGasto.frx":10D0
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmAsignaGasto.frx":1252
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmAsignaGasto.frx":13D4
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "FrmAsignaGasto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
    '---------------------------------------------------------------------------------------------
    'Sistema de Administración de Condominios-SINAHI TECH
    'Modula de Asignacion de Gastos/05/08/2002
    'Declaracion de Variables Globales a nivel de módulo
    '---------------------------------------------------------------------------------------------
    Dim AdogastosTran As ADODB.Recordset
Attribute AdogastosTran.VB_VarHelpID = -1
    Dim RstFacturas As ADODB.Recordset, RstAsignado As ADODB.Recordset
    Dim cnn As ADODB.Connection
    Dim stEditFC As Boolean, blnRet As Boolean
    Dim VecMes(1 To 12, 1 To 2)
    Dim datPeriodo As Date
    Dim WrkEspacio As Workspace
    Dim Fila%, vecFC()
    '
    Const strArchivo$ = "C:\TempCXC.log"
    Const Verde# = &HFF00&
    Const Blanco# = &H80000005
    '---------------------------------------------------------------------------------------------
    
    '---------------------------------------------------------------------------------------------
    Private Sub Check1_Click(Index As Integer) '-
    '---------------------------------------------------------------------------------------------
    'en caso de error continúa la ejecución
    On Error Resume Next
    '
    With TxtAGasto(4)
    '
        .Text = LblAgasto(6)
        .SelStart = 0
        .SelLength = Len(.Text)
        .SetFocus
    '
    End With
    fraAgasto(4).Visible = False
    '
    End Sub
    
    '---------------------------------------------------------------------------------------------
    Private Sub Check1_MouseDown(Index As Integer, Button As Integer, _
    Shift As Integer, X As Single, Y As Single) '
    '---------------------------------------------------------------------------------------------
    If Button = 2 Then      'Si presiona el segundo boton del mouse
        Select Case Index   'muestra un menu emergente
    '
            Case 0  'Gasto Comun / No_Comun
            '------------------------------
                If Check1(0).Value = False Then
                
                    Option1(0).Visible = True
                    Option1(1).Visible = True
                    fraAgasto(4).Visible = True
                    GridGastos(2).Visible = False
                    If Option1(0) Then
                        fraAgasto(4).Height = 960
                    Else
                        fraAgasto(4).Height = 3405
                    End If
                    '
                End If
    '
        End Select
    '
    End If
    '
    End Sub

    '---------------------------------------------------------------------------------------------
    Private Sub CmdHelp_Click(Index As Integer) '-
    '---------------------------------------------------------------------------------------------
    '
    Select Case Index
    '
        Case 0, 1 'Boton de ayuda por {1}codigo y {2}Descripcion
        '-------------------------------------------------------
            If TxtAGasto(Index) = "" Then Exit Sub
            With GridGastos(1)  'MUESTRA AL USUARIO LISTA DE CUENTAS/DESCRIPCION S/PATRON*
                If .Visible = True Then
                    .Visible = False
                    Exit Sub
                End If
                
                With AdoCuentas.Recordset
                    .Filter = cmdHelp(Index).Tag & " '" & TxtAGasto(Index) & "*'"
                    If Not .EOF Then
                        Set GridGastos(1).DataSource = AdoCuentas
                        If GridGastos(1).RowIsVisible(0) Then GridGastos(1).RowHeight(0) = 0
                        Call centra_titulo(GridGastos(1), True)
                    End If
                End With
                
            End With
            
        
        Case 2  'Boton aceptar de menu emergente
        '---------------------------------------
            fraAgasto(4).Visible = False
            With TxtAGasto(4)
                .SelStart = 0
                .SelLength = Len(.Text)
                .SetFocus
            End With
        
        Case 3  'Boton Buscar Ficha Lista
        '---------------------------------------
            Call Buscar_Factura
            
    End Select
    
    End Sub

    '--------------------------------------------------------------------
    Private Sub Combo1_KeyPress(Index As Integer, KeyAscii As Integer) '-
    '--------------------------------------------------------------------
    '
    '
    If KeyAscii = 13 Then   'Presionó {enter}
    '
        Select Case Index   'Ejecuta el procedimiento de acuero al
    '                        control que hace el llamdo
            Case 0  'Lista de los Meses
    '
                If Combo1(0) = "" Then
                    MsgBox "Seleccione un Mes..", vbExclamation
                    Exit Sub
                End If
                Combo1(1) = Format(Date, "YY")
                Combo1(1).SetFocus
                
            Case 1  'Combo del año {Guarda el registro actual}
    '
                If Validar_Guardar Then Exit Sub
                If Not Ajustes_Reintegros Then Call RtnGuardar
    '
        End Select
    '
        
    End If
    '
    End Sub

    '----------------------------------'Selecciona la lista de Facturas {Pendientes} / {Asignadas}
    Private Sub DataGrid1_DblClick() '-
    '-------------------------------
    '
    Call RtnEstado(6, Toolbar1)
    
    Select Case DataGrid1.Caption
    '
        Case "LISTA DE FACTURAS PENDIENTES"
    '   -----------------------------------
            If GridGastos(0).Rows > 2 Or GridGastos(0).Rows = 1 Then GridGastos(0).Rows = 2
            Call rtnLimpiar_Grid(GridGastos(0))
            Toolbar1.Buttons(10).Enabled = False
    '
        Case "LISTA DE FACTURAS ASIGNADAS", "LISTA DE FACTURAS CANCELADAS"
    '   ----------------------------------
            Call Mostrar_Cargado
            Toolbar1.Buttons("New").Enabled = False
            
    End Select
    Fila = DataGrid1.Bookmark
    '
    Call RtnAsignaCampos
    
    SSTab1.TabEnabled(0) = True
    '
    End Sub

Private Sub DataGrid1_HeadClick(ByVal ColIndex As Integer)
'ordena la lista por el encabezado de la columna
Static orden(6) As String
If orden(ColIndex) = "" Then orden(ColIndex) = "ASC"
RstFacturas.Sort = DataGrid1.Columns(ColIndex).DataField & " " & orden(ColIndex) & ",Ndoc"
orden(ColIndex) = IIf(orden(ColIndex) = "ASC", "DESC", "ASC")
End Sub

    '----------------------------------------------------------------
    Private Sub DtcAGasto_Click(Index As Integer, Area As Integer) '-
    '----------------------------------------------------------------
    Dim strCriterio As String
    If Area = 2 Then    'SI EL USUARIO SELECCIONA UN ELEMENTO DE LA LISTA
        
        Select Case Index
            
            Case 0, 1
                
                If Index = 0 Then
                    strCriterio = "Ndoc  = '" & DtcAGasto(Index) & "'"
                Else
                    strCriterio = "Name  = '" & DtcAGasto(Index) & "'"
                End If
                With RstFacturas
                    .MoveFirst
                    .Find strCriterio
                    If Not .EOF Or .BOF Then
                        DtcAGasto(0) = !NDoc
                        DtcAGasto(1) = !Name
                        LblAgasto(4) = IIf(IsNull(!Fact), "", Format(!Fact, "000000 "))
                        LblAgasto(5) = Format(!Total, "#,##0.00")
                    End If
                End With
            
        End Select
        '
    End If
    End Sub


    'Rev.27/08/2002-------------------------Configura la presentación en pantalla y---------------
    Private Sub Form_Load() '---------------el origen de los datos--------------------------------
    '---------------------------------------------------------------------------------------------
    'variables locales
    Dim IntLis%, strSQL$
    '
    
    Set GridGastos(0).FontFixed = LetraTitulo(LoadResString(527), 6, True)
    Set DataGrid1.HeadFont = LetraTitulo(LoadResString(527), 7.5, True)
    Set DataGrid1.Font = LetraTitulo(LoadResString(528), 8)
    Set AdogastosTran = New ADODB.Recordset
    Set cnn = New ADODB.Connection
    Set RstFacturas = New ADODB.Recordset
    
    On Error GoTo ValidaInicio
    '
    cmdHelp(3).Picture = LoadResPicture("Buscar", vbResIcon)
    ImgOK(0).Picture = LoadResPicture("UNCHECKED", vbResBitmap)
    ImgOK(1).Picture = LoadResPicture("CHECKED", vbResBitmap)
    '
    'Crea una conexion al inmueble seleccionado
    
    cnn.CursorLocation = adUseClient
    cnn.Open cnnOLEDB + mcDatos
    '
    'Busca el último período facturado
    RstFacturas.CursorLocation = adUseClient
    RstFacturas.Open "SELECT MAX(Periodo) FROM Factura WHERE Fact Not Like 'CH%' OR Fact Is Nul" _
    & "l;", cnn, adOpenKeyset, adLockOptimistic, adCmdText
    
    datPeriodo = IIf(IsNull(RstFacturas.Fields(0)), "15/01/1975", RstFacturas.Fields(0))
    RstFacturas.Close
    'ABRE LOS ADODB.Recordset NECESARIOS EN EL FORMULARIO
    '---------------------------------------------------------------------------------------------
    'strSQl = ftnSQL("PENDIENTE")
    strSQL = "SELECT Cpp.*,Proveedores.NombProv as Name FROM Proveedores INNER JOIN Cpp ON " _
    & "Proveedores.Codigo=Cpp.CodProv WHERE  CodInm ='" & gcCodInm & "' ORDER BY Frecep DESC, Ndoc"
    RstFacturas.Open strSQL, cnnConexion, adOpenKeyset, adLockOptimistic, adCmdText
    RstFacturas.Filter = "Estatus='PENDIENTE'"
    Set DataGrid1.DataSource = RstFacturas
    'DataGrid1.ReBind
    AdoCuentas.ConnectionString = cnn.ConnectionString
    AdoCuentas.RecordSource = "Tgastos"
    AdoCuentas.Refresh
    '------------------|
    With GridGastos(0)
        .Row = 0
        IntLis = CInt(Format(DateAdd("yyyy", -2, Date), "YY"))
        For I = 0 To .Cols - 1
            .Col = I
            .CellAlignment = flexAlignCenterCenter
            .ColWidth(I) = TxtAGasto(I).Width
            Combo1(1).AddItem Format((IntLis + I), "00")
        Next
        'Combo1(1).AddItem ("99")
    '   -------------|
        .ColWidth(5) = Combo1(0).Width + Combo1(1).Width    '
        .ColAlignment(1) = flexAlignLeftCenter
        '.Col = 0
    End With
    '--------------|
    For I = 0 To 1
    '   CARGA DEL ARCHIVO DE RECURSOS LAS IMAGENES PARA LOS BOTONES
        cmdHelp(I).Picture = LoadResPicture("DataCombo", vbResBitmap)
    Next
    '--------------|
    For I = 1 To 12 'LLENAR UN VECTOR NUMERADO CON LOS MESES DEL AÑO
        VecMes(I, 1) = UCase(MonthName(I, True))
        VecMes(I, 2) = I
    Next
    '--------------|
    '
    SSTab1.tab = 1
    cnn.Execute "DELETE * FROM AsignaGasto WHERE Ndoc='0' AND Cargado>#" & _
    Format(datPeriodo, "mm/dd/yy") & "#;"
ValidaInicio:
'-----------
    '
    If Err.Number <> 0 Then
        MsgBox Err.Description & " " & Err.Source, vbCritical, App.ProductName
        Unload FrmAsignaGasto
        MousePointer = vbDefault
    End If
    '
    End Sub
    


    Private Sub Form_Resize()
    '
    Dim ancho As Long
    On Error Resume Next
    With SSTab1
        .Height = Me.ScaleHeight - .Top
        .Width = Me.ScaleWidth - .Left
        fraAgasto(0).Left = .Left + 100
        fraAgasto(0).Height = .Height - fraAgasto(0).Top - 200
        fraAgasto(0).Width = SSTab1.Width - (fraAgasto(0).Left * 2)
        DataGrid1.Left = fraAgasto(0).Left + 100
        DataGrid1.Height = .Height - fraAgasto(5).Height - DataGrid1.Top - 200
        DataGrid1.Width = fraAgasto(0).Width - DataGrid1.Left
        'DataGrid1.Columns(4).Width = 0
        For I = 0 To DataGrid1.Columns.count - 1
            If I <> 4 Then ancho = ancho + DataGrid1.Columns(I).Width
            
        Next
        DataGrid1.Columns(4).Width = DataGrid1.Width - ancho - 1000
        fraAgasto(5).Left = DataGrid1.Left
        fraAgasto(3).Left = fraAgasto(5).Left + fraAgasto(5).Width + 100
        fraAgasto(6).Left = fraAgasto(5).Left + fraAgasto(5).Width + fraAgasto(3).Width + 200
        fraAgasto(5).Top = DataGrid1.Top + DataGrid1.Height + 100
        fraAgasto(6).Top = DataGrid1.Top + DataGrid1.Height + 100
        fraAgasto(3).Top = DataGrid1.Top + DataGrid1.Height + 100
        fraAgasto(8).Left = (fraAgasto(0).Width - fraAgasto(8).Width) / 2
        fraAgasto(2).Left = fraAgasto(8).Left
        fraAgasto(1).Left = fraAgasto(8).Left
        GridGastos(1).Left = fraAgasto(2).Left + TxtAGasto(0).Left
        fraAgasto(4).Left = TxtAGasto(2).Left + fraAgasto(2).Left
        fraAgasto(1).Height = fraAgasto(0).Height - fraAgasto(1).Top - 200
        GridGastos(0).Height = fraAgasto(1).Height - GridGastos(0).Top - 200
        
    End With
    '
    End Sub

    '------------------------------------------------
    Private Sub GridGastos_Click(Index As Integer) '-
    '------------------------------------------------
    '
    Select Case Index
    '
        Case 0  'Edición de distribución del gasto
    '   ---------------------
    '
        Case 1  'Listado de Códigos Catalogos de Gasto
        '--------------------
            With GridGastos(1)
                For I = 0 To 1
                    TxtAGasto(I) = .TextMatrix(.RowSel, I)
                Next
                With AdoCuentas.Recordset
                    .MoveFirst
                    .Find "CodGasto = '" & TxtAGasto(0) & "'"
                    Check1(0).Value = IIf(.Fields("Comun") = -1, 1, 0)
                    Check1(1).Value = IIf(.Fields("Alicuota") = -1, 1, 0)
                End With
                .Visible = False
            End With
            TxtAGasto(4).SetFocus
        End Select
    '
    End Sub


    '---------------------------------------------------------------------------------------------
    Private Sub GridGastos_EnterCell(Index As Integer) '-
    '---------------------------------------------------------------------------------------------
    If Index = 2 And GridGastos(2).Col = 2 And GridGastos(2).RowSel > 0 Then
        With GridGastos(2)
            .Text = CCur(GridGastos(2).Text)
            Dim curMonto As Currency
            
            For I = 1 To .Rows - 1
                If .TextMatrix(I, 2) = "" Then .TextMatrix(I, 2) = 0
                curMonto = curMonto + .TextMatrix(I, 2)
            Next
            TxtAGasto(4) = Format(CLng(curMonto), "#,##0.00 ")
        End With
    End If
    End Sub

    '---------------------------------------------------------------------------------------------
    Private Sub GridGastos_KeyPress(Index As Integer, _
    KeyAscii As Integer) 'Permite editar la columna de montos -
    '---------------------------------------------------------------------------------------------
    '
    If Index = 2 And GridGastos(2).Col = 2 Then
    '
        With GridGastos(2)
    '
            If .RowSel > 0 Then
            Call Validacion(KeyAscii, "0123456789.,")
    '
                If KeyAscii = 46 Then KeyAscii = 44 'CONVIERTE PUNTO(.) EN COMA(,)
                If KeyAscii >= 48 And KeyAscii <= 57 Or KeyAscii = 44 Then
    '
                    .TextMatrix(.RowSel, 2) = _
                    .TextMatrix(.RowSel, 2) & Chr(KeyAscii)
    '
                ElseIf KeyAscii = 13 Then
    '
                    GridGastos_LeaveCell 2
                    If .RowSel = (.Rows - 1) Then
    '                   Si esta en la última cuadricula vuelve a la priemra
                        .Row = 1
                    Else
    '                   Baja a la cuadricula inferior siguiente
                        .Row = (.RowSel + 1)
                    End If
                    
                    GridGastos_EnterCell 2
                    GridGastos(2).SetFocus
    '
                End If
    '
            End If
    '
        End With
    '
    End If
    '
    End Sub
    
    '--------------------------------------------------------------------------------------------'
    Private Sub GridGastos_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
    '--------------------------------------------------------------------------------------------'
    'Permite Borrar el contenido de la celda (total/parcial)
    If Index = 2 And GridGastos(2).Col = 2 And GridGastos(2).RowSel > 0 Then
    '
        With GridGastos(2)
    '
            Select Case KeyCode 'Selecciona la tecla que presionó
    '       -----------------------------------------------------
                Case 46 '{SUPRIMIR} 'Borra todo
    '           --------------------------------
                    '.TextMatrix(.RowSel, 2) = ""
                    .Text = ""
    '
                Case 8  '{BACKSPACE} 'Borra el último caracter
    '           ----------------------------------------------
                    If Len(.Text) > 0 Then
                        .Text = Left(.Text, Len(.Text) - 1)
                    End If
    '
            End Select
    '
        End With
    '
    End If
    '
    End Sub

    '----------------------------------------------------
    Private Sub GridGastos_LeaveCell(Index As Integer) '-
    '----------------------------------------------------
    If Index = 2 And GridGastos(2).Col = 2 Then
        GridGastos(2).Text = Format(GridGastos(2).Text, "#,##0.00 ")
    End If
    End Sub

    

    '------------------------------------------------------------------------
    Private Sub LisAGasto_DblClick(Index As Integer) 'Selección de Propietarios
    '------------------------------------------------------------------------
    '
    Select Case Index
        Case 0  'Lista de Todos los propietarios
    '   -----------------------------------------------
            LisAGasto(1).AddItem (LisAGasto(0).Text)
        
        Case 1  'Lista de propietarios pre-seleccionados
    '   -------------------------------------------------
            LisAGasto(1).RemoveItem (LisAGasto(1).ListIndex)
    '
    End Select
    '
    End Sub


    '---------------------------------------------
    Private Sub Option1_Click(Index As Integer) '-
    '---------------------------------------------
    '
    Select Case Index
    '----------------
        'Todos los propietarios
        Case 0: fraAgasto(4).Height = 960
        'Determinados propietarios
        Case 1
            LisAGasto(1).Clear
            Set AdogastosTran = New ADODB.Recordset
            With AdogastosTran
    '           Selecciona todos los codigos de propietarios del inmueble
    '           y llena la lista para que el usuario realice su selección
                .Open "SELECT Codigo,Nombre  FROM Propietarios " _
                & "ORDER BY Codigo", cnn, adOpenKeyset, adLockReadOnly
                If .EOF Then Exit Sub
                .MoveFirst
                Do Until .EOF
                    LisAGasto(0).AddItem (!Codigo)
                    .MoveNext
                Loop
                AdogastosTran.Close
                Set AdogastosTran = Nothing
    '
            End With
            If fraAgasto(4).Height < 3405 Then
                fraAgasto(4).Height = 3405
            End If
    '
    End Select
    '---------
    End Sub

    'Rev27/08/2002------------------------------------------------Muestra la selección de Facturas
    Private Sub OptTipoFactura_Click(Index As Integer) '|
    '---------------------------------------------------------------------------------------------
    '
    MousePointer = vbHourglass
    Select Case Index
        Case 0, 1, 8
            OptTipoFactura(9).Visible = OptTipoFactura(8)
            RstFacturas.Filter = "Estatus='" & OptTipoFactura(Index).Tag & "'"
            DataGrid1.Caption = OptTipoFactura(Index).ToolTipText
        
        Case 2, 3, 4, 5, 6, 7
            RstFacturas.Sort = OptTipoFactura(Index).Tag
            MousePointer = vbDefault
            
    End Select
    MousePointer = vbDefault
    '
    End Sub

    '---------------------------------------------------------------------------------------------
    Private Sub SSTab1_Click(PreviousTab As Integer)                                            '|
    '---------------------------------------------------------------------------------------------
    '
    Select Case SSTab1.tab
    '
        Case 0  'Ficha Datos Generales
    '   -----------------------------------------
            DtpAGasto.Value = Date
            Set DtcAGasto(0).RowSource = RstFacturas
            Set DtcAGasto(1).RowSource = RstFacturas
            fraAgasto(0).Visible = True
            DataGrid1.Visible = False
            fraAgasto(5).Visible = False
            fraAgasto(3).Visible = False
            fraAgasto(6).Visible = False
            
    '
        Case 1  'Ficha lista
    '   -----------------------------------------
            If Not RstFacturas.EOF And Not RstFacturas.BOF Then RstFacturas.Move (Fila - 1)
            For I = 5 To Toolbar1.Buttons.count - 1
                Toolbar1.Buttons(I).Enabled = False
            Next
            fraAgasto(0).Visible = False
            DataGrid1.Visible = True
            fraAgasto(5).Visible = True
            fraAgasto(3).Visible = True
            fraAgasto(6).Visible = True

            '
    End Select
    '---------
    End Sub

    
    '--------------------------------------------------------------------
    Private Sub Toolbar1_ButtonClick(ByVal Button As ComctlLib.Button) '-
    '--------------------------------------------------------------------
    'variables locales
    Dim Reg As Long
    '
    With RstFacturas
    '
        Select Case Button.Index
        '
            Case 1  'IR AL PRIMER REGISTRO
                .MoveFirst
                If .BOF Then Exit Sub
                Call RtnAsignaCampos
    '
            Case 2  'IR AL REGISTRO ANTERIOR
                .MovePrevious
                If .BOF Then .MoveLast
                Call RtnAsignaCampos
    '
            Case 3  'IR AL REGISTRO SIGUIENTE
                .MoveNext
                If .EOF Then .MoveFirst
                Call RtnAsignaCampos
    '
            Case 4  'IR AL ULTIMO REGISTRO
                .MoveLast
                If .EOF Then Exit Sub
                Call RtnAsignaCampos
    '
            Case 5  'NUEVO REGISTRO
            
                Call RtnEstado(Button.Index, Toolbar1)
                SSTab1.TabEnabled(1) = False
                TxtAGasto(0).Enabled = True
                For I = 0 To 2: fraAgasto(I).Enabled = True
                Next
                LblAgasto(6) = LblAgasto(5)
                TxtAGasto(0).SetFocus
                cnn.BeginTrans
    '
            Case 6   'GUARDAR EL REGISTRO ACTIVO
                If DtpAGasto = "" Or DtpAGasto = Null Then
                    MsgBox ("Debe ingresar la Fecha")
                    DtpAGasto.SetFocus
                    Exit Sub
                End If
    '
                If Respuesta("Esta seguro de realizar toda la transacción?") = True _
                And LblAgasto(6) = 0 Then
    '
                    SSTab1.TabEnabled(1) = True
                    SSTab1.tab = 1
                    cnn.CommitTrans
                    Call rtnBitacora("Actualizado la asignación de gastos Doc.:" & DtcAGasto(0))
                    
                    If Not OptTipoFactura(8) Then   'SI NO ES UNA FACTURA PAGADA
                        cnnConexion.Execute "UPDATE Cpp SET Estatus='ASIGNADO',Usuario='" _
                        & gcUsuario & "',Freg=DATE() WHERE Ndoc ='" & DtcAGasto(0) & "'"
                    End If
                    '
                    If Not .EOF Or Not .BOF Then
                        .MoveFirst
                        .Find "Ndoc='" & DtcAGasto(0) & "'"
                        If Not .EOF Then .Update "Detalle", TxtAGasto(6) 'actualiza la descrip
                    End If
                    cnnConexion.Execute "UPDATE Cheque SET Concepto ='" & TxtAGasto(6) & "' WHE" _
                    & "RE Clave IN (SELECT Clave FROM ChequeDetalle WHERE NDoc='" & _
                    TxtAGasto(6) & "')", n
                    If n > 0 Then
                        Call rtnBitacora("Actualizado el conpto de (" & n & ") cheque(s)")
                    End If
    '
                Else
                '
                    cnn.RollbackTrans
                    Call rtnBitacora("Actualización Asignación de Gastos Doc.:" & DtcAGasto(0) _
                    & " NO ACEPTADA POR EL USUARIO")
                    MsgBox "Transacción Cancelada..." & IIf(LblAgasto(6) <> 0, vbCrLf _
                    & "Falta Distribuir " & LblAgasto(6), "")
                    '
                End If
    '
                For I = 1 To 4: Toolbar1.Buttons(I).Enabled = True
                Next
                Toolbar1.Buttons("Close").Enabled = True
    '
                stEditFC = False
                For I = 0 To 1
                    DtcAGasto(I) = ""
                    LblAgasto(I + 4) = ""
                Next
                GridGastos(0).Rows = 2
                Call rtnLimpiar_Grid(GridGastos(0))
                For I = 1 To 2: fraAgasto(0).Enabled = False
                Next
                'DataGrid1.ReBind
                Toolbar1.MousePointer = vbDefault
                .Requery
    '
            Case 7  'BUSCAR UN REGISTRO
    '       ---------------------------
                SSTab1.tab = 1
    '
            Case 8  'CANCELAR REGISTRO
    '       --------------------------
                cnn.RollbackTrans
                fraAgasto(4).Visible = False
                Call rtnBitacora("Actualización Asignación de Gastos Doc.:" & DtcAGasto(0) _
                    & " CANCELADA POR EL USUARIO")
                Call RtnEstado(Button.Index, Toolbar1)
                GridGastos(1).Visible = False
                LblAgasto(6) = LblAgasto(5)
                For I = 1 To 2
                    fraAgasto(I).Enabled = False
                    TxtAGasto(I) = ""
                    TxtAGasto(I + 2) = ""
                Next
                SSTab1.TabEnabled(1) = True
                For I = 0 To 1
                    Combo1(I).ListIndex = -1
                    TxtAGasto(I) = ""
                Next
                GridGastos(0).Rows = 2
                Call rtnLimpiar_Grid(GridGastos(0))
                stEditFC = False
                If Dir(strArchivo) <> "" Then Kill (strArchivo)
                If OptTipoFactura(8) Or OptTipoFactura(1) Then Toolbar1.Buttons("New").Enabled = False
                MsgBox ("Registro Cancelado...."), vbInformation, App.ProductName
    '
            Case 9  'ELIMINAR REGISTRO ACTIVO
    '       ----------------------------------
                Dim K%
                If RstFacturas("Estatus") = "ASIGNADO" Then
    '
                With GridGastos(0)
                    .Col = 5
                    For I = 1 To .Rows - 1
                        .Row = I
                        If .CellBackColor = Verde Then
                            MsgBox "Imposible eliminar esta factura.", vbInformation, _
                            App.ProductName
                            Exit Sub
                        End If
                    Next
                End With
                    If MsgBox("Esta seguro que desea eliminar la asigación del gasto de esta fa" _
                    & "ctura...?", vbYesNo + vbQuestion, "Asignar Gasto") = vbYes Then
                        If MsgBox("Esta operación colocará como 'PENDIENTE' esta factura," _
                        & vbCrLf & "Está seguro que desea continuar...?", vbYesNo + vbQuestion, _
                        "Asignar Gasto") = vbYes Then
                            
                            cnn.Execute "DELETE * FROM AsignaGasto WHERE ndoc = '" & RstFacturas("ndoc") & "'"
                            cnn.Execute "DELETE * FROM GastoNoComun WHERE CodGasto & Concepto" & _
                            "& Periodo & Fecha & Hora IN (SELECT CodGasto & Detalle & Periodo" & _
                            " & Fecha & Hora FROM Cargado WHERE Ndoc='" & RstFacturas("ndoc") & _
                            "');", Reg
                            MsgBox "Emilinado " & Reg & " Registro(s) de los Gastos No Comunes"
                            cnn.Execute "DELETE FROM Cargado WHERE ndoc='" & RstFacturas("ndoc") & "'"
                            With GridGastos(0)
                            For K = 1 To .Rows - 1
                            '
                            If booFondo(.TextMatrix(K, 0)) Then
                                'actualiza como eliminado el movimiento del fondo
                                cnn.Execute "UPDATE MovFondo SET Del=True WHERE CodGasto='" & _
                                .TextMatrix(K, 0) & "' AND TIPO='CH' AND Concepto LIKE '%" & _
                                .TextMatrix(K, 1) & "%" & DtcAGasto(0) & "%' AND Periodo=#" & _
                                Format("01-" & .TextMatrix(K, 5), "mm-dd-yyyy") & "# AND cstr(c" _
                                & "cur(Debe))='" & CCur(.TextMatrix(K, 4)) & "'", Reg
                                'call rtnbitacora("Eliminado
                                Call rtnBitacora("Mov. Fondo Eliminado Doc#" & DtcAGasto(0) & _
                                "( " & Reg & " Reg. Actualizado" & IIf(Reg > 1, "s)", ")"))
                            '
                            End If
                            '
                            '
                            Next K
                            End With
                            
                            cnnConexion.Execute "UPDATE Cpp SET Estatus='PENDIENTE' WHERE Ndoc ='" & RstFacturas("ndoc") & "'"
                            .Requery
                            MsgBox "Factura Actualizada...."
                            If GridGastos(0).Rows > 2 Then GridGastos(0).Rows = 2
                            If GridGastos(0).Rows > 1 Then Call rtnLimpiar_Grid(GridGastos(0))
                            For I = 1 To GridGastos(0).Rows - 1
                                GridGastos(0).Row = I
                                For j = 2 To 3
                                    GridGastos(0).Col = j
                                    Set GridGastos(0).CellPicture = Nothing
                                Next
                            Next
                            Call rtnBitacora("Asignación del Gastos Doc.:" & DtcAGasto(0) _
                            & " Eliminada")
                        End If
                    End If
    '
                Else
    '
                    MsgBox "Imposible ejecuatar esta acción ahora", vbInformation, App.Title & " Editar Registro"
    '
                End If
    '
            Case 10 'MODIFICAR REGISTRO ACTIVO
    '       -----------------------------------
                Call rtn_modificar
    '
            Case 11     'IMPRIMIR REPORTE DE ASIGNACION
    '       --------------------------------------------
                mcTitulo = "Reporte de Asignación de Gastos"
                mcReport = "AsignaGasto.Rpt"
                mcDatos = gcPath + gcUbica + "inm.mdb"
                mcOrdCod = ""
                mcOrdAlfa = ""
                mcCrit = "{QdfAsignaGasto.Cpp.CodInm}='" & gcCodInm & "' AND {AsignaGasto.Fecha}=Date(" & Format(Date, "yyyy,mm,dd") & ")"
                FrmReport.Frame2.Visible = True
                FrmReport.Show
    '
            Case 12 'SALIR
    '       --------------
                Unload Me
    '
        End Select
    '
    End With
    '
    End Sub

    '---------------------------------------------------------------------------------------------
    '   Rutina:     rtnBuscaTgasto
    '
    '
    '---------------------------------------------------------------------------------------------
    Private Sub RtnBuscaTGasto(Txt%)
    '
    With AdoCuentas.Recordset
    '
        .Filter = TxtAGasto(Txt).Tag & "='" & TxtAGasto(Txt) & "'"
        If Not .EOF Then
            
            TxtAGasto(0) = .Fields("CodGasto")
            TxtAGasto(1) = IIf(IsNull(.Fields("Titulo")), "--VACIO--", .Fields("Titulo"))
            Check1(0).Value = IIf(.Fields("Comun") = -1, 1, 0)
            Check1(1).Value = IIf(.Fields("Alicuota") = -1, 1, 0)
            'TxtAGasto(4).SetFocus
            '
            With TxtAGasto(4)
    '
                .Text = LblAgasto(6)
                .SelStart = 0
                .SelLength = Len(.Text)
                .SetFocus
    '
            End With
        Else
            MsgBox "No se hallaron Coincidencias", vbInformation, App.ProductName
        End If
        
        '
    End With
    '
    End Sub


    '---------------------------------------------------------------------------------------------
    Private Sub RtnAsignaCampos()   '
    '---------------------------------------------------------------------------------------------
    '
    With RstFacturas
    '
        DtcAGasto(0) = !NDoc
        DtcAGasto(1) = !Benef
        LblAgasto(4) = IIf(IsNull(!Fact), "", Format(!Fact, "000000 "))
        LblAgasto(5) = Format(!Total, "#,##0.00 ")
        
        TxtAGasto(6) = !detalle
        TxtAGasto(7) = !Benef
    '
    End With
    
    SSTab1.tab = 0
    '
    End Sub


    '-----------------------------------------------
    Private Sub TxtAGasto_Click(Index As Integer) '-
    '-----------------------------------------------
    Select Case Index
        Case 0, 1, 4
            With TxtAGasto(Index)
                .SelStart = 0
                .SelLength = Len(.Text)
            End With
    End Select
    End Sub

    '--------------------------------------------------
    Private Sub TxtAGasto_DblClick(Index As Integer) '-
    '--------------------------------------------------
    'Valida que la cantidad no sea mayor que el resto de
    'la factura
    If TxtAGasto(4) = "" Then Exit Sub
    If CCur(TxtAGasto(4)) > CCur(LblAgasto(6)) Then
        MsgBox "Monto del Gasto sobrepasa Total de la Factura", _
        vbCritical + vbOKOnly, "Monto Errado"
        With TxtAGasto(4)
            .SelStart = 0
            .SelLength = Len(.Text)
            .SetFocus
            End With
        Exit Sub
    End If
    
    If Check1(0).Value = 0 And Index = 4 And TxtAGasto(4) <> "" Then
    'And Option1(1).Value = True
        Option1(0).Visible = False
        GridGastos(2).Visible = True
        fraAgasto(4).Visible = True
        cmdHelp(2).BackColor = &H80000004
        TxtAGasto(4) = Format(TxtAGasto(4), "#,##0.00")
        With GridGastos(2)  'Configura el diseño del Grid
            .TextArray(0) = "Codigo": .ColWidth(0) = 700
            .TextArray(1) = "Propietario": .ColWidth(1) = 1350
            .TextArray(2) = "Monto"
            Call centra_titulo(GridGastos(2))
           .ColAlignment(0) = flexAlignCenterCenter
           .ColAlignment(1) = flexAlignCenterCenter
            Call RtnGrid: .Refresh
        End With
        
    End If
    End Sub

    '-----------------------------------------------------------------------
    Private Sub TxtAGasto_KeyPress(Index As Integer, KeyAscii As Integer) '-
    '-----------------------------------------------------------------------
    '
    KeyAscii = Asc(UCase(Chr(KeyAscii))) 'CONVIERTE TODO MAYUSCULAS
    Select Case Index 'VALIDA LA ENTRADA DE DATOS
    '---------------------------------------------
        Case 0
    '   ------------------------------------------
            Call Validacion(KeyAscii, "1234567890")
    
        Case 1
    '   ------------------------------------------
            
            
        Case 4
    '   ------------------------------------------
            If KeyAscii = 46 Then KeyAscii = 44 'CONVIERTE COMA A PUNTO
            Call Validacion(KeyAscii, "-1234567890,")
        
        Case 5   'Cuadro de Texto Buscar
    '   ------------------------------------------
            If OptTipoFactura(4) Or OptTipoFactura(7) Then   'N° FACTURA
                Call Validacion(KeyAscii, "0123456789")
            ElseIf OptTipoFactura(5) Then   'Monto
                If KeyAscii = 46 Then KeyAscii = 44 'CONVIERTE COMA A PUNTO
                Call Validacion(KeyAscii, "1234567890,")
            End If
            If KeyAscii = 13 Then Call Buscar_Factura
    End Select
    '
    If KeyAscii = 13 Then   'AL PRESIONAR {ENTER}
    '
        Select Case Index
    '   -----------------------------------------------
    '
            Case 0  'BUSCA CODIGO/DESCRIPCION DE GASTO
    '       -------------------------------------------
                If TxtAGasto(0) = "" Then TxtAGasto(1).SetFocus: Exit Sub
                Call RtnBuscaTGasto(0)
    '
            Case 1  'BUSCA CODIGO/DESCRIPCION DE GASTO
    '       ------------------------------------------
                If TxtAGasto(1) = "" Then
                    TxtAGasto(0).SetFocus
                ElseIf TxtAGasto(0) = "" Then
                    Call RtnBuscaTGasto(1)
                End If
    '
            Case 4  'DA FORMATO AL MONTO/PASA EL FOCO
    '       -----------------------------------------
                If Not IsNumeric(TxtAGasto(4)) Or Not IsNumeric(LblAgasto(6)) Then Exit Sub
                If CCur(TxtAGasto(4)) > CCur(LblAgasto(6)) Then
    '
                    MsgBox "Monto del Gasto sobrepasa Total de la Factura", _
                    vbCritical + vbOKOnly, "Monto Errado"
                    With TxtAGasto(4)
    '
                        .SelStart = 0
                        .SelLength = Len(.Text)
                        .SetFocus
    '
                    End With
                    Exit Sub
                End If
                Call RtnGrid
                TxtAGasto(4) = Format(TxtAGasto(4), "#,##0.00")
                Combo1(0) = Combo1(0).List(Month(Date) - 1)
                Combo1(0).SetFocus
    '   ------------------------------------------------------
        End Select
    '
    End If
    '
    End Sub

    '---------------------------------------------------------------------------------------------
    '   Rutina:     RtnGuardar
    '
    '   Rutina que guarda cada registro según la distribución del gassto: comun /
    '   No común. Por Alícuota o Partes Iguales. Para Propietarios Determinados.
    '   Montos Específicos. De acuerdo a la selección del usuario
    '---------------------------------------------------------------------------------------------
    Private Sub RtnGuardar() '
    'vairables locales
    Dim datPeriodo1 As Date, HoraA As Date
    Dim curMonto As Currency, curFC As Currency
    Dim errLocal As Long
    Dim rstlocal As ADODB.Recordset
    '
    On Error GoTo salir
    '
    curMonto = CCur(TxtAGasto(4))
    datPeriodo1 = "01/" & Combo1(0) & "/" & Combo1(1)
    '
    'Aqui actualiza el vector factura cancelada y viene de una edicion de factura canelada
    If stEditFC Then
    
        curFC = curMonto
        For I = 0 To UBound(vecFC)
        
            If vecFC(I, 1) <> 0 Then
30            If vecFC(I, 1) < curFC Then
                    cnn.Execute "INSERT INTO ChequeDetalle(IDCheque,Cargado,CodInm,CodGasto,Det" _
                    & "alle,Monto,Clave,Ndoc) IN '" & gcPath & "\sac.mdb' VALUES (" & vecFC(I, 0) _
                    & ",'" & datPeriodo1 & "','" & gcCodInm & "','" & TxtAGasto(0) & "','" & _
                    TxtAGasto(1) & "','" & vecFC(I, 1) & "','" & vecFC(I, 2) & "','" & DtcAGasto(0) & "')"
                    '
                    'si es cuenta de fondo agrega el movimiento al movfondo
                    cnn.Execute "INSERT INTO MovFondo (CodGasto,Fecha,Tipo,Periodo,Concepto,Deb" _
                    & "e,Haber) VALUES ('" & TxtAGasto(0) & "',Date(),'CH','" & datPeriodo1 & "'" _
                    & ",'" & TxtAGasto(1) & " CH#" & vecFC(I, 0) & "','" & vecFC(I, 1) & "',0)"
                    curFC = curFC - vecFC(I, 1)
                    vecFC(I, 1) = 0
                    I = I + 1
                    GoTo 30
                    '
                Else
                    cnn.Execute "INSERT INTO ChequeDetalle(IDCheque,Cargado,CodInm,CodGasto,Det" _
                    & "alle,Monto,Clave,Ndoc) IN '" & gcPath & "\sac.mdb' VALUES (" & vecFC(I, 0) & _
                    ",'" & datPeriodo1 & "','" & gcCodInm & "','" & TxtAGasto(0) & "','" & _
                    TxtAGasto(1) & "','" & curFC & "','" & vecFC(I, 2) & "','" & DtcAGasto(0) & "')"
                    '
                    If booFondo(TxtAGasto(0)) Then
                        If Len(TxtAGasto(1) & " CH#" & vecFC(I, 0)) > 50 Then
                            MsgBox "Resuma el campo descripción.", vbInformation, App.ProductName
                        End If
                    'si es cuenta de fondo agrega el movimiento al movfondo
                        cnn.Execute "INSERT INTO MovFondo (CodGasto,Fecha,Tipo,Periodo,Concepto,Deb" _
                        & "e,Haber) VALUES ('" & TxtAGasto(0) & "',Date(),'CH','" & datPeriodo1 & "'" _
                        & ",'" & TxtAGasto(1) & " CH#" & vecFC(I, 0) & "','" & curFC & "',0)"
                    End If
                    vecFC(I, 1) = vecFC(I, 1) - curFC
                    '
                End If
                Exit For
            End If
        Next
    '
    If blnRet Then Exit Sub
    End If
    
    'consulta si ya se ha guardado el mismo gasto en el mismo período
    Set rstlocal = New ADODB.Recordset
    '
    Set rstlocal = cnn.Execute("SELECT * FROM Cargado WHERE CodGasto='" & TxtAGasto(0) & _
    "' AND Periodo =#" & Format(datPeriodo1, "mm/dd/yyyy") & "# AND Monto = " _
    & Replace(curMonto, ",", "."))
    '
    If Not rstlocal.EOF And Not rstlocal.BOF Then
        MsgBox "Ya ha sido cargado un gasto similar al mismo período", vbInformation, Me.Caption
        'registrar en la bitacora
        Call rtnBitacora("Registro Similar " & TxtAGasto(0) & "|" & datPeriodo1 & "|" & _
        Format(curMonto, "#,##0.00"))
    End If
    rstlocal.Close
    Set rstlocal = Nothing
    HoraA = Time()
    If booFondo(TxtAGasto(0)) = True Then 'Si el gasto afecta el Fondo de Reserva
        'Agrega el cargado-----------------------
        cnn.Execute "INSERT INTO Cargado (Ndoc,CodGasto,Detalle,Monto,Fecha,Hora,Usuario,Period" _
        & "o) VALUES('" & DtcAGasto(0) & "','" & TxtAGasto(0) & "','" & TxtAGasto(1) & "','" & _
        curMonto & "',Date(),'" & HoraA & "' ,'" & gcUsuario & "','" & datPeriodo1 & "');"
    Else
        cnn.Execute "INSERT INTO Cargado (Ndoc,CodGasto,Detalle,Periodo,Monto,Fecha,Hora,Usuari" _
        & "o) VALUES('" & DtcAGasto(0) & "','" & TxtAGasto(0) & "','" & TxtAGasto(1) & "','" & _
        datPeriodo1 & "','" & curMonto & "',Date(),'" & HoraA & "' ,'" & gcUsuario & "')"
        If RstFacturas!Tipo = "**" Then GoTo 20
        If TxtAGasto(0) = "100025" Or TxtAGasto(0) = "101025" Or TxtAGasto(0) = "102025" Then GoTo 20
    '
        If Check1(0).Value = 0 Then  'Gasto {NoComun}
        '
            If Option1(0).Value = True Then 'Todos los Propietarios
        '
                If Check1(1).Value = 1 Then  'Aplicado por alícuota
                    Call RtnAnexarReg("TRUE", HoraA)
        '
                ElseIf Check1(1).Value = 0 Then 'Aplicado en partes iguales
                    Call RtnAnexarReg("FALSE", HoraA)
        '
                End If
        '
            ElseIf Option1(1).Value = True Then 'Propietarios Seleccionados
        '
                With GridGastos(2)
        '       ----------------------|
                For I = 1 To .Rows - 1
                    cnn.Execute "INSERT INTO GastoNoComun (CodApto,CodGasto,Concepto,Monto,Periodo," _
                    & "Fecha,Hora,Usuario) VALUES ('" & LisAGasto(1).List(I - 1) & "','" _
                    & TxtAGasto(0) & "','" & TxtAGasto(1) & "','" & .TextMatrix(I, 2) & "','" _
                    & datPeriodo1 & "',Date(),'" & HoraA & "' ,'" & gcUsuario & "')"
                Next
        '       ---------------------|
                End With
        '
            End If
        '
        ElseIf Check1(0) Then   'Gasto Comun
            'Guarda el registro en la Tabla AsignaGasto
            cnn.Execute "INSERT INTO AsignaGasto (Ndoc,Cargado,CodGasto,Descripcion,Comun, " _
            & "Alicuota,Monto,Usuario,Fecha,Hora) VALUES ('" & DtcAGasto(0) & "','01/" _
            & Combo1(0) & "/" & Combo1(1) & "','" & TxtAGasto(0) & "','" & TxtAGasto(1) & _
            "'," & Check1(0) & "," & Check1(1) & ",'" & curMonto & "','" & gcUsuario _
            & "',Date() ,'" & HoraA & "')"
        End If
    End If
salir:
    If Err.Number = 0 Then
20    Call Agregar_Gasto
    Else
        MsgBox Err.Description, vbCritical, "Error " & Err.Number
    End If
    End Sub
    
    '---------------------------------------------------------------------------------------------
    '   Rutina:     RtnAnexarReg
    '
    '   Entrada:    Variable boolLlamada. Por alicuota=True sino False
    '
    '   Agrega Nuevos registros a la tabla GastosNoComunes, según parametros
    '   enviados por el usuario,{alícuota / partes iguales}
    '---------------------------------------------------------------------------------------------
    Private Sub RtnAnexarReg(BooLlamada As Boolean, Hora As Date) '-
    'variable locales
    Dim strMonto As String  'variables locales
    Dim curMonto As Currency
    Dim datPeri As Date
    Set AdogastosTran = New ADODB.Recordset
    datPeri = "01/" & Combo1(0) & "/" & Combo1(1)
    curMonto = CCur(TxtAGasto(4))
    '------------------|
    With AdogastosTran
        .Open "SELECT * FROM Propietarios WHERE Codigo <> 'U" & gcCodInm & "' ORDER BY Codigo", _
        cnn, adOpenKeyset, adLockOptimistic
        .MoveFirst
        If BooLlamada Then
            strMonto = "('" & curMonto & "' * Alicuota / 100)"
        Else
            strMonto = Replace((curMonto / .RecordCount), ",", ".")
        End If
        cnn.Execute "INSERT INTO GastoNoComun (CodApto,CodGasto,Concepto,Monto,Periodo,Fecha,Ho" _
        & "ra,Usuario) SELECT Codigo,'" & TxtAGasto(0) & "','" & TxtAGasto(1) & "'" _
        & "," & strMonto & ",'" & datPeri & "',Date(),'" & Hora & "','" _
        & gcUsuario & "' FROM Propietarios WHERE Codigo <> 'U" & gcCodInm & "'"
    End With
    '----------------|
    AdogastosTran.Close
    Set AdogastosTran = Nothing
    End Sub

    '-------------------------------------------------------------------------------------------------
    Private Sub RtnGrid() 'RUTINA QUE DISTRIBUYE EL GRID PROP.SELECCIONADOS
    '-------------------------------------------------------------------------------------------------
    'variables locales
    Dim CurPorcion As Currency
    Dim ByFila As Integer
    Dim rstProp As ADODB.Recordset
    '
    If Check1(0).Value = 0 And TxtAGasto(4) <> "" Then
        
        If Option1(1).Value = True Then
        
            With GridGastos(2)
            CurPorcion = CLng(TxtAGasto(4) / LisAGasto(1).ListCount)
            For I = 0 To (LisAGasto(1).ListCount - 1)
            'Llena el Grid con los datos de la lista
                ByFila = I + 1
                If ByFila >= .Rows Then .AddItem ("")
                .TextMatrix(ByFila, 0) = RTrim(Left(LisAGasto(1).List(I), 6))
                .TextMatrix(ByFila, 1) = Right(LisAGasto(1).List(I), _
                Len(LisAGasto(1).List(I)) - 0)
                .TextMatrix(I + 1, 2) = Format(CurPorcion, "#,##0.00 ")
            Next
            End With
        
        Else
            Set rstProp = New ADODB.Recordset
            rstProp.CursorLocation = adUseClient
            rstProp.Open "Propietarios", cnn, adOpenKeyset, adLockReadOnly, admcdtable
            rstProp.Sort = "Codigo"
            rstProp.Filter = "Codigo <> 'U" & gcCodInm & "'"
            Call rtnLimpiar_Grid(GridGastos(2))
            If Not (rstProp.EOF And rstProp.BOF) Then
                rstProp.MoveFirst
                 CurPorcion = TxtAGasto(4) / (rstProp.RecordCount)
                Do
                    With GridGastos(2)
                        ByFila = ByFila + 1
                        If Check1(1).Value = vbChecked Then CurPorcion = _
                        TxtAGasto(4) * rstProp!Alicuota / 100
                        If ByFila >= .Rows Then .AddItem ("")
                        .TextMatrix(ByFila, 0) = rstProp!Codigo
                        .TextMatrix(ByFila, 1) = rstProp!Nombre
                        .TextMatrix(ByFila, 2) = Format(CurPorcion, "#,##0.00 ")
                        rstProp.MoveNext
                    End With
                Loop Until rstProp.EOF
                
            End If
            rstProp.Close
            Set rstProp = Nothing
        End If
    End If
    'configura la presentacion del grid en pantalla
    With GridGastos(2)
        'AltoFila = .RowHeight(1)
        .Height = (.Rows + 1) * .RowHeight(1)
        
        fraAgasto(4).Height = .Height + 100 + .Top
        If fraAgasto(4).Height > (fraAgasto(0).Height - fraAgasto(4).Top) Then
            fraAgasto(4).Height = fraAgasto(0).Height - fraAgasto(4).Top - 100
            .Height = fraAgasto(4).Height - .Top - 100
            If .RowHeight(1) * (.Rows + 1) > .Height Then
                fraAgasto(4).Left = cmdHelp(1).Left + 100
                fraAgasto(4).Width = TxtAGasto(4).Left + TxtAGasto(4).Width - cmdHelp(1).Left
                
                .Width = fraAgasto(4).Width - 200
'                .Left = 200
            End If
            
        End If
    End With
    
    End Sub



    '---------------------------------------------------------------------------------------------
    '   Rutina:     rtn_modificar
    '
    '   Permite modificar la distribucón del gasto de una factura ya asignada
    '---------------------------------------------------------------------------------------------
    Private Sub rtn_modificar()
    'variables locales                      '*
    Dim datCargado As Date                  '*
    Dim strSQL As String                    '*
    Dim RstCheque As New ADODB.Recordset    '*
    Dim numArchivo As Integer               '*
    Dim I As Integer, j As Integer          '*
    Dim numCheq As Long                     '*
    Dim Reg As Long                         '*
    Dim Archivo As Detalle_Cheque           '*
    '----------------------------------------
    With GridGastos(0)
        .Col = 5
        LblAgasto(6) = 0
        
        cnn.BeginTrans  'Comienza el proceso por lotes
        j = 1
        For I = 1 To (.Rows - 1)
            .Row = I
            
            If .CellBackColor = Blanco Then
                '
              If .TextMatrix(.Row, 4) <> "" Then
                LblAgasto(6) = CCur(LblAgasto(6)) + CCur(.TextMatrix(.Row, 4))
                datCargado = Format("01-" & .TextMatrix(.Row, 5), "mm-dd-yyyy")
                
                If OptTipoFactura(8) Then   'factura cancelada
                    stEditFC = True
                    'Selecciona los datos generales del cheque (ID,Total,CLave)
                    strSQL = "SELECT IDCheque,Monto as Total,Clave  FROM ChequeDetalle WHERE Co" _
                    & "dGasto & Detalle & Cargado IN (SELECT CodGasto & Detalle & Periodo  FROM" _
                    & " Cargado IN '" & mcDatos & "' WHERE Ndoc='" & DtcAGasto(0) & "' AND CodG" _
                    & "asto='" & .TextMatrix(.Row, 0) & "' AND Periodo=#" & datCargado & "# AND" _
                    & " Monto=CCur('" & .TextMatrix(.Row, 4) & "')) AND Clave IN (SELECT Clave " _
                    & "FROM ChequeFactura WHERE Ndoc='" & DtcAGasto(0) & "') AND Ndoc='" & DtcAGasto(0) & "';"
                    '
                    RstCheque.Open strSQL, cnnConexion, adOpenStatic, adLockReadOnly, adCmdText
                    '
                    If Not RstCheque.EOF Or Not RstCheque.BOF Then
                        '
                        numFichero = FreeFile
                        RstCheque.MoveFirst
                        numCheq = RstCheque!IDCheque
                        
                        Open strArchivo For Random As numFichero Len = Len(Archivo)
                            
                            Do
                                
                                With Archivo
                                    .ID = RstCheque!IDCheque
                                    .Monto = Format(CCur(RstCheque!Total), "###0.00")
                                    .Clave = RstCheque!Clave
                                End With
                                Put #numFichero, j, Archivo
                                j = j + 1
                                RstCheque.MoveNext
                            Loop Until RstCheque.EOF
                        Close numFichero
                        
                    Else
                        
                        MsgBox "No se encuentra o está alterada la información del detalle de p" _
                        & "ago del que cancela esta facutra. Cancele esta operación y póngase e" _
                        & "n contacto con el administrador del sistema.", vbCritical, App.EXEName
                        RstCheque.Close
                        Set RstCheque = Nothing
                        If Dir(strArchivo) <> "" Then Kill strArchivo
                        Exit Sub
                        
                    End If
                    RstCheque.Close
                    Set RstCheque = Nothing
                    
                     If booFondo(.TextMatrix(.Row, 0)) Then
                    '
                        'efectua un reintegro en el movimiento del fondo
                        cnn.Execute "INSERT INTO MovFondo (CodGasto,Fecha,Tipo,Periodo,Concepto" _
                        & ",Debe,Haber) SELECT CodGasto,Date(),'NC','01/" & _
                        Format(Date, "MM/YYYY") & "',Concepto,0,Debe FROM MovFondo WHERE Concep" _
                        & "to LIKE '%#" & numCheq & "%' AND CodGasto='" & .TextMatrix(.Row, 0) _
                        & "'", Reg
                    '
                    End If
                    '
                    'elimina el detalle del cheque emitido
'                   '
                    strSQL = "DELETE * FROM ChequeDetalle IN '" & gcPath & "\sac.mdb' WHERE Co" _
                    & "dGasto & Detalle & Cargado IN (SELECT CodGasto & Detalle & Periodo  FROM" _
                    & " Cargado WHERE Ndoc='" & DtcAGasto(0) & "' AND CodGasto='" & _
                    .TextMatrix(.Row, 0) & "' AND Periodo=#" & datCargado & "# AND" _
                    & " Monto=CCur('" & .TextMatrix(.Row, 4) & "')) AND Clave IN (SELECT Clave " _
                    & "FROM ChequeFactura IN '" & gcPath & "\sac.mdb' WHERE Ndoc='" & DtcAGasto(0) & "') AND Ndoc='" & DtcAGasto(0) & "';"
                    
                    cnn.Execute strSQL, Reg
                    '
                    End If
                '
                
                '
                'Elimina el registro de TDF Cargado
                cnn.Execute "DELETE FROM Cargado WHERE Ndoc='" & DtcAGasto(0) & "' AND CodGasto" _
                & "='" & .TextMatrix(.Row, 0) & "' AND Periodo=#" & datCargado & "# and Monto=C" _
                & "cur('" & .TextMatrix(.Row, 4) & "')"
                '
                'Elimina registro si estubiera asignado como gasto no comun
                cnn.Execute "DELETE FROM GastoNoComun WHERE CodGasto & Concepto & Periodo & Fec" _
                & "ha & Hora IN (SELECT CodGasto & Detalle & Periodo & Fecha & Hora FROM Cargad" _
                & "o WHERE Ndoc='" & DtcAGasto(0) & "');"
                '
                'Elimina el registro de TDF AsignaGasto
                cnn.Execute "DELETE FROM AsignaGasto WHERE Ndoc='" & DtcAGasto(0) & "' AND " _
                & "CodGasto='" & .TextMatrix(.Row, 0) & "' AND Cargado=#" & datCargado & "# and" _
                & " Monto=Ccur('" & .TextMatrix(.Row, 4) & "')", Reg
                'Elimina el registro de ChequeDetalle si la factura esta cancelada
                '
            End If
            '
          End If
          '
          '
        Next
        
        LblAgasto(6) = _
        Format(IIf(LblAgasto(6) > CCur(LblAgasto(5)), LblAgasto(5), LblAgasto(6)), "#,##0.00")
        
    End With
    '
    If LblAgasto(6) <> 0 Then    'existen gastos sin facturar
    '
        
        Call RtnEstado(5, Toolbar1)
        SSTab1.TabEnabled(1) = False
        TxtAGasto(0).Enabled = True
        For I = 0 To 2
            fraAgasto(I).Enabled = True
        Next
        TxtAGasto(0).SetFocus
        '
        With GridGastos(0)  'Elimina las filas seleccionadas
            .Col = 5
            X = .Rows - 1
            Do
                .Row = X
                If .CellBackColor = Blanco Then
                    If .Row <> 1 Then
                        .Rows = .Rows - 1
                        .Row = .Rows - 1
                    Else
                        For K = 0 To .Cols - 1
                            .Col = K
                            .Text = ""
                            If .Col = 2 Or .Col = 3 Then Set .CellPicture = Nothing
                        Next
                    End If
                End If
                X = X - 1
                '
            Loop Until X = 0
            '
        End With
        If OptTipoFactura(8) Then Call Cargar_Matriz
    Else
    
        cnn.RollbackTrans
        Call rtnBitacora(LoadResString(538) & DtcAGasto(0) & " FALLIDO")
        MsgBox "Imposible editar el cargado de esta factura" & vbCrLf & "" _
        & "", vbCritical, App.ProductName
        
    End If
    '
    End Sub

    
    '---------------------------------------------------------------------------------------------
    '   Function:   Validar_Guardar
    '
    '   Función que verifica los datos necesarios para procesar el registro
    '   periodo facturado, campos nulos, etc. Si todo es correcto devuelve
    '   un valor false, si encuenta un error devuelve True
    '---------------------------------------------------------------------------------------------
    Private Function Validar_Guardar() As Boolean
    'variables locales
    Dim strMensaje As String
    Dim datCargado As Date
    '------------------------
    datCargado = "01-" + Combo1(0) + "-" + Combo1(1)
    '
    If Combo1(0) = "" Or Combo1(1) = "" Then 'Sale de la rutina si esta vacio
        strMensaje = "Verifique los datos del Período"
    ElseIf TxtAGasto(0) = "" Or TxtAGasto(1) = "" Then
        strMensaje = "Verifique el Código y/o la Descripción de la Cuenta"
    ElseIf TxtAGasto(4) = "" Then
        strMensaje = strMensaje + "Debe Introducir una cantidad" & vbCrLf
    'ElseIf datCargado <= datPeriodo Then
    '    strMensaje = "Imposible Asignar el Gasto a un Período Facturado...."
    End If
    '
    If strMensaje <> "" Then Validar_Guardar = MsgBox(strMensaje, vbCritical, "Error..")
    '
    End Function


    '---------------------------------------------------------------------------------------------
    '   Funcion:    Ajustes_Reintegros
    '
    '   Si La factura en cuestión corresponde con un mantenimiento o un
    '   servicio presupuestado efectua, si corresponde, el ajuste o reintegro
    '   respectivo, devuelve True, de lo contrario devuelve False
    '---------------------------------------------------------------------------------------------
    Private Function Ajustes_Reintegros() As Boolean
    '
    'Variables locales
    Dim strSQL$, strGasto$   'Cadena Sql, selección de registros
    Dim rstAR As ADODB.Recordset  'Conjunto de Registros seleccionados
    Dim datFacturado As Date    'mes al que estamos cargando la factura
    Dim curFacturado@, curTemp@
    '
    Set rstAR = New ADODB.Recordset
    datFacturado = "01/" & Combo1(0) & "/" & Combo1(1)
    datFacturado = Format(datFacturado, "mm/dd/yy")
    '
    strGasto = TxtAGasto(0)
    'selecciona gasto fijo con este código de gasto
    strSQL = "SELECT * FROM AsignaGasto WHERE CodGasto='" & strGasto & "' AND Carga" _
        & "do=#" & datFacturado & "# AND Fijo=True AND CodGasto Not In (Select CodGasto FROM TGastos WHERE Fondo=True);"
    rstAR.Open strSQL, cnn, adOpenKeyset, adLockOptimistic, adCmdText
    '
    If Not (rstAR.EOF And rstAR.BOF) Then 'existe coincidencia
    '
        Ajustes_Reintegros = True   '
        If rstAR("Ndoc") <> 0 Then  'ya fué aplicado una factura a este gasto
            MsgBox "Para este Período ya fué asignado la Factura N° " & _
            rstAR("Ndoc"), vbInformation, App.ProductName
        Else    'si no verifica el monto que se esta aplicanco
            'Actualiza el nº de factura del gasto
            'elimne esta línea para no actualizar la descripcion--
            ',Descripcion='" & TxtAGasto(1) & "'------------------
            '-----------------------------------------------------
            cnn.Execute "UPDATE AsignaGasto SET Ndoc='" & DtcAGasto(0) & _
            "' WHERE Codgasto ='" _
            & strGasto & "' AND Cargado=#" & datFacturado & _
            "# and Fijo=True;", n
            '-----------------------------------------------------
            'cnn.Execute "UPDATE Cargado SET Ndoc='" & DtcAGasto(0) & "' WHERE Codgasto ='" _
            & strGasto & "' AND Periodo=#" & datFacturado & "#;"
            curFacturado = CCur(TxtAGasto(4))
            '
            If curFacturado > rstAR!Monto Then
                curTemp = Format(curFacturado - rstAR!Monto, "#,##0.00")
                MsgBox "Debe hacer un Ajuste al gasto por Bs. " & curTemp
                TxtAGasto(4) = rstAR!Monto
            ElseIf curFacturado < rstAR!Monto Then
                curTemp = curFacturado - rstAR!Monto
                MsgBox "Debe hacer un reintegro al gasto por Bs. " & Format(curTemp, "#,##0.00")
                TxtAGasto(4) = Format(rstAR!Monto, "#,##0.00")
                
            End If
            '
            
            If stEditFC Then
                blnRet = True
                Call RtnGuardar
                blnRet = False
            End If
            cnn.Execute "INSERT INTO Cargado (Ndoc,CodGasto,Detalle,Periodo,Monto,Fecha,Hora,Usuari" _
            & "o) VALUES('" & DtcAGasto(0) & "','" & TxtAGasto(0) & "','" & TxtAGasto(1) & "',#" _
            & datFacturado & "#,'" & TxtAGasto(4) & "',Date(),Time() ,'" & gcUsuario & "');"
            Call Agregar_Gasto
            TxtAGasto(4) = Format(curTemp, "#,##0.00")
            
        End If
    Else
        'Ajustes_Reintegros = False
        If CDate("01/" & Combo1(0) & "/" & Combo1(1)) <= datPeriodo Then
            Ajustes_Reintegros = MsgBox("Período ya facturado...", vbExclamation, App.EXEName)
        End If
'        rstAR.Close
'        rstAR.Open "SELECT * FROM Tgastos WHERE CodGasto='" & strGasto & "' AND Fijo=True;", _
'        cnn, adOpenKeyset, adLockOptimistic
'        If rstAR.RecordCount > 0 Then
'            cnn.Execute "UPDATE Tgastos SET Fijo=False,Usuario='AJUSTADO' WHERE Codgasto='" _
'            & strGasto & "'"
'        End If
    End If
    rstAR.Close
    Set rstAR = Nothing
    '
    End Function

    '---------------------------------------------------------------------------------------------
    Private Sub Agregar_Gasto() '
    '---------------------------------------------------------------------------------------------
    'Variables locales
    Dim strMonto@
    '
    strMonto = CCur(TxtAGasto(4))
    LblAgasto(6) = Format(CCur(LblAgasto(6)) - strMonto, "#,##0.00")
    '           ------------------'envia el registro a la cuadrícula
    With GridGastos(0)
        .Row = .Rows - 1
        If .TextMatrix(.Row, 0) <> "" Then .AddItem (""): .Row = .Row + 1
        
        For I = 0 To 5
        
            Select Case Val(I)
                Case 0, 1
                    .TextMatrix(.Row, I) = TxtAGasto(I)
                    TxtAGasto(I) = ""
                Case 2, 3
                    .TextMatrix(.Row, I) = IIf(Check1(I - 2) = 1, "SI", "NO")
                    .Col = I
                    .CellAlignment = flexAlignCenterCenter
                    Check1(I - 2).Value = 0
                Case 4
                    .TextMatrix(.Row, I) = Format(strMonto, "#,##0.00 ")
                Case 5
                    .TextMatrix(.Row, I) = Combo1(0) + "-" + Combo1(1)
                    For j = 0 To 1: Combo1(j).ListIndex = -1
                    Next
                End Select
        Next
        TxtAGasto(4) = ""
        GridGastos(0).AddItem ("")
        TxtAGasto(0).SetFocus
        LisAGasto(1).Clear
        Option1(0).Value = True
    '
    End With
    '
    End Sub

    '---------------------------------------------------------------------------------------------
    Private Sub Buscar_Factura() '
    '---------------------------------------------------------------------------------------------
    'Variable Local
    Dim strCriterio$
    Dim rstFiltro As ADODB.Recordset
    '
    If OptTipoFactura(4) Then
        strCriterio = "Fact='" & TxtAGasto(5) & "'"
    ElseIf OptTipoFactura(5) Then
        strCriterio = "Total=" & CCur(TxtAGasto(5))
    ElseIf OptTipoFactura(6) Then
        strCriterio = "Detalle LIKE '*" & TxtAGasto(5) & "*'"
    ElseIf OptTipoFactura(7) Then
        strCriterio = "Ndoc='" & TxtAGasto(5) & "'"
    ElseIf OptTipoFactura(9) Then   'filtra por cheque pagado
        If TxtAGasto(5) = "" Then
            RstFacturas.Filter = ""
        Else
            Set rstFiltro = New ADODB.Recordset
            rstFiltro.Open "SELECT * FROM ChequeFactura WHERE IDCHEQUE=" & TxtAGasto(5), cnnConexion, adOpenKeyset, adLockOptimistic, adCmdText
            If Not rstFiltro.EOF And Not rstFiltro.BOF Then
                rstFiltro.MoveFirst
                Do
                    If strCriterio <> "" Then strCriterio = strCriterio & " or "
                    strCriterio = strCriterio & "Ndoc = '" & rstFiltro.Fields("Ndoc") & "'"
                    rstFiltro.MoveNext
                Loop Until rstFiltro.EOF
            End If
            If strCriterio <> "" Then
                RstFacturas.Filter = strCriterio
            Else
                RstFacturas.Filter = 0
            End If
            rstFiltro.Close
            Set rstFiltro = Nothing
            Exit Sub
        End If
    End If
    
    If strCriterio = "" Then Exit Sub
    '
    With RstFacturas
        If Not .EOF Then .MoveNext
        .Find strCriterio   'Primera Busqueda
        '
        If .EOF And Not .BOF Then
            .MoveFirst
            .Find strCriterio   'Segunda Busqueda
            If .EOF Then MsgBox "Información No encontrada", vbInformation, App.ProductName
        End If
        '
    End With
    '
    End Sub

    Private Sub TxtAGasto_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
    'Variables locales
    Dim intC%, StrP$
    '
    If Index = 1 Then
        If KeyCode = 8 Then KeyCode = 46
        If KeyCode = 46 Or TxtAGasto(1) = "" Then Exit Sub
        
        intC = TxtAGasto(1).SelStart
        If intC <= 0 Then Exit Sub
        StrP = Left(TxtAGasto(1), intC)
        If StrP = "" Then Exit Sub
        With AdoCuentas.Recordset
                .Find "Titulo Like '" & StrP & "*'"
                If .EOF = True Then
                    If .BOF = True Then Exit Sub
                    'Buca desde el principio
                    .MoveFirst
                    .Find "Titulo Like '" & StrP & "*'"
                    If Not .EOF Then GoTo Marcar
                Else
Marcar:        With TxtAGasto(1)
                        .Text = AdoCuentas.Recordset!Titulo
                        .SelStart = intC
                        .SelLength = Len(.Text)
                        .SetFocus
                    End With
                    '
                End If
                '
        End With
        '
    End If
    '
    End Sub

    '---------------------------------------------------------------------------------------------
    '   Rutina: Mostrar_Cargado
    '
    '   Muestra como fue distribuido el gasto de la factura señalada
    '---------------------------------------------------------------------------------------------
    Private Sub Mostrar_Cargado()
    '
    Set RstAsignado = New ADODB.Recordset
    
    Dim strSQL$
    
    strSQL = "SELECT CodGasto, Descripcion, Comun,Alicuota, Format(Monto,'#,##0.00') ,Cargado F" _
    & "ROM AsignaGasto WHERE Ndoc ='" & RstFacturas("Ndoc") & "' UNION ALL SELECT CodGasto,Deta" _
    & "lle,'','', Format(Monto,'#,##0.00'),Periodo FROM Cargado WHERE Ndoc =" _
    & "'" & RstFacturas("Ndoc") & "' AND  CodGasto & Detalle & Monto & Periodo  NOT IN (SELECT " _
    & "CodGasto & Descripcion & Monto & Cargado FROM AsignaGasto WHERE Ndoc ='" & _
    RstFacturas("Ndoc") & "') ORDER BY Cargado,CodGasto;"

    RstAsignado.Open strSQL, cnn, adOpenStatic, adLockReadOnly, adCmdText
    
    If Not RstAsignado.RecordCount > 0 Then
        
        RstAsignado.Close
        RstAsignado.Open "SELECT CodGasto,Detalle,'','', Format(Monto,'#,##0.00'),Periodo FROM " _
        & "Cargado WHERE Ndoc ='" & RstFacturas("Ndoc") & "' ORDER BY Periodo DESC;", cnn, _
        adOpenStatic, adLockReadOnly, adCmdText
        
    End If
    '
    Call rtnLimpiar_Grid(GridGastos(0))
    If RstAsignado.RecordCount > 0 Then
        With GridGastos(0)
            .Rows = RstAsignado.RecordCount + 1
            RstAsignado.MoveFirst
            Do Until RstAsignado.EOF
                .Row = RstAsignado.AbsolutePosition
                For K = 0 To 5
                    If Not K = 2 And Not K = 3 Then
                        If K = 5 Then
                            .TextMatrix(.Row, K) = UCase(Format(RstAsignado.Fields(K), "MMM-YYYY"))
                        Else
                            .TextMatrix(.Row, K) = RstAsignado.Fields(K)
                        End If
                        '.TextMatrix(.Row, k) = IIf(k = 5, UCase(Format(RstAsignado.Fields(k), "MMM-YYYY")), RstAsignado.Fields(k))
                        '.TextMatrix(.Row, k) = RstAsignado.Fields(k)
                    Else
                        .Col = K
                        Set GridGastos(0).CellPicture = IIf(RstAsignado.Fields(K) = "-1", _
                        ImgOK(1), ImgOK(0))
                        GridGastos(0).CellPictureAlignment = flexAlignCenterCenter
                    End If
                    If K = 5 Then
                        .Col = 5
                        If CDate("01-" & Left(GridGastos(0).Text, 3) & "-" & _
                        Right(GridGastos(0).Text, 4)) <= datPeriodo Then
                            .CellBackColor = Verde   'fondo verde
                        Else
                            .CellBackColor = Blanco    'fondo blanco
                        End If
                    End If
                Next
                RstAsignado.MoveNext
            Loop
        End With
    Else
        GridGastos(0).Rows = 2
        For I = 2 To 3
            GridGastos(0).Col = I
            Set GridGastos(0).CellPicture = Nothing
        Next
    End If
    GridGastos(0).Row = 0
    '
    End Sub

    '---------------------------------------------------------------------------------------------
    '   Rutina Cargar_Matriz
    '
    '   Carga en un vector la información del detalle del cheque
    '---------------------------------------------------------------------------------------------
    Private Sub Cargar_Matriz()
    '
    Dim numFichero%  'vairables locales
    Dim I%, m@
    Dim Temp, Archivo As Detalle_Cheque
    Dim Cta As String
    '
    numFichero = FreeFile
    X = -1
    '
    Open strArchivo For Random As numFichero Len = Len(Archivo)
        For I = 1 To LOF(numFichero) / Len(Archivo)
            Get #numFichero, I, Archivo
            If Temp <> Archivo.ID Then X = X + 1
            Temp = Archivo.ID
            m = m + Trim(Archivo.Monto)
            Cta = Trim(Archivo.Clave)
        Next
        ReDim vecFC(X, 2)
    Close numFichero
    '----------------
    If X = 0 Then
        vecFC(X, 0) = Temp
        vecFC(X, 1) = m
        vecFC(X, 2) = Cta
    Else
        numFichero = FreeFile
        X = 0: Temp = 0
        '
        Open strArchivo For Random As numFichero Len = Len(Archivo)
            For I = 1 To LOF(numFichero) / Len(Archivo)
                Get #numFichero, I, Archivo
                If Temp <> Archivo.ID Then
                    vecFC(X, 0) = Archivo.ID
                    vecFC(X, 1) = 0
                    vecFC(X, 2) = Trim(Archivo.Clave)
                    X = X + 1
                End If
                Temp = Archivo.ID
            Next
        Close numFichero
        '-------------
        numFichero = FreeFile
        '
        For X = 0 To UBound(vecFC)
        
            Open strArchivo For Random As numFichero Len = Len(Archivo)
                'Do 'Until EOF(numFichero)
                For I = 1 To LOF(numFichero) / Len(Archivo)
                    Get #numFichero, I, Archivo
                    If vecFC(X, 0) = Archivo.ID Then
                        vecFC(X, 1) = CCur(vecFC(X, 1)) + Val(Archivo.Monto)
                    End If
                    
                Next
            Close numFichero
        Next
    End If
    '-------------
    If Dir(strArchivo) <> "" Then Kill strArchivo
    End Sub


    Sub config_fom()
    With SSTab1
        .Height = Me.ScaleHeight - .Top
        .Width = Me.ScaleWidth - .Left
        fraAgasto(0).Left = .Left + 100
        fraAgasto(0).Height = .Height - fraAgasto(0).Top - 200
        fraAgasto(0).Width = SSTab1.Width - (fraAgasto(0).Left * 2)
        DataGrid1.Left = fraAgasto(0).Left + 100
        DataGrid1.Height = .Height - fraAgasto(5).Width - 100
        DataGrid1.Width = fraAgasto(0).Width - DataGrid1.Left
        fraAgasto(5).Left = DataGrid1.Left
        fraAgasto(3).Left = fraAgasto(5).Left + fraAgasto(5).Width + 100
        fraAgasto(6).Left = fraAgasto(5).Left + fraAgasto(5).Width + fraAgasto(3).Width + 200
        fraAgasto(5).Top = DataGrid1.Top + DataGrid1.Height + 100
        fraAgasto(6).Top = DataGrid1.Top + DataGrid1.Height + 100
        fraAgasto(3).Top = DataGrid1.Top + DataGrid1.Height + 100
        fraAgasto(8).Left = (fraAgasto(0).Width - fraAgasto(8).Width) / 2
        fraAgasto(2).Left = fraAgasto(8).Left
        fraAgasto(1).Left = fraAgasto(8).Left
        GridGastos(1).Left = fraAgasto(2).Left + TxtAGasto(0).Left
        fraAgasto(4).Left = TxtAGasto(2).Left + fraAgasto(2).Left
        fraAgasto(1).Height = fraAgasto(0).Height - fraAgasto(1).Top - 200
        GridGastos(0).Height = fraAgasto(1).Height - GridGastos(0).Top - 200
    End With
    End Sub
