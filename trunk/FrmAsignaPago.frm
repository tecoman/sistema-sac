VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form FrmAsignaPago 
   Caption         =   "Emisión de Pagos"
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
      TabIndex        =   37
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
            Object.ToolTipText     =   "Anular Cheque"
            Object.Tag             =   ""
            ImageIndex      =   9
         EndProperty
         BeginProperty Button10 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Enabled         =   0   'False
            Object.Visible         =   0   'False
            Key             =   "Edit"
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
            Object.Width           =   1e-4
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
   Begin TabDlg.SSTab STabAPago 
      Height          =   7110
      Left            =   180
      TabIndex        =   0
      Top             =   570
      Width           =   11625
      _ExtentX        =   20505
      _ExtentY        =   12541
      _Version        =   393216
      Tabs            =   4
      Tab             =   2
      TabsPerRow      =   4
      TabHeight       =   520
      TabCaption(0)   =   "&Datos Generales"
      TabPicture(0)   =   "FrmAsignaPago.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "fraApago(0)"
      Tab(0).Control(1)=   "fraApago(1)"
      Tab(0).Control(2)=   "Calendario"
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "Lista &Facturas"
      TabPicture(1)   =   "FrmAsignaPago.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "fraApago(4)"
      Tab(1).Control(1)=   "fraApago(5)"
      Tab(1).Control(2)=   "fraApago(8)"
      Tab(1).Control(3)=   "DataGrid1"
      Tab(1).ControlCount=   4
      TabCaption(2)   =   "Lista &Pagos"
      TabPicture(2)   =   "FrmAsignaPago.frx":0038
      Tab(2).ControlEnabled=   -1  'True
      Tab(2).Control(0)=   "imgPago"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "fraApago(11)"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "MshgAPago(3)"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "Adodc1(0)"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).Control(4)=   "Adodc1(1)"
      Tab(2).Control(4).Enabled=   0   'False
      Tab(2).Control(5)=   "MshgAPago(4)"
      Tab(2).Control(5).Enabled=   0   'False
      Tab(2).Control(6)=   "fraApago(3)"
      Tab(2).Control(6).Enabled=   0   'False
      Tab(2).Control(7)=   "fraApago(7)"
      Tab(2).Control(7).Enabled=   0   'False
      Tab(2).ControlCount=   8
      TabCaption(3)   =   "Consecutivos"
      TabPicture(3)   =   "FrmAsignaPago.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "fraApago(10)"
      Tab(3).Control(1)=   "MshgAPago(6)"
      Tab(3).Control(2)=   "MshgAPago(7)"
      Tab(3).ControlCount=   3
      Begin MSComCtl2.MonthView Calendario 
         Height          =   2370
         Left            =   -70035
         TabIndex        =   88
         Top             =   4365
         Visible         =   0   'False
         Width           =   2625
         _ExtentX        =   4630
         _ExtentY        =   4180
         _Version        =   393216
         ForeColor       =   -2147483646
         BackColor       =   -2147483636
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         StartOfWeek     =   51838977
         TitleBackColor  =   -2147483646
         TitleForeColor  =   -2147483639
         CurrentDate     =   37938
      End
      Begin MSDataGridLib.DataGrid DataGrid1 
         Height          =   4215
         Left            =   -74715
         TabIndex        =   83
         Top             =   960
         Width           =   10935
         _ExtentX        =   19288
         _ExtentY        =   7435
         _Version        =   393216
         AllowUpdate     =   0   'False
         AllowArrows     =   0   'False
         BorderStyle     =   0
         ColumnHeaders   =   -1  'True
         HeadLines       =   1
         RowHeight       =   15
         RowDividerStyle =   4
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
         Caption         =   "LISTA DE FACTURAS ASIGNADAS INMUEBLE"
         ColumnCount     =   6
         BeginProperty Column00 
            DataField       =   "Ndoc"
            Caption         =   "N.Doc."
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
         BeginProperty Column05 
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
            ScrollBars      =   2
            AllowRowSizing  =   0   'False
            AllowSizing     =   0   'False
            BeginProperty Column00 
               Alignment       =   2
               DividerStyle    =   5
               ColumnAllowSizing=   0   'False
               ColumnWidth     =   915,024
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   870,236
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   870,236
            EndProperty
            BeginProperty Column03 
               ColumnWidth     =   5190,236
            EndProperty
            BeginProperty Column04 
               Alignment       =   1
               DividerStyle    =   5
               ColumnWidth     =   1275,024
            EndProperty
            BeginProperty Column05 
               Alignment       =   2
               ColumnWidth     =   1214,929
            EndProperty
         EndProperty
      End
      Begin VB.Frame fraApago 
         Caption         =   "Filtrar:"
         Height          =   1785
         Index           =   10
         Left            =   -74625
         TabIndex        =   64
         Top             =   5040
         Width           =   11070
         Begin VB.TextBox txtAPago 
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
            Height          =   315
            Index           =   17
            Left            =   8400
            TabIndex        =   116
            Top             =   615
            Width           =   1290
         End
         Begin VB.TextBox txtAPago 
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
            Height          =   315
            Index           =   16
            Left            =   8400
            TabIndex        =   85
            Top             =   240
            Width           =   1290
         End
         Begin VB.TextBox txtAPago 
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "#,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   8202
               SubFormatType   =   1
            EndProperty
            Height          =   315
            Index           =   15
            Left            =   5850
            TabIndex        =   82
            Top             =   1380
            Width           =   1290
         End
         Begin VB.TextBox txtAPago 
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "#,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   8202
               SubFormatType   =   1
            EndProperty
            Height          =   315
            Index           =   14
            Left            =   5850
            TabIndex        =   81
            Top             =   990
            Width           =   1290
         End
         Begin VB.TextBox txtAPago 
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
            Height          =   315
            Index           =   13
            Left            =   5850
            TabIndex        =   80
            Top             =   615
            Width           =   1290
         End
         Begin VB.TextBox txtAPago 
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "#,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   8202
               SubFormatType   =   1
            EndProperty
            Height          =   315
            Index           =   12
            Left            =   1065
            TabIndex        =   65
            Top             =   1380
            Width           =   3210
         End
         Begin MSDataListLib.DataCombo DtcAPago 
            Height          =   315
            Index           =   8
            Left            =   1065
            TabIndex        =   66
            ToolTipText     =   "Banco"
            Top             =   270
            Width           =   3210
            _ExtentX        =   5662
            _ExtentY        =   556
            _Version        =   393216
            Style           =   2
            ListField       =   "NombreBanco"
            BoundColumn     =   "NumCuenta"
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo DtcAPago 
            Height          =   315
            Index           =   9
            Left            =   1065
            TabIndex        =   67
            ToolTipText     =   "Número de Cuenta"
            Top             =   645
            Width           =   3210
            _ExtentX        =   5662
            _ExtentY        =   556
            _Version        =   393216
            Style           =   2
            ListField       =   "NumCuenta"
            BoundColumn     =   "NombreBanco"
            Text            =   ""
         End
         Begin MSComCtl2.DTPicker dtpApago 
            Height          =   315
            Index           =   5
            Left            =   1065
            TabIndex        =   68
            Top             =   1020
            Width           =   1460
            _ExtentX        =   2566
            _ExtentY        =   556
            _Version        =   393216
            CalendarTitleBackColor=   -2147483646
            CalendarTitleForeColor=   -2147483639
            CheckBox        =   -1  'True
            Format          =   51838977
            CurrentDate     =   37540
         End
         Begin MSComCtl2.DTPicker dtpApago 
            Height          =   315
            Index           =   6
            Left            =   2820
            TabIndex        =   69
            Top             =   1020
            Width           =   1460
            _ExtentX        =   2566
            _ExtentY        =   556
            _Version        =   393216
            CalendarTitleBackColor=   -2147483646
            CalendarTitleForeColor=   -2147483639
            CheckBox        =   -1  'True
            Format          =   51838977
            CurrentDate     =   37540
         End
         Begin MSDataListLib.DataCombo DtcAPago 
            Height          =   315
            Index           =   0
            Left            =   5865
            TabIndex        =   79
            ToolTipText     =   "Cargado"
            Top             =   240
            Width           =   1275
            _ExtentX        =   2249
            _ExtentY        =   556
            _Version        =   393216
            Style           =   2
            ListField       =   "P"
            BoundColumn     =   "Cargado"
            Text            =   ""
         End
         Begin VB.Label LblAPago 
            Caption         =   "Doc. Nº:"
            Height          =   255
            Index           =   39
            Left            =   7545
            TabIndex        =   115
            Top             =   645
            Width           =   750
         End
         Begin VB.Label LblAPago 
            Caption         =   "Cheque #:"
            Height          =   255
            Index           =   38
            Left            =   7545
            TabIndex        =   84
            Top             =   270
            Width           =   750
         End
         Begin VB.Label LblAPago 
            Caption         =   "Código del Gasto:"
            Height          =   255
            Index           =   37
            Left            =   4520
            TabIndex        =   78
            Top             =   1410
            Width           =   1545
         End
         Begin VB.Label LblAPago 
            Caption         =   "Cód. Inmueble:"
            Height          =   255
            Index           =   27
            Left            =   4520
            TabIndex        =   77
            Top             =   1020
            Width           =   1185
         End
         Begin VB.Label LblAPago 
            Caption         =   "Cheque Monto:"
            Height          =   255
            Index           =   26
            Left            =   4520
            TabIndex        =   76
            Top             =   645
            Width           =   1185
         End
         Begin VB.Label LblAPago 
            Caption         =   "Cargado:"
            Height          =   255
            Index           =   25
            Left            =   4520
            TabIndex        =   75
            Top             =   270
            Width           =   705
         End
         Begin VB.Label LblAPago 
            Caption         =   "Entre:"
            Height          =   255
            Index           =   28
            Left            =   255
            TabIndex        =   74
            Top             =   1050
            Width           =   700
         End
         Begin VB.Label LblAPago 
            Caption         =   "Cuenta:"
            Height          =   255
            Index           =   29
            Left            =   255
            TabIndex        =   73
            Top             =   675
            Width           =   700
         End
         Begin VB.Label LblAPago 
            Caption         =   "Banco:"
            Height          =   255
            Index           =   35
            Left            =   255
            TabIndex        =   72
            Top             =   300
            Width           =   700
         End
         Begin VB.Label LblAPago 
            Caption         =   "Y"
            Height          =   255
            Index           =   24
            Left            =   2600
            TabIndex        =   71
            Top             =   1080
            Width           =   105
         End
         Begin VB.Label LblAPago 
            Caption         =   "Proveedor:"
            Height          =   255
            Index           =   36
            Left            =   240
            TabIndex        =   70
            Top             =   1440
            Width           =   825
         End
      End
      Begin VB.Frame fraApago 
         Height          =   435
         Index           =   8
         Left            =   -74715
         TabIndex        =   52
         Top             =   405
         Width           =   2265
         Begin VB.Label LblAPago 
            Alignment       =   1  'Right Justify
            BackColor       =   &H80000002&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000005&
            Height          =   270
            Index           =   22
            Left            =   810
            TabIndex        =   54
            Top             =   120
            Width           =   1410
         End
         Begin VB.Label LblAPago 
            BackColor       =   &H80000002&
            Caption         =   "Total:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000005&
            Height          =   270
            Index           =   21
            Left            =   30
            TabIndex        =   53
            Top             =   120
            Width           =   780
         End
      End
      Begin VB.Frame fraApago 
         Caption         =   "Impresión: "
         Height          =   1785
         Index           =   7
         Left            =   255
         TabIndex        =   49
         Top             =   5085
         Width           =   2280
         Begin VB.Frame fraApago 
            BorderStyle     =   0  'None
            Caption         =   "Frame1"
            Height          =   615
            Index           =   9
            Left            =   120
            TabIndex        =   89
            Top             =   1080
            Width           =   1935
            Begin VB.OptionButton Option1 
               Caption         =   "Pantalla"
               Height          =   285
               Index           =   5
               Left            =   600
               TabIndex        =   91
               Top             =   360
               Width           =   1335
            End
            Begin VB.OptionButton Option1 
               Caption         =   "Impresora"
               Height          =   285
               Index           =   4
               Left            =   600
               TabIndex        =   90
               Top             =   0
               Value           =   -1  'True
               Width           =   1095
            End
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Selección"
            Height          =   285
            Index           =   3
            Left            =   300
            TabIndex        =   51
            Top             =   570
            Width           =   1095
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Todo"
            Height          =   285
            Index           =   2
            Left            =   300
            TabIndex        =   50
            Top             =   270
            Value           =   -1  'True
            Width           =   735
         End
         Begin VB.Line Line1 
            BorderColor     =   &H00FFFFFF&
            Index           =   1
            X1              =   240
            X2              =   2040
            Y1              =   960
            Y2              =   960
         End
         Begin VB.Line Line1 
            BorderColor     =   &H00808080&
            BorderWidth     =   2
            Index           =   0
            X1              =   240
            X2              =   2040
            Y1              =   960
            Y2              =   960
         End
      End
      Begin VB.Frame fraApago 
         Caption         =   "Filtrar:"
         Height          =   1785
         Index           =   3
         Left            =   2640
         TabIndex        =   38
         Top             =   5085
         Width           =   4515
         Begin VB.OptionButton optApago 
            Caption         =   "Anulado"
            Height          =   315
            Index           =   2
            Left            =   1020
            Style           =   1  'Graphical
            TabIndex        =   55
            Top             =   1395
            Width           =   1005
         End
         Begin VB.OptionButton optApago 
            Caption         =   "Pendiente"
            Height          =   315
            Index           =   1
            Left            =   3225
            Style           =   1  'Graphical
            TabIndex        =   47
            Top             =   1395
            Value           =   -1  'True
            Width           =   1005
         End
         Begin VB.OptionButton optApago 
            Caption         =   "Impreso"
            Height          =   315
            Index           =   0
            Left            =   2115
            Style           =   1  'Graphical
            TabIndex        =   46
            Top             =   1395
            Width           =   1005
         End
         Begin MSDataListLib.DataCombo DtcAPago 
            Height          =   315
            Index           =   5
            Left            =   1065
            TabIndex        =   39
            ToolTipText     =   "Banco"
            Top             =   270
            Width           =   3210
            _ExtentX        =   5662
            _ExtentY        =   556
            _Version        =   393216
            ListField       =   "NombreBanco"
            BoundColumn     =   "NumCuenta"
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo DtcAPago 
            Height          =   315
            Index           =   6
            Left            =   1065
            TabIndex        =   41
            ToolTipText     =   "Número de Cuenta"
            Top             =   645
            Width           =   3210
            _ExtentX        =   5662
            _ExtentY        =   556
            _Version        =   393216
            ListField       =   "NumCuenta"
            BoundColumn     =   "NombreBanco"
            Text            =   ""
         End
         Begin MSComCtl2.DTPicker dtpApago 
            Height          =   315
            Index           =   1
            Left            =   1065
            TabIndex        =   43
            Top             =   1020
            Width           =   1515
            _ExtentX        =   2672
            _ExtentY        =   556
            _Version        =   393216
            CalendarTitleBackColor=   -2147483646
            CalendarTitleForeColor=   -2147483639
            CheckBox        =   -1  'True
            Format          =   51838977
            CurrentDate     =   37540
         End
         Begin VB.Label LblAPago 
            Caption         =   "Estado:"
            Height          =   255
            Index           =   9
            Left            =   255
            TabIndex        =   45
            Top             =   1425
            Width           =   705
         End
         Begin VB.Label LblAPago 
            Caption         =   "Fecha:"
            Height          =   255
            Index           =   8
            Left            =   255
            TabIndex        =   44
            Top             =   1050
            Width           =   700
         End
         Begin VB.Label LblAPago 
            Caption         =   "Cuenta:"
            Height          =   255
            Index           =   6
            Left            =   255
            TabIndex        =   42
            Top             =   675
            Width           =   700
         End
         Begin VB.Label LblAPago 
            Caption         =   "Banco:"
            Height          =   255
            Index           =   7
            Left            =   255
            TabIndex        =   40
            Top             =   300
            Width           =   700
         End
      End
      Begin VB.Frame fraApago 
         Caption         =   "Buscar por:"
         Height          =   1400
         Index           =   5
         Left            =   -71250
         TabIndex        =   8
         Top             =   5400
         Width           =   3720
         Begin VB.CommandButton CmdRegistrar 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Index           =   0
            Left            =   1950
            Style           =   1  'Graphical
            TabIndex        =   12
            Top             =   735
            Width           =   1515
         End
         Begin VB.TextBox txtAPago 
            Height          =   315
            Index           =   0
            Left            =   1950
            TabIndex        =   11
            Top             =   300
            Width           =   1515
         End
         Begin VB.OptionButton optBusca 
            Caption         =   "N° Documento"
            Height          =   285
            Index           =   0
            Left            =   330
            TabIndex        =   10
            Top             =   315
            Value           =   -1  'True
            Width           =   1860
         End
         Begin VB.OptionButton optBusca 
            Caption         =   "N° Factura"
            Height          =   285
            Index           =   1
            Left            =   330
            TabIndex        =   9
            Top             =   750
            Width           =   1155
         End
      End
      Begin VB.Frame fraApago 
         Caption         =   "Filtrar por:"
         Height          =   1400
         Index           =   4
         Left            =   -74670
         TabIndex        =   4
         Top             =   5400
         Width           =   3315
         Begin MSDataListLib.DataCombo DtcAPago 
            Height          =   315
            Index           =   4
            Left            =   1470
            TabIndex        =   7
            Top             =   720
            Width           =   1395
            _ExtentX        =   2461
            _ExtentY        =   556
            _Version        =   393216
            ListField       =   "CodProv"
            Text            =   ""
         End
         Begin VB.OptionButton optFiltro 
            Caption         =   "P&roveedor"
            Height          =   285
            Index           =   1
            Left            =   330
            TabIndex        =   6
            Top             =   735
            Width           =   1155
         End
         Begin VB.OptionButton optFiltro 
            Caption         =   "&Ninguno"
            Height          =   285
            Index           =   0
            Left            =   330
            TabIndex        =   5
            Top             =   315
            Value           =   -1  'True
            Width           =   1860
         End
      End
      Begin VB.Frame fraApago 
         Caption         =   "Asigna Pago"
         Enabled         =   0   'False
         Height          =   3225
         Index           =   1
         Left            =   -74730
         TabIndex        =   3
         Top             =   3690
         Width           =   10995
         Begin VB.CommandButton cmd 
            Height          =   255
            Left            =   8355
            Picture         =   "FrmAsignaPago.frx":0070
            Style           =   1  'Graphical
            TabIndex        =   87
            Top             =   720
            Width           =   225
         End
         Begin VB.TextBox txtAPago 
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
            Height          =   315
            Index           =   5
            Left            =   9525
            Locked          =   -1  'True
            TabIndex        =   36
            Text            =   "0"
            Top             =   1065
            Width           =   1245
         End
         Begin VB.TextBox txtAPago 
            Height          =   315
            Index           =   4
            Left            =   3420
            MaxLength       =   240
            MultiLine       =   -1  'True
            TabIndex        =   30
            ToolTipText     =   "Descripción Pago"
            Top             =   1065
            Width           =   5205
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Cuenta Inmueble"
            Height          =   510
            Index           =   0
            Left            =   200
            TabIndex        =   21
            Top             =   315
            Value           =   -1  'True
            Width           =   1450
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Cuenta Administradora"
            Height          =   540
            Index           =   1
            Left            =   200
            TabIndex        =   20
            Top             =   885
            Width           =   1450
         End
         Begin MSDataListLib.DataCombo DtcAPago 
            Height          =   315
            Index           =   2
            Left            =   7305
            TabIndex        =   22
            ToolTipText     =   "Número de Cuenta"
            Top             =   315
            Width           =   3465
            _ExtentX        =   6112
            _ExtentY        =   556
            _Version        =   393216
            ListField       =   "NumCuenta"
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo DtcAPago 
            Height          =   315
            Index           =   1
            Left            =   3420
            TabIndex        =   31
            ToolTipText     =   "Banco"
            Top             =   315
            Width           =   3180
            _ExtentX        =   5609
            _ExtentY        =   556
            _Version        =   393216
            ListField       =   "NombreBanco"
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo DtcAPago 
            Height          =   315
            Index           =   3
            Left            =   3420
            TabIndex        =   32
            ToolTipText     =   "Código Chequera"
            Top             =   690
            Width           =   675
            _ExtentX        =   1191
            _ExtentY        =   556
            _Version        =   393216
            ListField       =   "IDChequera"
            Text            =   ""
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid MshgAPago 
            Height          =   1530
            Index           =   1
            Left            =   195
            TabIndex        =   34
            Tag             =   "1300|6500|1200|1300|0|0"
            Top             =   1590
            Width           =   10530
            _ExtentX        =   18574
            _ExtentY        =   2699
            _Version        =   393216
            Cols            =   7
            FixedCols       =   0
            BackColorFixed  =   -2147483646
            ForeColorFixed  =   -2147483639
            BackColorSel    =   -2147483636
            BackColorBkg    =   -2147483636
            GridColor       =   -2147483636
            FocusRect       =   2
            HighLight       =   2
            GridLinesFixed  =   0
            ScrollBars      =   2
            BorderStyle     =   0
            Appearance      =   0
            MousePointer    =   99
            FormatString    =   "Código|Descripción|Cargado|Monto|"
            BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MouseIcon       =   "FrmAsignaPago.frx":01BA
            _NumberOfBands  =   1
            _Band(0).Cols   =   7
            _Band(0).TextStyleBand=   0
            _Band(0).TextStyleHeader=   1
         End
         Begin MSMask.MaskEdBox MskFecha 
            Bindings        =   "FrmAsignaPago.frx":031C
            DataSource      =   "DEFrmFactura"
            Height          =   315
            Left            =   7305
            TabIndex        =   86
            TabStop         =   0   'False
            Top             =   690
            Width           =   1305
            _ExtentX        =   2302
            _ExtentY        =   556
            _Version        =   393216
            MaxLength       =   10
            Format          =   "dd/MM/yyyy"
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin VB.Label LblAPago 
            Alignment       =   1  'Right Justify
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Index           =   32
            Left            =   9525
            TabIndex        =   35
            Top             =   690
            Width           =   1245
         End
         Begin VB.Label LblAPago 
            Alignment       =   1  'Right Justify
            Caption         =   "Cheque Bs."
            Height          =   255
            Index           =   5
            Left            =   8595
            TabIndex        =   33
            Top             =   1095
            Width           =   900
         End
         Begin VB.Label LblAPago 
            Caption         =   "Cuenta:"
            Height          =   255
            Index           =   3
            Left            =   6660
            TabIndex        =   29
            Top             =   360
            Width           =   705
         End
         Begin VB.Label LblAPago 
            Caption         =   "Chequera:"
            Height          =   255
            Index           =   31
            Left            =   2445
            TabIndex        =   28
            Top             =   720
            Width           =   900
         End
         Begin VB.Label LblAPago 
            Alignment       =   1  'Right Justify
            Caption         =   "Difencia:"
            Height          =   255
            Index           =   33
            Left            =   8595
            TabIndex        =   27
            Top             =   720
            Width           =   900
         End
         Begin VB.Label LblAPago 
            Alignment       =   1  'Right Justify
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Index           =   34
            Left            =   5220
            TabIndex        =   26
            ToolTipText     =   "Número Cheque"
            Top             =   690
            Width           =   1380
         End
         Begin VB.Label LblAPago 
            Alignment       =   1  'Right Justify
            Caption         =   "Nro.Cheque:"
            Height          =   255
            Index           =   19
            Left            =   4200
            TabIndex        =   25
            Top             =   720
            Width           =   960
         End
         Begin VB.Label LblAPago 
            Caption         =   "Fecha:"
            Height          =   255
            Index           =   20
            Left            =   6660
            TabIndex        =   24
            Top             =   735
            Width           =   630
         End
         Begin VB.Label LblAPago 
            Caption         =   "Banco:"
            Height          =   255
            Index           =   30
            Left            =   2445
            TabIndex        =   23
            Top             =   345
            Width           =   900
         End
         Begin VB.Label LblAPago 
            Caption         =   "Descripcion:"
            Height          =   255
            Index           =   2
            Left            =   2445
            TabIndex        =   19
            Top             =   1095
            Width           =   900
         End
      End
      Begin VB.Frame fraApago 
         Height          =   3165
         Index           =   0
         Left            =   -74730
         TabIndex        =   1
         Top             =   405
         Width           =   10995
         Begin VB.PictureBox Picture1 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   2985
            Left            =   6630
            LinkItem        =   "101"
            Picture         =   "FrmAsignaPago.frx":033E
            ScaleHeight     =   2985
            ScaleWidth      =   105
            TabIndex        =   117
            Tag             =   "0"
            Top             =   135
            Width           =   105
         End
         Begin VB.TextBox txtAPago 
            Height          =   315
            Index           =   2
            Left            =   1200
            TabIndex        =   18
            Top             =   625
            Width           =   5340
         End
         Begin VB.TextBox txtAPago 
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
            Height          =   315
            Index           =   3
            Left            =   1200
            Locked          =   -1  'True
            TabIndex        =   16
            Text            =   "0"
            Top             =   945
            Width           =   1605
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid MshgAPago 
            Height          =   1770
            Index           =   2
            Left            =   210
            TabIndex        =   14
            Tag             =   "900|900|3000|1150"
            Top             =   1305
            Width           =   6315
            _ExtentX        =   11139
            _ExtentY        =   3122
            _Version        =   393216
            Cols            =   4
            FixedCols       =   0
            BackColorFixed  =   -2147483646
            ForeColorFixed  =   -2147483639
            BackColorSel    =   65280
            ForeColorSel    =   -2147483646
            BackColorBkg    =   -2147483636
            GridColor       =   -2147483636
            AllowBigSelection=   0   'False
            FocusRect       =   0
            HighLight       =   0
            GridLinesFixed  =   0
            ScrollBars      =   2
            SelectionMode   =   1
            MousePointer    =   99
            GridLineWidthFixed=   1
            FormatString    =   "Nº Doc|Fact.|Descripción|Monto"
            BandDisplay     =   1
            RowSizingMode   =   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            FontWidthFixed  =   0
            MouseIcon       =   "FrmAsignaPago.frx":0455
            _NumberOfBands  =   1
            _Band(0).Cols   =   4
            _Band(0).GridLineWidthBand=   1
            _Band(0).TextStyleBand=   0
            _Band(0).TextStyleHeader=   3
         End
         Begin VB.TextBox txtAPago 
            Height          =   315
            Index           =   1
            Left            =   1200
            Locked          =   -1  'True
            TabIndex        =   13
            Top             =   300
            Width           =   5340
         End
         Begin TabDlg.SSTab tabDetalle 
            Height          =   2970
            Left            =   6795
            TabIndex        =   56
            Top             =   135
            Width           =   4140
            _ExtentX        =   7303
            _ExtentY        =   5239
            _Version        =   393216
            Tabs            =   2
            TabHeight       =   520
            WordWrap        =   0   'False
            ShowFocusRect   =   0   'False
            BackColor       =   -2147483646
            ForeColor       =   -2147483646
            TabCaption(0)   =   "Gastos"
            TabPicture(0)   =   "FrmAsignaPago.frx":05B7
            Tab(0).ControlEnabled=   -1  'True
            Tab(0).Control(0)=   "MshgAPago(0)"
            Tab(0).Control(0).Enabled=   0   'False
            Tab(0).ControlCount=   1
            TabCaption(1)   =   "Abonos"
            TabPicture(1)   =   "FrmAsignaPago.frx":05D3
            Tab(1).ControlEnabled=   0   'False
            Tab(1).Control(0)=   "LblAPago(23)"
            Tab(1).Control(1)=   "MshgAPago(5)"
            Tab(1).Control(2)=   "txtAPago(11)"
            Tab(1).ControlCount=   3
            Begin VB.TextBox txtAPago 
               Alignment       =   1  'Right Justify
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000002&
               Height          =   315
               Index           =   11
               Left            =   -72345
               Locked          =   -1  'True
               TabIndex        =   60
               ToolTipText     =   "Descripción Pago"
               Top             =   2565
               Width           =   1425
            End
            Begin MSHierarchicalFlexGridLib.MSHFlexGrid MshgAPago 
               Height          =   2340
               Index           =   0
               Left            =   75
               TabIndex        =   57
               Tag             =   "700|0|750|1100|1100|0|0"
               Top             =   435
               Width           =   3975
               _ExtentX        =   7011
               _ExtentY        =   4128
               _Version        =   393216
               Cols            =   7
               FixedCols       =   0
               BackColorFixed  =   -2147483646
               ForeColorFixed  =   -2147483639
               BackColorSel    =   65280
               ForeColorSel    =   -2147483646
               BackColorBkg    =   -2147483636
               GridColor       =   -2147483636
               AllowBigSelection=   0   'False
               Enabled         =   0   'False
               FocusRect       =   0
               HighLight       =   2
               GridLinesFixed  =   0
               ScrollBars      =   2
               SelectionMode   =   1
               BorderStyle     =   0
               MousePointer    =   99
               GridLineWidthFixed=   1
               FormatString    =   "Código|Descripción|Cargado|Monto|Resto|"
               BandDisplay     =   1
               RowSizingMode   =   1
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   -1  'True
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               FontWidthFixed  =   0
               MouseIcon       =   "FrmAsignaPago.frx":05EF
               _NumberOfBands  =   1
               _Band(0).Cols   =   7
               _Band(0).GridLineWidthBand=   1
               _Band(0).TextStyleBand=   0
               _Band(0).TextStyleHeader=   3
            End
            Begin MSHierarchicalFlexGridLib.MSHFlexGrid MshgAPago 
               Height          =   1995
               Index           =   5
               Left            =   -74925
               TabIndex        =   58
               Tag             =   "1000|1200|1500"
               Top             =   450
               Width           =   3975
               _ExtentX        =   7011
               _ExtentY        =   3519
               _Version        =   393216
               Cols            =   4
               FixedCols       =   0
               BackColorFixed  =   -2147483646
               ForeColorFixed  =   -2147483639
               BackColorSel    =   65280
               ForeColorSel    =   -2147483646
               BackColorBkg    =   -2147483636
               WordWrap        =   -1  'True
               AllowBigSelection=   0   'False
               Enabled         =   0   'False
               FocusRect       =   0
               HighLight       =   2
               GridLinesFixed  =   0
               ScrollBars      =   2
               SelectionMode   =   1
               BorderStyle     =   0
               MousePointer    =   99
               GridLineWidthFixed=   1
               FormatString    =   "^Cheque |^Fecha |>Monto"
               BandDisplay     =   1
               RowSizingMode   =   1
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   -1  'True
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               FontWidthFixed  =   0
               MouseIcon       =   "FrmAsignaPago.frx":0751
               _NumberOfBands  =   1
               _Band(0).Cols   =   4
               _Band(0).GridLinesBand=   3
               _Band(0).GridLineWidthBand=   1
               _Band(0).TextStyleBand=   0
               _Band(0).TextStyleHeader=   3
            End
            Begin VB.Label LblAPago 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Total abonos:"
               Height          =   255
               Index           =   23
               Left            =   -73425
               TabIndex        =   59
               Top             =   2640
               Width           =   1005
            End
         End
         Begin VB.Label LblAPago 
            Caption         =   "&Beneficiario:"
            Height          =   285
            Index           =   4
            Left            =   195
            TabIndex        =   17
            Top             =   660
            Width           =   1005
         End
         Begin VB.Label LblAPago 
            Caption         =   "Total:"
            Height          =   285
            Index           =   1
            Left            =   210
            TabIndex        =   15
            Top             =   1020
            Width           =   1005
         End
         Begin VB.Label LblAPago 
            Caption         =   "Proveedor:"
            Height          =   285
            Index           =   0
            Left            =   195
            TabIndex        =   2
            Top             =   315
            Width           =   1005
         End
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MshgAPago 
         Height          =   1600
         Index           =   4
         Left            =   225
         TabIndex        =   48
         Tag             =   "1200|900|5800|1400|1400"
         Top             =   3315
         Width           =   11085
         _ExtentX        =   19553
         _ExtentY        =   2831
         _Version        =   393216
         BackColor       =   16777215
         Cols            =   5
         FixedCols       =   0
         BackColorFixed  =   -2147483646
         ForeColorFixed  =   -2147483639
         BackColorSel    =   65280
         ForeColorSel    =   -2147483646
         GridColor       =   -2147483633
         AllowBigSelection=   0   'False
         FocusRect       =   0
         HighLight       =   2
         GridLinesFixed  =   1
         ScrollBars      =   2
         SelectionMode   =   1
         BorderStyle     =   0
         MousePointer    =   99
         GridLineWidthFixed=   1
         FormatString    =   "Cod - Cuenta|Cargado|Detalle|Débito|Crédito"
         BandDisplay     =   1
         RowSizingMode   =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FontWidthFixed  =   0
         MouseIcon       =   "FrmAsignaPago.frx":08B3
         _NumberOfBands  =   1
         _Band(0).Cols   =   5
         _Band(0).GridLineWidthBand=   1
         _Band(0).TextStyleBand=   0
         _Band(0).TextStyleHeader=   0
      End
      Begin MSAdodcLib.Adodc Adodc1 
         Height          =   330
         Index           =   1
         Left            =   3555
         Top             =   1185
         Visible         =   0   'False
         Width           =   2445
         _ExtentX        =   4313
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
         DataSourceName  =   "Sac"
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   ""
         Caption         =   "ChequeDetalle"
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
      Begin MSAdodcLib.Adodc Adodc1 
         Height          =   330
         Index           =   0
         Left            =   990
         Top             =   1185
         Visible         =   0   'False
         Width           =   2445
         _ExtentX        =   4313
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
         DataSourceName  =   "Sac"
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   ""
         Caption         =   "Cheque"
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
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MshgAPago 
         Height          =   2760
         Index           =   3
         Left            =   225
         TabIndex        =   61
         Tag             =   "1200|1200|1000|3500|0|0|11000|0"
         Top             =   555
         Width           =   11085
         _ExtentX        =   19553
         _ExtentY        =   4868
         _Version        =   393216
         BackColor       =   16777215
         Cols            =   8
         FixedCols       =   0
         BackColorFixed  =   -2147483646
         ForeColorFixed  =   -2147483639
         BackColorSel    =   65280
         ForeColorSel    =   -2147483646
         GridColor       =   -2147483633
         AllowBigSelection=   0   'False
         FocusRect       =   0
         HighLight       =   2
         GridLinesFixed  =   1
         SelectionMode   =   1
         BorderStyle     =   0
         MousePointer    =   99
         GridLineWidthFixed=   1
         FormatString    =   "^Cheque |>Monto |^Fecha |Beneficiario |Banco |Cuenta |Concepto |Impreso"
         BandDisplay     =   1
         RowSizingMode   =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FontWidthFixed  =   0
         MouseIcon       =   "FrmAsignaPago.frx":0A15
         _NumberOfBands  =   1
         _Band(0).Cols   =   8
         _Band(0).GridLineWidthBand=   1
         _Band(0).TextStyleBand=   0
         _Band(0).TextStyleHeader=   0
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MshgAPago 
         Height          =   2760
         Index           =   6
         Left            =   -74775
         TabIndex        =   62
         Tag             =   "1000|1200|1000|3500|0|0|11000"
         Top             =   555
         Width           =   11085
         _ExtentX        =   19553
         _ExtentY        =   4868
         _Version        =   393216
         BackColor       =   16777215
         Cols            =   8
         FixedCols       =   0
         BackColorFixed  =   -2147483646
         ForeColorFixed  =   -2147483639
         BackColorSel    =   65280
         ForeColorSel    =   -2147483646
         GridColor       =   -2147483633
         AllowBigSelection=   0   'False
         FocusRect       =   0
         HighLight       =   2
         GridLinesFixed  =   1
         SelectionMode   =   1
         BorderStyle     =   0
         MousePointer    =   99
         GridLineWidthFixed=   1
         FormatString    =   "^Cheque |>Monto |^Fecha |Beneficiario |Banco |Cuenta |Concepto |Impreso"
         BandDisplay     =   1
         RowSizingMode   =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FontWidthFixed  =   0
         MouseIcon       =   "FrmAsignaPago.frx":0B77
         _NumberOfBands  =   1
         _Band(0).Cols   =   8
         _Band(0).GridLineWidthBand=   1
         _Band(0).TextStyleBand=   0
         _Band(0).TextStyleHeader=   0
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MshgAPago 
         Height          =   1600
         Index           =   7
         Left            =   -74775
         TabIndex        =   63
         Tag             =   "1300|1100|7100|1200"
         Top             =   3315
         Width           =   11085
         _ExtentX        =   19553
         _ExtentY        =   2831
         _Version        =   393216
         BackColor       =   16777215
         Cols            =   4
         FixedCols       =   0
         BackColorFixed  =   -2147483646
         ForeColorFixed  =   -2147483639
         BackColorSel    =   65280
         ForeColorSel    =   -2147483646
         GridColor       =   -2147483633
         AllowBigSelection=   0   'False
         FocusRect       =   0
         HighLight       =   2
         GridLinesFixed  =   1
         ScrollBars      =   2
         SelectionMode   =   1
         BorderStyle     =   0
         MousePointer    =   99
         GridLineWidthFixed=   1
         FormatString    =   "Cod.Cuenta|Cargado|Detalle|Total"
         BandDisplay     =   1
         RowSizingMode   =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FontWidthFixed  =   0
         MouseIcon       =   "FrmAsignaPago.frx":0CD9
         _NumberOfBands  =   1
         _Band(0).Cols   =   4
         _Band(0).GridLineWidthBand=   1
         _Band(0).TextStyleBand=   0
         _Band(0).TextStyleHeader=   0
      End
      Begin VB.Frame fraApago 
         Height          =   1785
         Index           =   11
         Left            =   7230
         TabIndex        =   92
         Top             =   5085
         Width           =   4065
         Begin MSComCtl2.FlatScrollBar scrApago 
            Height          =   1485
            Left            =   3750
            TabIndex        =   114
            Top             =   225
            Visible         =   0   'False
            Width           =   270
            _ExtentX        =   476
            _ExtentY        =   2619
            _Version        =   393216
            MousePointer    =   99
            MouseIcon       =   "FrmAsignaPago.frx":0E3B
            Appearance      =   2
            LargeChange     =   5
            Max             =   10
            Orientation     =   1245184
         End
         Begin VB.Frame fraApago 
            BorderStyle     =   0  'None
            Caption         =   "fraApago"
            Height          =   1545
            Index           =   6
            Left            =   30
            TabIndex        =   93
            Top             =   195
            Width           =   3990
            Begin VB.Frame fraApago 
               BorderStyle     =   0  'None
               Height          =   3615
               Index           =   2
               Left            =   120
               TabIndex        =   94
               Top             =   30
               Width           =   3990
               Begin VB.CheckBox chkPago 
                  Caption         =   "Monto:"
                  Height          =   210
                  Index           =   1
                  Left            =   240
                  TabIndex        =   102
                  Top             =   2565
                  Visible         =   0   'False
                  Width           =   1000
               End
               Begin VB.CheckBox chkPago 
                  Caption         =   "Fecha:"
                  Height          =   210
                  Index           =   0
                  Left            =   210
                  TabIndex        =   101
                  Top             =   1665
                  Visible         =   0   'False
                  Width           =   1000
               End
               Begin VB.TextBox txtAPago 
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
                  Enabled         =   0   'False
                  Height          =   315
                  Index           =   9
                  Left            =   2025
                  TabIndex        =   100
                  Text            =   "0,00"
                  Top             =   2490
                  Visible         =   0   'False
                  Width           =   1275
               End
               Begin VB.TextBox txtAPago 
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
                  Enabled         =   0   'False
                  Height          =   315
                  Index           =   10
                  Left            =   2010
                  TabIndex        =   99
                  Text            =   "0,00"
                  Top             =   2940
                  Visible         =   0   'False
                  Width           =   1275
               End
               Begin VB.TextBox txtAPago 
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
                  Height          =   315
                  Index           =   6
                  Left            =   1500
                  MaxLength       =   6
                  TabIndex        =   98
                  Top             =   15
                  Width           =   1050
               End
               Begin VB.TextBox txtAPago 
                  BeginProperty DataFormat 
                     Type            =   1
                     Format          =   "#,##0.00"
                     HaveTrueFalseNull=   0
                     FirstDayOfWeek  =   0
                     FirstWeekOfYear =   0
                     LCID            =   8202
                     SubFormatType   =   1
                  EndProperty
                  Height          =   315
                  Index           =   7
                  Left            =   1500
                  TabIndex        =   97
                  Top             =   338
                  Width           =   2325
               End
               Begin VB.TextBox txtAPago 
                  BeginProperty DataFormat 
                     Type            =   1
                     Format          =   "#,##0.00"
                     HaveTrueFalseNull=   0
                     FirstDayOfWeek  =   0
                     FirstWeekOfYear =   0
                     LCID            =   8202
                     SubFormatType   =   1
                  EndProperty
                  Height          =   315
                  Index           =   8
                  Left            =   1500
                  TabIndex        =   96
                  Top             =   705
                  Width           =   2325
               End
               Begin VB.CommandButton CmdRegistrar 
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   330
                  Index           =   1
                  Left            =   2595
                  Style           =   1  'Graphical
                  TabIndex        =   95
                  Top             =   0
                  Width           =   870
               End
               Begin MSComCtl2.DTPicker dtpApago 
                  Height          =   315
                  Index           =   2
                  Left            =   2040
                  TabIndex        =   103
                  Top             =   1620
                  Visible         =   0   'False
                  Width           =   1260
                  _ExtentX        =   2223
                  _ExtentY        =   556
                  _Version        =   393216
                  Enabled         =   0   'False
                  Format          =   51838977
                  CurrentDate     =   37403
               End
               Begin MSComCtl2.DTPicker dtpApago 
                  Height          =   315
                  Index           =   3
                  Left            =   2040
                  TabIndex        =   104
                  Top             =   2010
                  Visible         =   0   'False
                  Width           =   1260
                  _ExtentX        =   2223
                  _ExtentY        =   556
                  _Version        =   393216
                  Enabled         =   0   'False
                  Format          =   51838977
                  CurrentDate     =   37403
               End
               Begin VB.Label LblAPago 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Buscar Cheque N°:"
                  Height          =   210
                  Index           =   10
                  Left            =   45
                  TabIndex        =   113
                  Top             =   60
                  Width           =   1500
               End
               Begin VB.Label LblAPago 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Del Banco:"
                  Height          =   210
                  Index           =   11
                  Left            =   45
                  TabIndex        =   112
                  Top             =   390
                  Width           =   2055
               End
               Begin VB.Label LblAPago 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Beneficiario:"
                  Height          =   210
                  Index           =   12
                  Left            =   45
                  TabIndex        =   111
                  Top             =   757
                  Width           =   1500
               End
               Begin VB.Label LblAPago 
                  Appearance      =   0  'Flat
                  BackColor       =   &H00C0C0C0&
                  BackStyle       =   0  'Transparent
                  Caption         =   "Opciones Avanzadas »»»"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   -1  'True
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00800000&
                  Height          =   240
                  Index           =   13
                  Left            =   0
                  MouseIcon       =   "FrmAsignaPago.frx":0F9D
                  MousePointer    =   99  'Custom
                  TabIndex        =   110
                  Top             =   1170
                  Width           =   3495
               End
               Begin VB.Label LblAPago 
                  Alignment       =   1  'Right Justify
                  Caption         =   "entre el"
                  Height          =   210
                  Index           =   14
                  Left            =   1260
                  TabIndex        =   109
                  Top             =   1665
                  Visible         =   0   'False
                  Width           =   630
               End
               Begin VB.Label LblAPago 
                  Alignment       =   1  'Right Justify
                  Caption         =   "y Bs."
                  Height          =   210
                  Index           =   17
                  Left            =   1395
                  TabIndex        =   108
                  Top             =   3015
                  Visible         =   0   'False
                  Width           =   495
               End
               Begin VB.Label LblAPago 
                  Alignment       =   1  'Right Justify
                  Caption         =   "y el"
                  Height          =   210
                  Index           =   15
                  Left            =   1215
                  TabIndex        =   107
                  Top             =   2055
                  Visible         =   0   'False
                  Width           =   630
               End
               Begin VB.Label LblAPago 
                  Alignment       =   1  'Right Justify
                  Caption         =   "entre Bs."
                  Height          =   210
                  Index           =   16
                  Left            =   1260
                  TabIndex        =   106
                  Top             =   2565
                  Visible         =   0   'False
                  Width           =   630
               End
               Begin VB.Label LblAPago 
                  Appearance      =   0  'Flat
                  BackColor       =   &H80000005&
                  BackStyle       =   0  'Transparent
                  BorderStyle     =   1  'Fixed Single
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   -1  'True
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H80000008&
                  Height          =   2010
                  Index           =   18
                  Left            =   0
                  TabIndex        =   105
                  Top             =   1395
                  Visible         =   0   'False
                  Width           =   3495
               End
            End
         End
      End
      Begin VB.Image imgPago 
         Height          =   225
         Left            =   420
         Picture         =   "FrmAsignaPago.frx":10EF
         Top             =   2250
         Width           =   240
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
            Picture         =   "FrmAsignaPago.frx":1414
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmAsignaPago.frx":1596
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmAsignaPago.frx":1718
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmAsignaPago.frx":189A
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmAsignaPago.frx":1A1C
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmAsignaPago.frx":1B9E
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmAsignaPago.frx":1D20
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmAsignaPago.frx":1EA2
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmAsignaPago.frx":2024
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmAsignaPago.frx":21A6
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmAsignaPago.frx":2328
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmAsignaPago.frx":24AA
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "FrmAsignaPago"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
    '---------------------------------------------------------------------------------------------
    'SAC-SINAI TECH
    'Módulo Asignación de pagos
    'Impresión de cheques
    '---------------------------------------------------------------------------------------------
    Dim RstFacturasA As ADODB.Recordset, RstGastos As ADODB.Recordset, RstBco As ADODB.Recordset
    Dim rst As ADODB.Recordset, RstCheque As ADODB.Recordset
    Dim GtoCnn As ADODB.Connection
    Dim CurDiferencia() As Currency
    Dim VecVinculo() As Integer
    Dim intRow%, strCtvo$
    Dim cAletra As New clsNum2Let
    Dim blnEdit As Boolean, blnOtro As Boolean
    Const CHEQUE_BANESCO$ = "chq_banesco.rpt"
    Const CHEQUE_VENEZUELA$ = "chq_vzla.rpt"
    Const CHEQUE_PROVINCIAL$ = "chq_provincial.rpt"
    Const Consulta_Consecutivo$ = "QfConsecutivoCondominio"
    '---------------------------------------------------------------------------------------------
    Private Sub Calendario_DateClick(ByVal DateClicked As Date)
    MskFecha = DateClicked
    Calendario.Visible = False
    End Sub

    Private Sub Calendario_KeyPress(KeyAscii As Integer)
    '
    If KeyAscii = 13 Then
        MskFecha = Calendario.Value
        Calendario.Visible = False
    End If
    '
    End Sub

    Private Sub Calendario_KeyUp(KeyCode As Integer, Shift As Integer)
    'si presiona esc oculta el calendario
    If KeyCode = 27 Then Calendario.Visible = False
    End Sub

    Private Sub Calendario_LostFocus()
    Calendario.Visible = False
    End Sub

    '---------------------------------------------------------------------------------------------
    Private Sub chkPago_Click(Index As Integer) '
    '---------------------------------------------------------------------------------------------
    '
    Select Case Index
    '
        Case 0
    '   ---------------------
            If chkPago(0).Value = 1 Then
                dtpApago(2).Enabled = True
                dtpApago(3).Enabled = True
                dtpApago(2).Value = Date
                dtpApago(3).Value = Date
            Else
                dtpApago(2).Enabled = False
                dtpApago(3).Enabled = False
            End If
    '
        Case 1
    '   ---------------------
            If chkPago(1).Value = 1 Then
                txtAPago(9).Enabled = True
                txtAPago(10).Enabled = True
            Else
                txtAPago(9).Enabled = False
                txtAPago(10).Enabled = False
            End If
    '
    End Select
    '
    End Sub

    
    Private Sub cmd_Click()
    Calendario.Visible = Not Calendario.Visible
    If Calendario.Visible = True Then Calendario.SetFocus
    End Sub

    '---------------------------------------------------------------------------------------------
    Private Sub CmdRegistrar_Click(Index As Integer)    '
    '---------------------------------------------------------------------------------------------
    '
    MousePointer = vbHourglass
    '
    Select Case Index
    '
        Case 0  'Busca {Facutra/Documento}
    '   ---------------------
            With RstFacturasA
                .MoveFirst
                .Find IIf(OptBusca(0), "Ndoc='", "Fact='") & txtAPago(0) & "'"
                If .EOF Then MsgBox IIf(OptBusca(0), "Documento no registrado", "Factura no reg" _
                & "istrada"), vbExclamation, App.ProductName: .MoveFirst
            End With
            
        Case 1  'Buscar Cheque
    '   ---------------------
            Call rtnBusca_Cheque
            
'        Case 2  'Imprimir Cheque
'    '   ---------------------
'            Call RtnImprimir_Cheque(IIf(optApago(0), 1, 0))
            
    End Select
    MousePointer = vbDefault
    '
    End Sub


    Private Sub DataGrid1_KeyDown(KeyCode As Integer, Shift As Integer)
    'Debug.Print KeyCode, Shift
    End Sub

    Private Sub DataGrid1_KeyUp(KeyCode As Integer, Shift As Integer)
    'Debug.Print KeyCode, Shift
    If KeyCode = 17 Then
        If DataGrid1.SelBookmarks.Count > 1 Then Call rtnMultiSel
    End If
    End Sub
    
    Private Sub DataGrid1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 0 And Shift = 0 Then Call rtnMultiSel
    End Sub

    Private Sub DtcAPago_Change(Index As Integer)
    If Index = 9 And DtcAPago(9) <> "" Then Call rtnConsecutivos
    End Sub

    'rev.28/08/2002-------------------------------------------------------------------------------
    Private Sub DtcAPago_Click(Index As Integer, Area As Integer)
    '---------------------------------------------------------------------------------------------
    '
    If Area = 2 Then
    '
        Select Case Index   'SELECCION CONTROL
            Case 0  'Cargado
    '       ---------------------------
                If DtcAPago(8) <> "" Or DtcAPago(9) <> "" Then Call rtnConsecutivos
    '
            Case 1  'Nombre BCO.
    '       ---------------------------
                Call RtnBuscaCta("NombreBanco", DtcAPago(1), 1, 2)
                Call RtnChqrs
                
            Case 2  'NUMEROS DE CTA.BCO.
    '       ---------------------------
                Call RtnBuscaCta("NumCuenta", DtcAPago(2), 1, 2)
                Call RtnChqrs
            
            Case 3  'NUMERO DE CHEQUERA
    '       ---------------------------
                Call RtnConsecutivo 'Busca el n° de cheque consecutivo
            
            Case 4  'Filtra por codigo de proveedor
    '       ---------------------------
                RstFacturasA.Filter = ""
                Call optFiltro_Click(1)
            
            Case 5, 6   'Busca Cheques emitidos según parametros de busqueda
    '       ---------------------------
                If Index = 5 Then
                    DtcAPago(6) = DtcAPago(5).BoundText
                Else
                    DtcAPago(5) = DtcAPago(6).BoundText
                End If
                Call rtnSelOPT
                
            Case 8  'Cosnecutivos
    '       ---------------------------
                DtcAPago(0).Text = ""
                Call RtnBuscaCta("NombreBanco", DtcAPago(8), 8, 9)
            
            Case 9  'Consecutivo
    '       ---------------------------
                DtcAPago(0).Text = ""
                Call RtnBuscaCta("NumCuenta", DtcAPago(9), 8, 9)
    '
        End Select
    '
    End If
    '
    End Sub
    
    
    '---------------------------------------------------------------------------------------------
    Private Sub DtcAPago_KeyPress(Index As Integer, KeyAscii As Integer)    '
    '---------------------------------------------------------------------------------------------
    Select Case Index
        Case 1, 2, 3  'Inhabilita la modificación
            KeyAscii = 0
    End Select
    End Sub

    
    'Rev.28/08/2002 ---Rutina que traspasa la información de la factura seleccionada a la ficha---
    Private Sub DataGrid1_Click()    'datos generales------------------------------
    '---------------------------------------------------------------------------------------------
    '
    Select Case Index   'De acuerdo al Grid que llama al procedimiento
    '
        Case 0  'Grid Facturas Asignadas
    '   -----------------------------------------
            Call rtnLimpiar_Grid(MshgAPago(1))
            Call rtnMultiSel
    '
        Case 1  'Grid Cheques
    
    '   -----------------------------------------
    
    End Select
    '
    End Sub

    

    '---------------------------------------------------------------------------------------------
    Private Sub dtpApago_Change(Index As Integer)   '
    '---------------------------------------------------------------------------------------------
    '
    Select Case Index
        Case 0  'evalua la fecha del cheque
    '   ---------------------
                If dtpApago(Index).Value < Date Then
                    MsgBox "La fecha seleccionada es anterior a la fecha de hoy...", _
                    vbCritical, App.ProductName
                    dtpApago(Index).Value = Date
                End If
        Case 1
    '   ---------------------mshgapago(3).ColWidth(2)
            Call rtnSelOPT   'Rutina Busca Cheques Emitidos para la fecha seleccionada
                            'cumpliendo con los otros parametros seleccionados
        Case 5, 6
    '   ---------------------
            If DtcAPago(8) <> "" Or DtcAPago(9) <> "" Then Call rtnConsecutivos
            
    End Select
    
    '
    End Sub

    
    'REV.27/08/2002----------------Carga el Formulario en pantalla, configura el origen de datos--
    Private Sub Form_Load()
    '---------------------------------------------------------------------------------------------
    'variables locales
    Dim rstList As ADODB.Recordset
    Dim strSQL$
    '
    CmdRegistrar(1).Picture = LoadResPicture("Buscar", vbResIcon)
    CmdRegistrar(0).Picture = LoadResPicture("Buscar", vbResIcon)
    '
    'configura los control ADO
    Adodc1(0).ConnectionString = cnnOLEDB + gcPath + "\sac.mdb"
    Adodc1(0).CommandType = adCmdText
    Adodc1(0).Mode = adModeShareDenyNone
    Adodc1(1).ConnectionString = cnnOLEDB + gcPath + "\sac.mdb"
    Adodc1(1).CommandType = adCmdText
    Adodc1(1).Mode = adModeShareDenyNone
    'SELECT CodInm & ' ' &  CodGasto as Cuenta, *  FROM ChequeDetalle
    
    Set MshgAPago(0).FontFixed = LetraTitulo(LoadResString(527), 9, , True)
    Set MshgAPago(0).Font = LetraTitulo(LoadResString(528), 8)
    
    For I = 1 To 7
      Set MshgAPago(I).FontFixed = MshgAPago(0).FontFixed
      Set MshgAPago(I).Font = MshgAPago(0).Font
    Next I
    
    Set DataGrid1.HeadFont = LetraTitulo(LoadResString(527), 7.5, True)
    Set DataGrid1.Font = LetraTitulo(LoadResString(528), 8)
    '
    'Set cAletra = New Num2Let.clsNum2Let
    Set RstFacturasA = New ADODB.Recordset
    Set rstList = New ADODB.Recordset
    Set RstBco = New ADODB.Recordset
    Set GtoCnn = New ADODB.Connection
    '
    GtoCnn.CursorLocation = adUseClient
    GtoCnn.Open cnnOLEDB & mcDatos
    '
    'Selecciona información de cuentas y chequeras asociadas al inmueble
    If gnCta = CUENTA_INMUEBLE Then 'si es cuenta del propio inmueble
    
        Call rtnBusca_Cuenta(mcDatos)
        
    Else    'si es cuenta pote
    
        Option1(1).Value = True
        Option1(0).Enabled = False
        
    End If
    '
    If gcCodInm = sysCodInm Then    'Facturas asignadas de la cuenta pote
        strSQL = "SELECT DISTINCT Cpp.CodProv FROM (Caja INNER JOIN Inmueble ON Caja.CodigoCaja" _
        & "= Inmueble.Caja) INNER JOIN Cpp ON Inmueble.CodInm = Cpp.CodInm WHERE Caja.CodigoCaj" _
        & "a='99' AND Cpp.Estatus='ASIGNADO' ORDER BY Cpp.CodProv"
    Else    'Facturas asignadas por condominio
        strSQL = "SELECT DISTINCT Cpp.CodProv FROM Proveedores INNER JOIN Cpp ON Proveedores.Co" _
        & "digo = Cpp.CodProv WHERE Cpp.estatus = 'ASIGNADO' and Cpp.CodInm = '" & gcCodInm _
        & "'  ORDER BY Cpp.CodProv"
    End If
    '---------------------------------------------------------------------------------------------
    rstList.Open strSQL, cnnConexion, adOpenKeyset, adLockOptimistic, adCmdText
    Set DtcAPago(4).RowSource = rstList             'pendientes de pago
    '---------------------------------------------------------------------------------------------
    strSQL = ftnSQL(" ORDER BY Cpp.codProv, cpp.FRecep")
    '
    RstFacturasA.CursorLocation = adUseClient
    RstFacturasA.Open strSQL, cnnConexion, adOpenKeyset, adLockOptimistic, adCmdText
    Set DataGrid1.DataSource = RstFacturasA
    DataGrid1.Caption = DataGrid1.Caption + " " + gcCodInm
    'DataGrid1.ReBind  'Lista las facturas pendientes de pago
    '---------------------------------------------------------------------------------------------
    For I = 0 To 4: If Not I = 3 Then Call centra_titulo(MshgAPago(I), True)
    Next
    '
    MshgAPago(0).ColAlignment(0) = flexAlignCenterCenter
    MshgAPago(0).ColAlignment(2) = flexAlignCenterCenter
    MshgAPago(1).Width = fraApago(1).Width - (MshgAPago(1).Left * 2)
    MshgAPago(1).ColAlignment(1) = flexAlignLeftCenter
    MshgAPago(1).ColAlignment(0) = flexAlignCenterCenter
    MshgAPago(1).ColAlignment(2) = flexAlignCenterCenter
    MshgAPago(2).ColAlignment(0) = flexAlignCenterCenter
    MshgAPago(2).ColAlignment(1) = flexAlignCenterCenter
    MshgAPago(4).ColAlignment(0) = flexAlignCenterCenter
    MshgAPago(4).ColAlignment(1) = flexAlignCenterCenter
    MshgAPago(5).RowHeight(0) = 315
    STabAPago.TabEnabled(0) = False
    dtpApago(1).Value = Date
    If gcNivel > nuAdministrador Then cmd.Enabled = False: MskFecha.Enabled = False
    STabAPago.tab = 1
    Call centra_titulo(MshgAPago(3), True)
    'MshgAPago(1).Cols = MshgAPago(1).Cols + 1
    'MshgAPago(1).ColWidth(5) = 0
    '
    End Sub

    
    '---------------------------------------------------------------------------------------------
    '   Function:       ftnPrintCheque
    '
    '   Entradas:       (IntChq) Valor entero que determina Clave en la tabla Cheques;
    '                   (intT) Impresion / Reimpresion,curTotal: monto del cheque
    '
    '   Salida:         Valos String del Número de cheque impreso
    '
    '   Genera un consulta temporal con toda la información del cheque seleccionado
    '   llama la .dll que convierte una cantidad numérica en letras(monto del cheque)
    '   desencadena la impresión del cheque
    '---------------------------------------------------------------------------------------------
    Private Function ftnPrintCheque(intChq$, intT%, curTotal@) As String '
    'Variables locales
    Dim strSQL As String                    '*
    Dim strAletra As String                 '*
    Dim strReporte As String                '*
    Dim strFormula As String                '*
    Dim strMonto As String                  '*
    Dim strNumCheq As String                '*
    Dim errLocal As Long                    '*
    Dim rstMultiCheq As New ADODB.Recordset '*
    Dim rpReporte As ctlReport
    '--------------------------------------------
    MousePointer = vbHourglass
    ftnPrintCheque = intChq
    'strSql = "SELECT Sum(ChequeDetalle.Monto) AS Total, Format(Cheque.IDCheque,'000000') AS Num" _
    & "eral, Cheque.FechaCheque, Cheque.Beneficiario, Cheque.Banco, Cheque.Cuenta, Cheque.Conce" _
    & "pto, Cheque.Impreso, ChequeDetalle.Cargado, ChequeDetalle.CodInm, ChequeDetalle.CodGasto" _
    & ", ChequeDetalle.Detalle, ChequeDetalle.Monto FROM Cheque INNER JOIN ChequeDetalle ON Che" _
    & "que.Clave = ChequeDetalle.Clave WHERE Cheque.Clave='" & intChq & "' GROUP BY Cheque.Fech" _
    & "aCheque, Cheque.Beneficiario, Cheque.Banco, Cheque.Cuenta, Cheque.Concepto, Cheque.Impre" _
    & "so, ChequeDetalle.Cargado, ChequeDetalle.CodInm, ChequeDetalle.CodGasto, ChequeDetalle.D" _
    & "etalle, ChequeDetalle.Monto, Cheque.IDCheque;"
    strSQL = "SELECT ChequeDetalle.Monto AS Total, Format(Cheque.IDCheque,'000000') AS Numeral," _
    & "Cheque.FechaCheque, Cheque.Beneficiario, Cheque.Banco, Cheque.Cuenta, Cheque.Concepto, C" _
    & "heque.Impreso, ChequeDetalle.Cargado, ChequeDetalle.CodInm, ChequeDetalle.CodGasto, Cheq" _
    & "ueDetalle.Detalle, ChequeDetalle.Monto FROM Cheque INNER JOIN ChequeDetalle ON Cheque.Cl" _
    & "ave = ChequeDetalle.Clave WHERE Cheque.Clave='" & intChq & "'"

    '
    Call rtnGenerator(gcPath & "\sac.mdb", strSQL, "ChequeImpresion")
    'rstMultiCheq.Open "ChequeImpresion", cnnConexion, adOpenStatic, adLockReadOnly, adCmdTable
    cAletra.Moneda = "Bolívares"
    '
    'Busca otros cheques registrados que cancelen el mismo documento
    
    rstMultiCheq.Open "SELECT DISTINCT ChequeFactura.IDCheque, Sum(ChequeDetalle.Monto) AS TOTA" _
    & "L,Cheque.FechaCheque FROM (Cheque INNER JOIN ChequeDetalle ON Cheque.Clave = ChequeDe" _
    & "talle.Clave) INNER JOIN ChequeFactura ON Cheque.Clave = ChequeFactura.Clave GROUP BY Che" _
    & "queFactura.IDCheque, ChequeFactura.Ndoc, Cheque.FechaCheque,ChequeFactura.Clave HAVING (" _
    & "((ChequeFactura.Clave)<>'" & intChq & "') AND ((ChequeFactura.Ndoc) In (SELECT Ndoc FROM" _
    & " ChequeFactura WHERE Clave='" & intChq & "'))) ORDER BY ChequeFactura.IDCheque DESC;", _
    cnnConexion, adOpenKeyset, adLockOptimistic
    '
    If rstMultiCheq.RecordCount > 0 Then
        rstMultiCheq.MoveFirst
        If rstMultiCheq!IDCheque > Left(intChq, Len(intChq) - Len(DtcAPago(6))) Then
            strFormula = ""
        Else
            
            Do Until rstMultiCheq.EOF
                'If strNumCheq <> rstMultiCheq!IdCheque Then
                strMonto = Format(rstMultiCheq!Total, "#,##0.00")
                strNumCheq = Format(rstMultiCheq!IDCheque, "000000")
                
                If Len(strMonto) < 13 Then strMonto = String$(13 - Len(strMonto), " ") & strMonto
                'End If
                strFormula = strFormula & "CHEQUE N° " & strNumCheq & " Bs. " & strMonto & " "
                'End If
                rstMultiCheq.MoveNext
            Loop
        End If
    strFormula = Left(strFormula, 240)
    Else
    End If
    rstMultiCheq.Close
    Set rstMultiCheq = Nothing
    '
    cAletra.Numero = curTotal
    strAletra = String(10, "x") & cAletra.ALetra
    If Len(strAletra) <= 98 Then
        strAletra = strAletra & " " & String(10, "x") & " " & String(Len(strAletra), "x")
    End If
    'crtCheque.Reset
    '
    'crtCheque.Formulas(0) = "aletras='" & strAletra & "'"
    'crtCheque.Destination = IIf(Option1(4), crptToPrinter, crptToWindow)
    
    If intT = 1 Then
    
        'Selecciona el modelo de cheque a imprimir
        If DtcAPago(5) = "PROVINCIAL" Then
            strReporte = CHEQUE_PROVINCIAL
            
        ElseIf DtcAPago(5) = "BANESCO" Or DtcAPago(5) = "CARIBE" Then
            strReporte = CHEQUE_BANESCO
            
        ElseIf DtcAPago(5) = "VENEZUELA" Then
            Dim strMensaje As String
            strMensaje = "El cheque en impresión pertenece al nuevo formato de chequera" _
            & vbCrLf & "del Banco Venezuela?"
            '
            If Respuesta(strMensaje) Then
                strReporte = CHEQUE_BANESCO
            Else
                strReporte = CHEQUE_VENEZUELA
            End If
            '
        End If
        'Configura el control report
        Set rpReporte = New ctlReport
        rpReporte.Formulas(0) = "aletras='" & strAletra & "'"
        rpReporte.Salida = IIf(Option1(4), crImpresora, crPantalla)
        rpReporte.Reporte = gcReport & strReporte
        rpReporte.OrigenDatos(0) = gcPath & "\sac.mdb"
        MsgBox "Inserte el cheque en la bandeja 1 de la impresora: " & Printer.DeviceName, _
        vbInformation, "Impresión Cheque " & Left(intChq, 6)
        'crtCheque.CopiesToPrinter = 1
        'errlocal = crtCheque.PrintReport
        errLocal = rpReporte.Imprimir
        Call rtnBitacora("Impresión Cheque #" & Left(intChq, Len(intChq) - Len(DtcAPago(6))))
        If errLocal <> 0 Then
            MsgBox "Error al imprimir el cheque. " & Err.Description, vbCritical, _
            "Error " & Err
            Call rtnBitacora("Ocurrió el Error " & Err & " durante la impresión")
            errLocal = 0
            ftnPrintCheque = ""
            Set rpReporte = Nothing
            Exit Function
            
        End If
        '
    End If
    '
    'With crtCheque
    Set rpReporte = New ctlReport
    With rpReporte
        rpReporte.Formulas(0) = "aletras='" & strAletra & "'"
        rpReporte.Salida = IIf(Option1(4), crImpresora, crPantalla)
        .Formulas(1) = "Detalle='" & strFormula & "'"
        If intT <> 1 Then .Formulas(2) = "Reimpresion='" & Date & "'"
        .OrigenDatos(0) = mcDatos
        .OrigenDatos(1) = gcPath & "\sac.mdb"
        .Reporte = gcReport + "chq_voucher.rpt"
        .TituloVentana = "Voucher #" & Left(intChq, 6)
        errLocal = .Imprimir
        
        If errLocal <> 0 Then
            MsgBox "Error al imprimir el voucher. " & Err.Description, vbCritical, _
            "Error " & Err
            errLocal = 0
        End If
    End With
    Set rpReporte = Nothing
    
    Set rpReporte = New ctlReport
    With rpReporte
        .Reporte = gcReport + "chq_consecutivo.rpt"
        .OrigenDatos(0) = mcDatos
        .OrigenDatos(1) = gcPath & "\sac.mdb"
        rpReporte.Formulas(0) = "aletras='" & strAletra & "'"
        rpReporte.Salida = IIf(Option1(4), crImpresora, crPantalla)
        .TituloVentana = "Consecutivo #" & Left(intChq, 6)
        errLocal = .Imprimir
        If errLocal <> 0 Then
            MsgBox "Error al imprimir el consecutivo. " & Err.Description, vbCritical, _
            "Error " & Err
            errLocal = 0
        End If
     End With
     Set rpReporte = Nothing
    '
    MousePointer = vbDefault
    End Function

    Private Sub Form_Resize()
    'configura la presetacion del formulario
    Dim Ficha As Long
    
    On Error Resume Next
    
    With STabAPago
        Ficha = .tab
        .Width = Me.ScaleWidth - .Left
        .Height = Me.ScaleHeight - .Top
        'ficha consecutivos
        .tab = 3
        MshgAPago(6).Left = 100
        MshgAPago(6).Width = .Width - 200
        MshgAPago(6).Height = .Height * 0.45
        MshgAPago(7).Top = MshgAPago(6).Top + MshgAPago(6).Height
        MshgAPago(7).Left = MshgAPago(6).Left
        MshgAPago(7).Width = MshgAPago(6).Width
        MshgAPago(7).Height = .Height * 0.2
        
        '
        fraApago(10).Left = 100
        fraApago(10).Top = MshgAPago(7).Top + MshgAPago(7).Height + 100
        'ficha lista pagos
        .tab = 2
        MshgAPago(3).Left = 100
        MshgAPago(4).Left = 100
        MshgAPago(3).Width = .Width - 200
        MshgAPago(4).Width = .Width - 200
        MshgAPago(3).Height = MshgAPago(6).Height
        MshgAPago(4).Height = MshgAPago(7).Height
        MshgAPago(4).Top = MshgAPago(3).Top + MshgAPago(3).Height
        '
        fraApago(3).Top = MshgAPago(4).Top + MshgAPago(4).Height + 100
        fraApago(7).Top = MshgAPago(4).Top + MshgAPago(4).Height + 100
        fraApago(11).Top = MshgAPago(4).Top + MshgAPago(4).Height + 100
        fraApago(7).Left = 100
        fraApago(3).Left = fraApago(7).Width + 200
        fraApago(11).Left = fraApago(7).Width + fraApago(3).Width + 300
        '
        'ficha lista facturas
        .tab = 1
        DataGrid1.Width = .Width - (DataGrid1.Left * 2)
        DataGrid1.Height = .Height - DataGrid1.Top - fraApago(4).Height - 200
        DataGrid1.Columns(3).Width = DataGrid1.Width - 1000 - DataGrid1.Columns(0).Width - DataGrid1.Columns(1).Width - DataGrid1.Columns(2).Width - DataGrid1.Columns(4).Width - DataGrid1.Columns(5).Width
        fraApago(4).Top = DataGrid1.Top + DataGrid1.Height + 100
        fraApago(5).Top = DataGrid1.Top + DataGrid1.Height + 100
        
        
        .tab = Ficha
        
    End With
    '
    End Sub

    '---------------------------------------------------------------------------------------------
    Private Sub Form_Unload(Cancel As Integer)  '
    'Rutina Destructor----------------------------------------------------------------------------
    On Error Resume Next
    Set cAletra = Nothing
    Set rstList = Nothing
    Set rst = Nothing
    Set RstFacturasA = Nothing
    RstBco.Close
    Set RstBco = Nothing
    Set RstCheque = Nothing
    Set RstGastos = Nothing
    GtoCnn.Close
    Set GtoCnn = Nothing
    '
    End Sub

Private Sub fraApago_MouseMove(Index As Integer, Button As Integer, Shift As Integer, _
X As Single, Y As Single)
If (Picture1.Tag = 0 And Picture1.LinkItem = 102) Or (Picture1.Tag = 1 And Picture1.LinkItem = 104) Then
    Picture1.Picture = LoadResCustom(IIf(Picture1.Tag = 0, 101, 103), "CUSTOM")
    Picture1.LinkItem = IIf(Picture1.Tag = 0, 101, 103)
    Picture1.Height = 2970
End If

End Sub

Private Sub Image1_Click()
If Image1.Tag = 0 Then
    Image1.Left = 10830
    Image1.Tag = 1
    tabDetalle.Visible = False
    Image1.Picture = LoadResCustom(103, "custom")
Else
    Image1.Left = 6645
    Image1.Tag = 0
    tabDetalle.Visible = True
    Image1.Picture = LoadResCustom(101, "custom")
End If
Image1.Height = 2970
End Sub

Private Sub Image1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'
If (Y >= 1200 And Y <= 1750) Then
    Image1.Picture = LoadResCustom(IIf(Image1.Tag = 0, 102, 104), "CUSTOM")
    Image1.Height = 2970
End If
'
End Sub


    '---------------------------------------------------------------------------------------------
    Private Sub LblAPago_Click(Index As Integer)    '
    '---------------------------------------------------------------------------------------------
    '
    If Index = 13 Then
    '
        If LblAPago(13).ForeColor = &H800000 Then
            With LblAPago(13)
              .Caption = "Opciones Avanzadas «««"
              .BorderStyle = 1
              .ForeColor = &HFFFFFF
              .BackColor = &H800000
            End With
            
            For I = 14 To 18
                  LblAPago(I).Visible = True
            Next
            For I = 0 To 1
                chkPago(I).Visible = True
                dtpApago(I + 2).Visible = True
                txtAPago(I + 9).Visible = True
                txtAPago(I + 7).Width = 1980
            Next
            scrApago.Visible = True
            CmdRegistrar(1).Width = 870
        Else
            fraApago(2).Top = 60
            With LblAPago(13)
                .Caption = "Opciones Avanzadas »»»"
                .BorderStyle = 0
                .ForeColor = &H800000
                .BackColor = &H80000004
            End With
            For I = 14 To 18
                LblAPago(I).Visible = False
            Next
            For I = 0 To 1
                chkPago(I).Visible = False
                chkPago(I).Value = 0
                dtpApago(I + 2).Visible = False
                txtAPago(I + 9).Visible = False
                txtAPago(I + 7).Width = 2325
            Next
            scrApago.Visible = False
            CmdRegistrar(1).Width = 1230
        
        End If
    
    End If
    '
    End Sub

    '---------------------------------------------------------------------------------------------
    Private Sub MshgAPago_Click(Index As Integer)   '
    '---------------------------------------------------------------------------------------------
    '
    Select Case Index
    '
        Case 0
    '   ---------------------
            With MshgAPago(0)   'Flex Cargado'
    '
                If .TextMatrix(.RowSel, 4) = "" Then Exit Sub
                If CCur(.TextMatrix(.RowSel, 4)) = 0 Then
                    MsgBox "Cod.de Gasto Asignado Totalmente", vbInformation, "Cod.Gasto: " _
                    & .TextMatrix(.RowSel, 0)
                    Exit Sub
                End If
    '
            End With
    '
            
            Dim intFila, IntColumna As Integer
            intFila = MshgAPago(0).RowSel
            If IsNull(VecVinculo(intFila, 0)) Or VecVinculo(intFila, 0) = 0 Then
                VecVinculo(intFila, 0) = intFila 'LLENA LA PRIMERA COLUMNA DEL VECTOR
            Else
                Exit Sub
            End If
            With MshgAPago(1)
                .Row = 1
                Do Until .TextMatrix(.Row, 0) = ""
                    If .Row < .Rows Then .Row = .Row + 1
                Loop
                IntColumna = .RowSel
                Dim I As Integer
                For I = 0 To 2
                    .TextMatrix(IntColumna, I) = MshgAPago(0).TextMatrix(intFila, I)
                Next
                .TextMatrix(IntColumna, 3) = MshgAPago(0).TextMatrix(intFila, 4)
                .TextMatrix(IntColumna, 4) = MshgAPago(0).TextMatrix(intFila, 5)
                .TextMatrix(IntColumna, 5) = intFila
                MshgAPago(0).TextMatrix(MshgAPago(0).Row, 4) = Format(0, "#,##0.00")
                .TextMatrix(IntColumna, 6) = MshgAPago(0).TextMatrix(intFila, 6)
            End With
            txtAPago(5) = Format(ftnMonto, "#,##0.00")
            LblAPago(32) = Format(ftnDif, "#,##0.00")
        
        Case 3  'Cheques emitidos
    '   ---------------------
        If MshgAPago(3).RowSel = 0 Then Exit Sub
        If optApago(2).Value = False Then
            If MshgAPago(3).TextMatrix(MshgAPago(3).RowSel, 0) <> "" Then
                If Option1(3).Value = True Then
                    If MshgAPago(3).CellPicture = imgPago Then
                        Set MshgAPago(3).CellPicture = Nothing
                    Else
                        Set MshgAPago(3).CellPicture = imgPago
                    End If
                    
                End If
            With Adodc1(1)
                .RecordSource = "SELECT CodInm & ' ' &  CodGasto as Cuenta, *  FROM ChequeDetal" _
                & "le WHERE Clave='" & CLng(MshgAPago(3).TextMatrix(MshgAPago(3).RowSel, 0)) _
                & MshgAPago(3).TextMatrix(MshgAPago(3).RowSel, 5) & "'"
                .Refresh
                If .Recordset.RecordCount > 0 Then
                    'variables locales
                    Dim curCheque As Currency
                    Dim j As Integer
                    '
                    MshgAPago(4).Rows = .Recordset.RecordCount + 1: j = 0
                    MshgAPago(4).ColAlignment(2) = flexAlignLeftCenter
'                    If MshgAPago(4).Rows > 5 Then
'                        MshgAPago(4).ColWidth(3) = 1400
'                        MshgAPago(4).ColWidth(4) = 1400
'                    Else
'                        MshgAPago(4).ColWidth(3) = 1500
'                        MshgAPago(4).ColWidth(4) = 1500
'                    End If
                    Do Until .Recordset.EOF
                        j = j + 1
                        MshgAPago(4).TextMatrix(j, 0) = .Recordset!Cuenta
                        MshgAPago(4).TextMatrix(j, 1) = Format(.Recordset!cargado, "mm-yy")
                        MshgAPago(4).TextMatrix(j, 2) = .Recordset!detalle
                        MshgAPago(4).TextMatrix(j, 3) = Format(.Recordset!Monto, "#,##0.00")
                        MshgAPago(4).TextMatrix(j, 4) = Format(0, "#,##0.00")
                        curCheque = curCheque + .Recordset!Monto
                        .Recordset.MoveNext
                    Loop
                    MshgAPago(4).AddItem ("")
                    MshgAPago(4).TextMatrix(j + 1, 2) = "BANCO"
                    MshgAPago(4).TextMatrix(j + 1, 3) = Format(0, "#,##0.00")
                    MshgAPago(4).TextMatrix(j + 1, 4) = Format(curCheque, "#,##0.00")
                Else
                MshgAPago(4).Rows = 1
                End If
            End With
            '
        End If
        '
    End If
    Case 6  'consecutivos
    If MshgAPago(6).RowSel = 0 Or MshgAPago(6).TextMatrix(MshgAPago(6).RowSel, 0) = "" Then _
    Exit Sub
    'Dim NumCheq As Long
    'NumCheq = MshgAPago(6).TextMatrix(MshgAPago(6).RowSel, 0)
    Call busca_detalle(MshgAPago(6).TextMatrix(MshgAPago(6).RowSel, 0))
    '
    End Select
    '
    End Sub

    '---------------------------------------------------------------------------------------------
    Private Sub MshgAPago_EnterCell(Index As Integer)   '
    '---------------------------------------------------------------------------------------------
    '
    'Dim NumCheq As Long 'variables locales
    '
    If Index = 1 And MshgAPago(1).Col = 3 Then
        With MshgAPago(1)
            If .Text <> "" Then
                .Text = CCur(.Text)
            End If
        End With
    'ElseIf Index = 6 Then
    End If
    '
    End Sub

    '---------------------------------------------------------------------------------------------
    Private Sub MshgAPago_KeyPress(Index As Integer, KeyAscii As Integer)   '
    '---------------------------------------------------------------------------------------------
    '
    Select Case Index
        '
        Case 1
    '   ---------------------
            With MshgAPago(1)
                'SOLO PERMITE EDITAR LA COLUMNA DE MONTOS
                If .Row > 0 And .Col = 3 And .TextMatrix(.Row, 1) <> "" Then
                    
                    If KeyAscii = 46 Then KeyAscii = 44 'CONVIERTE PUNTO(.) EN COMA(,)
                    If KeyAscii >= 48 And KeyAscii <= 57 Or KeyAscii = 44 Then
                        'SOLO ACEPTA DATOS NUMERICOS
                        .TextMatrix(.RowSel, 3) = .TextMatrix(.RowSel, 3) & Chr(KeyAscii)
                    End If
                    
                    If KeyAscii = 13 Then 'EL USUARIO PRESIONO ENTER
                        
                        If .Text <> "" Then 'SI LA CELDA CONTIENE ALGUNA CANTIDAD
                            .Text = Format(.Text, "#,##0.00")
                            Call rtnDif(CCur(.Text), .TextMatrix(.RowSel, 5))
                        Else    'SI EL USUARIO DEJA LA CELDA VACIA
                        
                            If .TextMatrix(.Row, 1) <> "" Then .Text = Format(0, "#,##0.00")
                            MshgAPago(0).TextMatrix(.Row, .Col + 1) = _
                            Format(CurDiferencia(.Row) - .Text, "#,##0.00")
                        
                        End If
                        'PASA EL FOCO A LA SIGUIETE CELDA
                        If .Row = .Rows - 1 Then
                            .Row = 1
                        Else
                            'SI ES LA ULTIMA CELDA VUELVE A LA PRIMERA
                            .Row = .Row + 1
                            If .Text = "" Then .Row = 1
                        End If
                        '
                    End If
                    '
                End If
                '
            End With
        'cheques emitidos y consecutivos
        Case 6, 3: MshgAPago_Click (Index)
        
        
    End Select
    '
    End Sub

    '---------------------------------------------------------------------------------------------
    Private Sub MshgAPago_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
    '---------------------------------------------------------------------------------------------
    '
    Select Case Index
    '
        Case 1
            With MshgAPago(1)
                If .Row > 0 And .Col = 3 Then
                    Select Case KeyCode
                    
                        Case 46 'SUPRIMIR 'BORRA TODO EL CONTENIDO'
                            .TextMatrix(.RowSel, 3) = ""
                        
                        Case 8  'BACKSPACE 'BORRA EL ULTIMO CARACETER'
                            If Len(.Text) > 0 Then
                                .Text = Left(.Text, Len(.Text) - 1)
                            End If
                            '
                    End Select
                    '
                End If
                '
            End With
            '
    End Select
    '
    End Sub

    '---------------------------------------------------------------------------------------------
    Private Sub MshgAPago_LeaveCell(Index As Integer)   '
    '---------------------------------------------------------------------------------------------
    'AL PERDER EL FOCO UNA CELDA DA FORMATO '#,##0.00' A LA MISMA*
    
    Select Case Index
    '
        Case 1
        '
            With MshgAPago(1)
                If .Text = "" Then Exit Sub
            '
                If .Col = 3 And .Row <> 0 Then
                '
                    If .Text <> "" Then
                    '
                        .Text = Format(.Text, "#,##0.00")
                        '
                    Else
                    '
                        If .TextMatrix(.Row, 1) <> "" Then .Text = Format(0, "#,##0.00")
                        MshgAPago(0).TextMatrix(.Row, .Col + 1) = MshgAPago(0).TextMatrix(.Row, .Col)
                        '
                    End If
                    '
                    Call rtnDif(CCur(.Text), .Row)
                    txtAPago(5) = Format(ftnMonto, "#,##0.00")
                    LblAPago(32) = Format(ftnDif, "#,##0.00")
                End If
                '
            End With
            '
    End Select
    '
    End Sub

    Private Sub MshgAPago_MouseDown(Index As Integer, Button As Integer, Shift As Integer, _
    X As Single, Y As Single)
    '
    If Index = 3 And Button = 2 Then    'Si presiona el boton secundario sobre el grid
        With MshgAPago(3)                   'de cheques impresos
            .Col = 0
            .Row = .RowSel
            If .CellPicture = imgPago Then PopupMenu AC204, , _
            .Left + imgPago.Width + STabAPago.Left
        End With
    End If
    '
    End Sub


    Private Sub MskFecha_KeyDown(KeyCode As Integer, Shift As Integer)
    If gcNivel > nuAdministrador Then KeyCode = 0
    End Sub

    Private Sub MskFecha_KeyPress(KeyAscii As Integer)
    If gcNivel > nuAdministrador Then
        KeyAscii = 0
    Else
        Call Validacion(KeyAscii, "0123456789")
    End If
    End Sub

    '---------------------------------------------------------------------------------------------
    Private Sub optApago_Click(Index As Integer)    '
    '---------------------------------------------------------------------------------------------
    MousePointer = vbHourglass
    Select Case Index
        Case 0  'Impreso
    '   ---------------------
            Toolbar1.Buttons(9).Enabled = True
            Toolbar1.Buttons("Edit").Enabled = True
            fraApago(7).Enabled = True
            
        Case 1  'Pendiente
    '   ---------------------
            Toolbar1.Buttons(9).Enabled = False
            Toolbar1.Buttons("Edit").Enabled = False
            fraApago(7).Enabled = True
            
        Case 2  'Anulado
    '   ---------------------
            Toolbar1.Buttons(9).Enabled = False
            Toolbar1.Buttons("Edit").Enabled = False
            fraApago(7).Enabled = False
            
    End Select
    Call rtnCheqE(Index)
    MousePointer = vbDefault
    '
    End Sub

    '---------------------------------------------------------------------------------------------
    Private Sub optFiltro_Click(Index As Integer)   'aplica filtro a la cuadrícula
    '---------------------------------------------------------------------------------------------
    '
    Select Case Index
        Case 0  'No se aplica filtro (predeterminado)
    '   -------------------------------
'            RstFacturasA.Filter = ""
'            RstFacturasA.Sort = "codProv, Frecep"
            strSQL = ftnSQL(" ORDER BY Cpp.codProv, cpp.FRecep")
            
        Case 1  'se filtra por un proveedor en especial
    '   -------------------------------
            If DtcAPago(4) = "" Then
                DtcAPago(4).SetFocus
                Exit Sub
            End If
            strSQL = ftnSQL(" AND CodProv='" & DtcAPago(4) & "' ORDER BY Cpp.FRecep,Cpp.Ndoc")
            
    End Select
    RstFacturasA.Close
    RstFacturasA.Open strSQL, cnnConexion, adOpenKeyset, adLockOptimistic, adCmdText
    Set DataGrid1.DataSource = RstFacturasA
    '
    End Sub

    '---------------------------------------------------------------------------------------------
    Private Sub Option1_Click(Index As Integer) '
    '---------------------------------------------------------------------------------------------
    '
    Select Case Index
    '
        Case 0  'CUENTA DEL INMUEBLE
    '   ---------------------
            Call rtnBusca_Cuenta(mcDatos)
        
        Case 1  'CUENTA DE LA ADMINISTRADORA
    '   ---------------------
            Call rtnBusca_Cuenta(gcPath & "\" + sysCodInm + "\inm.mdb")
    '
        Case 2  'Impresión Todos los Cheques pendientes
    '   ---------------------
    
        Case 3  'Impresión cheque seleccionado
    '   ---------------------
        
    '
    End Select
    '
    End Sub
    
Private Sub Picture1_Click()
If Picture1.Tag = 0 Then
    Picture1.Left = 10830
    Picture1.Tag = 1
    tabDetalle.Visible = False
    Picture1.Picture = LoadResCustom(103, "custom")
    Picture1.LinkItem = 103
    MshgAPago(2).Width = 10600
    MshgAPago(2).ColWidth(2) = 3000 + (10600 - 6315)
Else
    Picture1.Left = 6645
    Picture1.Tag = 0
    tabDetalle.Visible = True
    Picture1.Picture = LoadResCustom(101, "custom")
    Picture1.LinkItem = 101
    MshgAPago(2).Width = 6315
    MshgAPago(2).ColWidth(2) = 3000
End If
Picture1.Height = 2970

End Sub

Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If (Y >= 1200 And Y <= 1750) And ((Picture1.Tag = 0 And Picture1.LinkItem = 101) Or (Picture1.Tag = 1 And Picture1.LinkItem = 103)) Then
    Picture1.Picture = LoadResCustom(IIf(Picture1.Tag = 0, 102, 104), "CUSTOM")
    Picture1.LinkItem = IIf(Picture1.Tag = 0, 102, 104)
    Picture1.Height = 2970
End If

End Sub

    '---------------------------------------------------------------------------------------------
    Private Sub scrApago_Change()   '
    '---------------------------------------------------------------------------------------------
    '
    fraApago(2).Top = 60 - (scrApago.Value * 200)
    '
    End Sub

    '---------------------------------------------------------------------------------------------
    Private Sub STabAPago_Click(PreviousTab As Integer) '
    '---------------------------------------------------------------------------------------------
    '
    Select Case STabAPago.tab   'SELECCIONA UNA FICHA
    
        Case 0  'DATOS GENERALES
    '   ---------------------
            Toolbar1.Buttons(9).Enabled = False
            Toolbar1.Buttons("Print").Enabled = False
            'aqui llama rutina de verificación
            'RETARDO
            For I = 0 To 100000
            Next
            'VARIFICA QUE LO FACTURA = ABONADO + SALDO
            Dim CurFac@, curAbo@, curSal@
            'facturado
            With MshgAPago(2)
                For I = 1 To .Rows - 1
                    If .TextMatrix(I, 3) <> "" Then CurFac = CurFac + CCur(.TextMatrix(I, 3))
                Next
            End With
            'saldo
            With MshgAPago(0)
                For I = 1 To .Rows - 1
                    If .TextMatrix(I, 4) <> "" Then curSal = curSal + CCur(.TextMatrix(I, 4))
                Next
            End With
            'abonado
            tabDetalle.tab = 1
            tabDetalle.tab = 0
            With MshgAPago(5)
                For I = 1 To .Rows - 1
                    If .TextMatrix(I, 2) <> "" Then curAbo = curAbo + CCur(.TextMatrix(I, 2))
                Next
            End With
            If CurFac - curAbo <> curSal Then
                MsgBox "Existe direncia en los pagos de este documento. Revise los abonos.", vbExclamation, App.ProductName
            End If
            
        Case 1  'LISTA DE FACTURAS ASIGNADAS
    '   ---------------------
            Toolbar1.Buttons(9).Enabled = False
            Toolbar1.Buttons("Print").Enabled = False
            RstFacturasA.Requery
            
        Case 2  'LISTA DE PAGOS ASIGNADOS
    '   ---------------------
            MousePointer = vbHourglass
            Toolbar1.Buttons("Print").Enabled = True
            If optApago(0) Then Toolbar1.Buttons(9).Enabled = True
            If MshgAPago(3).Rows >= 1 And optApago(0) Then
                Toolbar1.Buttons("Edit").Enabled = True
            End If
            Call rtnSelOPT
            MousePointer = vbDefault
       
       Case 3 'Consecutivos
    '   ---------------------
            Toolbar1.Buttons(9).Enabled = False
            Toolbar1.Buttons("Print").Enabled = True
            Call centra_titulo(MshgAPago(6), True)
            Call centra_titulo(MshgAPago(7), True)
            dtpApago(5).Value = Date
            dtpApago(6).Value = Date
            '
    End Select
    '
    End Sub

    Private Sub tabDetalle_Click(PreviousTab As Integer)
    If tabDetalle.tab = 1 Then
        Dim strSQL As String, strRango As String
        With MshgAPago(2)
            I = 1
            Do Until .TextMatrix(I, 0) = ""
                If I = 1 Then
                    strRango = "('" & .TextMatrix(I, 0)
                Else
                    strRango = strRango & "','" & .TextMatrix(I, 0)
                End If
                I = I + 1
            Loop
            strRango = strRango & "')"
        End With
        If Len(strRango) < 5 Then Exit Sub
        strSQL = "SELECT ChequeDetalle.IDCheque, Format(cheque.Fecha,'dd/mm/yy'), Format(Sum(Ch" _
        & "equeDetalle.Monto),'#,##0.00') AS Total FROM cheque INNER JOIN ChequeDetalle ON cheq" _
        & "ue.IDCheque = ChequeDetalle.IDCheque WHERE (((ChequeDetalle.IDCheque) In (SELECT IDC" _
        & "heque From ChequeFactura WHERE Ndoc IN " & strRango & "))) GROUP BY ChequeDetalle.ID" _
        & "Cheque, cheque.Fecha;"
        '
        Dim rstAbono As New ADODB.Recordset
        rstAbono.Open strSQL, cnnConexion, adOpenStatic, adLockReadOnly
        
        With MshgAPago(5)
            Set .DataSource = rstAbono
            .FormatString = "^Cheque |^Fecha |>Monto"
            Call centra_titulo(MshgAPago(5), True)
            Dim curTotalAbono As Currency
            For I = 1 To .Rows - 1
                curTotalAbono = curTotalAbono + .TextMatrix(I, 2)
            Next
            If rstAbono.RecordCount > 0 Then .Row = 1
            txtAPago(11) = Format(curTotalAbono, "#,##0.00")
        End With
        rstAbono.Close
        Set rstAbono = Nothing
        
    End If
    End Sub

    '---------------------------------------------------------------------------------------------
    Private Sub Toolbar1_ButtonClick(ByVal Button As ComctlLib.Button)  '
    '---------------------------------------------------------------------------------------------
    'variables locales
    Dim errLocal As Long
    Dim intCheque As Long
    Dim intConfirma As Integer
    Dim I As Integer
    Dim strClave As String, strSQL As String
    Dim rpReporte As ctlReport
    '
    Select Case UCase(Button.Key)
    '
'        Case "FIRST"    'Primer Registro
'    '   ---------------------
'        Case "PREVIOUS" 'Registro Previo
'    '   ---------------------
'        Case "NEXT" 'Siguiente Registro
'    '   ---------------------
'        Case "END"  'Último Registro
'    '   ---------------------
       Case "NEW"  'Nuevo Registro
    '   ---------------------
            Call rtnNewCan("True")
            Call RtnEstado(5, Toolbar1)

        Case "SAVE" 'Actualizar Registro
    '   ---------------------
            
            If ftnValida = True Then Exit Sub
            On Error GoTo NotSave
            
            Call rtnRegistar
            
            If CCur(LblAPago(32)) = 0 Then
                
                With MshgAPago(2)
                    Dim strDocs As String
                    I = 1
                    While I < .Rows And .TextMatrix(I, 0) <> ""
                        strDocs = IIf(strDocs = "", "'" & .TextMatrix(I, 0) _
                        & "'", strDocs & ",'" & .TextMatrix(I, 0) & "'")
                        I = I + 1
                    Wend
                    'actualiza el estatus a 'PAGADO' en la tabla de cuentas por pagar
                    cnnConexion.Execute "UPDATE Cpp SET Estatus='PAGADO', Usuario='" & gcUsuario _
                    & "',Freg=Date() WHERE Ndoc IN (" & strDocs & ");"
                    RstFacturasA.Requery
                    DataGrid1.ReBind
                End With
                '
                If Err.Number = 0 Then
                    cnnConexion.CommitTrans
                    MsgBox "Transacción Procesada..", vbInformation, App.ProductName
                Else
                    cnnConexion.RollbackTrans
                    Call RtnEstado(5, Toolbar1)
                    MsgBox "Operación Cancelada.." & Err.Description, vbExclamation, App.ProductName
                End If
                Call Fin_Emitir
            Else
                If Respuesta("¿Desea emitir otro cheque?") Then
                    blnOtro = True
                    Call rtnLimpiar_Grid(MshgAPago(1))
                    For I = 1 To MshgAPago(0).Rows - 1
                        CurDiferencia(I) = MshgAPago(0).TextMatrix(I, 4)
                    Next
                    ReDim VecVinculo(MshgAPago(0).Rows, 2)
                    MshgAPago(1).Rows = MshgAPago(0).Rows
                Else
NotSave:
                    If Err.Number = 0 Then
                        cnnConexion.CommitTrans
                        MsgBox "Transacción Procesada", vbInformation, App.ProductName
                    Else
                        cnnConexion.RollbackTrans
                        Call RtnEstado(5, Toolbar1)
                        MsgBox "Operación Cancelada.." & Err.Description, vbInformation, App.ProductName
                    End If
                    Call Fin_Emitir
                    '
                End If
            End If
            txtAPago(5) = "0,00"
            '
        Case "FIND" 'Buscar
    '   ---------------------
            STabAPago.tab = 2
            
        Case "UNDO" 'Deshacer
    '   ---------------------
            Call rtnNewCan("FALSE")
            Call Fin_Emitir
            Call RtnEstado(6, Toolbar1)
            Toolbar1.Buttons(9).Enabled = False
            
        Case "DELETE"   'Eliminar registro
    '   ---------------------
    
            On Error Resume Next
            cnnConexion.BeginTrans
            strClave = CLng(MshgAPago(3).TextMatrix(MshgAPago(3).RowSel, 0)) & _
            MshgAPago(3).TextMatrix(MshgAPago(3).RowSel, 5)
            '
            With Adodc1(0).Recordset
            '
                .MoveFirst
                
                .Find "ID='" & MshgAPago(3).TextMatrix(MshgAPago(3).RowSel, 0) & "'"
                '
                'elimina cualquier el registro de la tabla movfondo
                'cnnConexion.Execute "UPDATE MovFondo IN '" & gcpathSET Del = True WHERE Concepto LIKE '*" & _
                MshgAPago(3).TextMatrix(MshgAPago(3).RowSel, 0) & "'"
                '
                'Coloca el cheque en la tabla de cheques anulados
                '
                cnnConexion.Execute "INSERT INTO ChequeAnulado(IDCheque,FechaCheque,Monto,Benef" _
                & "iciario,Banco,Cuenta,CodInm,Concepto,Usuario,Fecha,Hora,Clave) SELECT TOP 1 " _
                & !ID & ",'" & !fecha & "','" & !Monto & "','" & !Beneficiario & "','" & !Banco _
                & "','" & !Cuenta & "','" & gcCodInm & "','" & Left(!Concepto, 200) & "','" & _
                gcUsuario & "','" & Date & "','" & Time() & "','" & !ID & !Cuenta & "' FROM Che" _
                & "queDetalle WHERE Clave='" & strClave & "' GROUP BY IDCheque,CodInm;"
                '
                'Actualiza el estado de la factura a 'ASIGNADO'
                Dim rstAbonos As New ADODB.Recordset
                '
                rstAbonos.Open "SELECT * FROM ChequeFactura WHERE Clave = '" & strClave & "' AN" _
                & "D Ndoc IN (SELECT Ndoc FROM ChequeFactura WHERE Clave='" & strClave & "');", _
                cnnConexion, adOpenStatic, adLockReadOnly, adCmdText
                '
                If rstAbonos.RecordCount > 0 Then
                
                    cnnConexion.Execute "UPDATE Cpp SET  Estatus='ASIGNADO' WHERE Ndoc IN (SELE" _
                    & "CT Ndoc FROM ChequeFactura WHERE Clave='" & strClave & "');"
                    
                End If
                '
                rstAbonos.Close
                '
                'Reintegra el movimiento del fondo si el cheque tiene que ver
                rstAbonos.Open "SELECT * FROM ChequeDetalle WHERE Clave = '" & strClave & "'", _
                cnnConexion, adOpenKeyset, adLockOptimistic, adCmdText
                
                If Not rstAbonos.EOF And Not rstAbonos.BOF Then
                
                    cnnConexion.Execute "INSERT INTO MovFondo (CodGasto,Fecha,Tipo,Periodo,Conc" _
                    & "epto,Debe,Haber) IN '" & gcPath & "\" & rstAbonos("CodInm") & "\inm.mdb'" _
                    & " SELECT CodGasto, Date(), 'NC', (SELECT DateAdd('M',1,MAX(Periodo)) FROM" _
                    & " Factura IN '" & gcPath & "\" & rstAbonos("CodInm") & "\inm.mdb' WHERE Fa" _
                    & "ct Not Like 'CH%'), '(ANULADO) ' & right(Concepto, 40) ,0,  debe FROM MovFondo IN '" _
                    & gcPath & "\" & rstAbonos("CodInm") & "\inm.mdb' WHERE Concepto LIKE '%" _
                    & rstAbonos("IDCheque") & "' AND Tipo='CH'", N

                    
                End If
                rstAbonos.Close
                '
                Set rstAbonos = Nothing
                '
                'Elimina el registro de los cheques emitidos
                cnnConexion.Execute "DELETE * FROM Cheque WHERE Clave='" & strClave & "'"
                
            End With
            '
            If Respuesta("Está seguro de Anular el Cheque #" & _
            MshgAPago(3).TextMatrix(MshgAPago(3).RowSel, 0)) And Err.Number = 0 Then
            '
                cnnConexion.CommitTrans
                Call rtnBitacora("Cheque " & MshgAPago(3).TextMatrix(MshgAPago(3).RowSel, 0) _
                & " eliminado...")
                Call rtnLimpiar_Grid(MshgAPago(3))
                Call rtnLimpiar_Grid(MshgAPago(4))
                
                MsgBox "Cheque Anulado", vbInformation, App.ProductName
                '
            Else
            
                cnnConexion.RollbackTrans
                If Err.Number <> 0 Then MsgBox Err.Description, vbCritical, App.ProductName
                
            End If
    
        Case "EDIT" 'Editar Registro
    '   ---------------------
    
            With MshgAPago(3)
                .Col = 0
                .Row = 1
                If Not .Text = "" Then
                    blnEdit = True
                Else
                    blnEdit = False
                End If
            End With
            
        Case "PRINT"    'Imprimir Reporte
    '   ---------------------
            If STabAPago.tab = 3 Then
                'Imprime el reporte de consecutivos
                If strCtvo <> "" Then
                    Call rtnGenerator(gcPath & "\sac.mdb", strCtvo, Consulta_Consecutivo)
                    'Call clear_Crystal(crtCheque)
                    Set rpReporte = New ctlReport
                    With rpReporte
                        '
                        .Reporte = gcReport + "consecutivos.rpt"
                        .OrigenDatos(0) = gcPath & "\sac.mdb"
                        .Formulas(0) = "SubTitulo='BANCO " & DtcAPago(8) & " - Cuenta Nº: " _
                        & DtcAPago(9) & "'"
                        .Formulas(1) = "Filtro='" & Filtrar & "'"
                        If Respuesta(LoadResString(537)) Then
                            .Salida = crImpresora
                        Else
                            .Salida = crPantalla
                            .TituloVentana = "Consecutivos"
                        End If
                        .Imprimir
                        Call rtnBitacora("Printer Consecutivos.." & DtcAPago(8) & "/" & DtcAPago(9))
                    End With
                Else
                    MsgBox LoadResString(535), vbInformation, LoadResString(536)
                End If
            ElseIf STabAPago.tab = 2 Then   'impresión de cheques
                Call RtnImprimir_Cheque(IIf(optApago(0), 1, 0))
            End If
            '
        Case "CLOSE"    'Cerrar Formulario
    '   ---------------------
            Unload Me
    '
    End Select
    
    '
    End Sub

    '---------------------------------------------------------------------------------------------
    Private Sub RtnBuscaCta(StrCampo$, StrValor$, Indice%, Indice1%) '
    '---------------------------------------------------------------------------------------------
    '
    Dim rstlocal As ADODB.Recordset
    Dim strSQL As String
    
    With RstBco
        .MoveFirst
        .Find StrCampo & "='" & StrValor & "'"
        If Not .EOF Then    'VALOR ENCONTRADO
            'verifica las chequeras en existencia
            Set rstlocal = New ADODB.Recordset
            strSQL = "SELECT Cuentas.IDCuenta, Count(Chequera.IDchequera) " _
            & "AS Registrada FROM (Bancos INNER JOIN Cuentas ON Bancos.IDBanco " _
            & "= Cuentas.IDBanco) INNER JOIN Chequera ON Cuentas.IDCuenta = " _
            & "Chequera.IDCuenta Where Cuentas.Inactiva = False And " _
            & "Chequera.Ultimo = 0 AND Cuentas.IDCuenta=" & !IDCuenta _
            & " GROUP BY Cuentas.IDCuenta, Cuentas.IDCuenta " _
            & "ORDER BY Cuentas.IDCuenta;"
            '
            With rstlocal
                .Open strSQL, RstBco.ActiveConnection, adOpenKeyset, adLockOptimistic, _
                adCmdText
                If Not (.EOF And .BOF) Then
                    If !registrada < 2 Then
                        MsgBox "Quedan " & !registrada & " chequera(s) registrada(s)." & _
                        vbCrLf & "Recurde solicitar al banco", vbInformation, App.ProductName
                    End If
                Else
                    MsgBox "No existen chequeras en reserva. Recuerde solicitar al banco", _
                    vbCritical, App.ProductName
                End If
            .Close
            End With
            Set rstlocal = Nothing
            GoSub Coincidencia
            '
        Else    'VALOR NO ENCONTRADO
            MsgBox "No Registrado '" & StrValor & "'", vbInformation, App.ProductName
        End If
    
    Exit Sub
Coincidencia:
    If StrCampo = "NombreBanco" Then
       DtcAPago(9) = .Fields("NumCuenta")
       DtcAPago(2) = .Fields("NumCuenta")
    Else
       DtcAPago(8) = .Fields("NombreBanco")
       DtcAPago(0) = .Fields("NombreBanco")
    End If
    Call DtcAPago_Change(9)
    Return
    End With
    
    '
    End Sub

    '---------------------------------------------------------------------------------------------
    '   Rutina:  rtnChqrs
    '
    '   Procedimiento que busca las chequera activas para determinado inmuelbe
    '   Llama la rutina que busca el cheque consecutivo
    '---------------------------------------------------------------------------------------------
    Private Sub RtnChqrs()
    '
    Dim RstChqrs As ADODB.Recordset
    Set RstChqrs = New ADODB.Recordset
    '
    strSQL = "SELECT Chequera.IDchequera, Cuentas.NumCuenta, Bancos.NombreBanco, Cheque" _
    & "ra.DESDE, Chequera.Hasta, AsignaChequera.Activa FROM Bancos INNER JOIN (Cuentas " _
    & "INNER JOIN (Chequera LEFT JOIN AsignaChequera ON Chequera.IDchequera = AsignaChe" _
    & "quera.IDChequera) ON Cuentas.IDCuenta = Chequera.IDCuenta) ON Bancos.IDBanco = Cuentas.I" _
    & "DBanco WHERE Asignachequera.Activa=True AND Bancos.NombreBanco='" & DtcAPago(1) & _
    "' AND Cuentas.NumCuenta='" & DtcAPago(2) & "';"
    RstChqrs.Open strSQL, RstBco.ActiveConnection, adOpenKeyset, adLockOptimistic
    '
    Set DtcAPago(3).RowSource = RstChqrs
    
    If RstChqrs.EOF And RstChqrs.BOF Then    'No hay chequeras registradas
        MsgBox "No existen Chequeras Disponibles del Banco " & DtcAPago(1) & vbCrLf _
        & "Cuenta N° " & DtcAPago(2), vbInformation, App.ProductName
        DtcAPago(3) = ""
        LblAPago(34) = ""
    Else
        DtcAPago(3) = RstChqrs.Fields(0)
        Call RtnConsecutivo 'Busca el cheque consecutivo
    End If
    '
    End Sub

    '---------------------------------------------------------------------------------------------
    '   Rutina:     rtnConsecutivo
    '
    '   Procedimiento que localiza el último cheque utilizado y asigna al pago
    '   el cheque consecutivo, o el primer cheque si la cheque según el caso
    '---------------------------------------------------------------------------------------------
    Private Sub RtnConsecutivo()
    '
    Set rst = New ADODB.Recordset
    rst.Open "SELECT MAX(Ultimo) AS Cheque,Desde  FROM Chequera WHERE IDChequera =" & _
    DtcAPago(3) & " GROUP BY Desde;", RstBco.ActiveConnection, adOpenKeyset, adLockOptimistic, _
    adCmdText
    If rst!Cheque = 0 Then
        Cheque = rst!Desde
    Else
        Cheque = rst!Cheque + 1
    End If
    LblAPago(34) = Format(Cheque, "000000")
    rst.Close
    Set rst = Nothing
    '
    End Sub


    '27/08/2002-----------------------------------------------------------------------------------
    '       Function:   ftnSQL
    '
    '       entradas:   una que contiene una condición para la clausula WHERE
    '                   de la instrucción SQL
    '
    '       Salida:     Una cadena Intrucción SQL para originar un grupo
    '                   de registros
    '       Función que genera la instrucción SQL que determina el grupo de
    '       origen de los registros a mostrar en el Grid de Facturas
    '---------------------------------------------------------------------------------------------
    Private Function ftnSQL(strCadena$) As String
    '
    If gcCodInm = sysCodInm Then    'Facturas asignadas de la cuenta pote
        ftnSQL = "SELECT Cpp.*, Caja.CodigoCaja, Proveedores.NombProv FROM Proveedores INNER JO" _
        & "IN ((Caja INNER JOIN Inmueble ON Caja.CodigoCaja = Inmueble.Caja) INNER JOIN Cpp ON " _
        & "Inmueble.CodInm = Cpp.CodInm) ON Proveedores.Codigo = Cpp.CodProv WHERE (((Caja.Codi" _
        & "goCaja)='99') AND ((Cpp.Estatus)='ASIGNADO')) " & strCadena

    Else    'Facturas asignadas por condominio
        ftnSQL = "SELECT Cpp.*,Proveedores.NombProv FROM Proveedores INNER JOIN Cpp ON Proveedo" _
            & "res.Codigo = Cpp.CodProv WHERE Cpp.estatus = 'ASIGNADO' and Cpp.CodInm = '" _
            & gcCodInm & "'" & strCadena
    '
    End If
    '
    End Function
    

    '---------------------------------------------------------------------------------------------
    Private Sub rtnMultiSel()    '-  Selección de Factura
    '---------------------------------------------------------------------------------------------
    'variables locales
    Dim I As Integer
    Dim j As Integer
    Dim K As Integer
    Dim intLinea As Integer
    Dim strAnterior As String
    Dim strCI As String
    Dim regMark As Variant
    Dim cnnTemp As New ADODB.Connection
    '
    Set RstGastos = New ADODB.Recordset
    '
    j = 0
    K = 0
    MshgAPago(2).Rows = 2
    txtAPago(3) = 0
    Call rtnLimpiar_Grid(MshgAPago(0))
    MshgAPago(0).Rows = 2
    '
    For Each regMark In DataGrid1.SelBookmarks  'hacer por cada fila marcada
    '
        With RstFacturasA   'Desplega la información básica del documento seleccionado
    '
            .AbsolutePosition = regMark
            If strAnterior = "" Then
                j = j + 1
                strAnterior = ftnAnterior(j)
                strCI = !CodInm
                cnnTemp.Open cnnOLEDB & gcPath & "\" & strCI & "\inm.mdb"
            Else
                If strAnterior = !codProv Then
                    If strCI <> !CodInm Then
                        cnnTemp.Close
                        cnnTemp.Open cnnOLEDB & gcPath & "\" & !CodInm & "\inm.mdb"
                    End If
                    j = j + 1
                    strAnterior = ftnAnterior(j)
                Else
                    DataGrid1.SelBookmarks.Remove (DataGrid1.SelBookmarks.Count - 1)
                    Exit Sub
                End If
            End If
            RstGastos.Open "SELECT * FROM Cargado WHERE Ndoc='" & !NDoc & "';", _
            cnnTemp, adOpenStatic, adLockReadOnly, adCmdText
            If Not RstGastos.EOF And Not RstGastos.BOF Then
                intLinea = ftnCargado(RstGastos, !CodInm, intLinea, !NDoc)
            Else
                MsgBox "Revise la asignación del gasto de la factura Nº " & !NDoc, _
                vbInformation, "Emisión de Cheques"
            End If
            RstGastos.Close
    '
        End With
    '
    Next
    
   MshgAPago(2).SetFocus
    '   ------------------------------------------------------------------------------------------
    
    Set RstGastos = Nothing
    On Error Resume Next
    cnnTemp.Close
    Set cnnTemp = Nothing
    ReDim CurDiferencia(1 To intLinea)
    For j = 1 To intLinea
        CurDiferencia(j) = CCur(MshgAPago(0).TextMatrix(j, 4))
    Next
    MshgAPago(0).Col = 1
    '
    End Sub
    
    '26/09/2002-----------------------------------------------------------------------------------
    Private Sub rtnRegistar()   '-  Registra el cheque actual, actualiza Chequera
    '---------------------------------------------------------------------------------------------
    '
    Dim I%
    If Not blnOtro Then Call RtnConsecutivo 'actualiza el número de cheque al guardar
    
    'Agrega el cheque a la tabla cheque
    cnnConexion.Execute "INSERT INTO Cheque(IDCheque,FechaCheque,Beneficiario,Banco,Cuenta,Conc" _
    & "epto,Usario,Fecha,Hora,IDEstado,Clave) VALUES(" & LblAPago(34) & ",Date(),'" & _
    txtAPago(2) & "','" & DtcAPago(1) & "','" & DtcAPago(2) & "','" & txtAPago(4) & "','" & _
    gcUsuario & "','" & MskFecha & "',Time(),-1,'" & CLng(LblAPago(34)) & DtcAPago(2) & "');"
    '
    'Agregar el registro en la relacion Cheque-factura
    I = 1
    With MshgAPago(2)
    '
        Do Until I > .Rows - 1 Or .TextMatrix(I, 0) = ""
            cnnConexion.Execute "INSERT INTO ChequeFactura(IDCheque,NDoc,Clave) " _
            & "VALUES (" & LblAPago(34) & ",'" & .TextMatrix(I, 0) & "','" & _
            CLng(LblAPago(34)) & DtcAPago(2) & "');"
            I = I + 1
        Loop
        '
    End With
    '
    With MshgAPago(1)   'Agrega los registros al detalle del cheque
        
        For I = 1 To .Rows - 1
            
            If .TextMatrix(I, 3) <> "" Then
                '
                cnnConexion.Execute "INSERT INTO ChequeDetalle(IDcheque,Cargado," _
                & "CodInm,CodGasto,Detalle,Monto,Clave,Ndoc) VALUES(" & LblAPago(34) & _
                ",'01-" & .TextMatrix(I, 2) & "','" & .TextMatrix(I, 4) & "','" & _
                .TextMatrix(I, 0) & "','" & .TextMatrix(I, 1) & "','" & _
                CCur(.TextMatrix(I, 3)) & "','" & CLng(LblAPago(34)) & DtcAPago(2) _
                & "','" & .TextMatrix(I, 6) & "');"
                'si el gasto es fondo lo agrega al mov. fondo
                If booFondo(.TextMatrix(I, 0)) Then
                    '
                    cnnConexion.Execute "INSERT INTO MovFondo(CodGasto,Fecha,Tipo,Periodo,Conce" _
                    & "pto,Debe,Haber) IN '" & gcPath & "\" & .TextMatrix(I, 4) & "\inm.mdb' VAL" _
                    & "UES('" & .TextMatrix(I, 0) & "',Date() ,'CH','01/" & .TextMatrix(I, 2) & _
                    "','" & .TextMatrix(I, 1) & " #" & LblAPago(34) & "','" & _
                    CCur(.TextMatrix(I, 3)) & "',0)"
                    'Actuliza el Saldo Actual----------------
                    cnnConexion.Execute "UPDATE Tgastos IN '" & gcPath & "\" & .TextMatrix(I, 4) _
                    & "\inm.mdb' SET SaldoActual = SaldoActual - '" & CCur(.TextMatrix(I, 3)) & _
                    "', Freg=DATE(),Usuario='" & gcUsuario & "' WHERE CodGasto='" & _
                    .TextMatrix(I, 0) & "';"
                    '
                End If
                '
            End If
            '
        Next
        '
    End With
    'actualiza los campos ref. al último pago efectuado al proveedor
    cnnConexion.Execute "Update Proveedores SET FecUltPag=Date(), UltPago='" & txtAPago(5) & _
    "', Deuda=Deuda-'" & txtAPago(5) & "' WHERE Codigo='" & RstFacturasA.Fields("CodProv") & "'"
    '
    Call rtnBitacora("Cheque Registrado #" & LblAPago(34) & "; Inm: " & gcCodInm)
    'Actualiza el útlimo de cheque de la chequera que se esta utilizando
    If Option1(0) Then
        Call finChequera(GtoCnn.Properties("Data Source Name"))
    Else
        Call finChequera(gcPath & "\" + sysCodInm + "\inm.mdb")
    End If
    '
    End Sub
    
    '---------------------------------------------------------------------------------------------
    Private Function ftnMonto() As Currency 'Devuelve el monto total de un cheque
    '---------------------------------------------------------------------------------------------
    '
    With MshgAPago(1)
        I = 1
        While I < .Rows
            If .TextMatrix(I, 3) <> "" Then
                ftnMonto = ftnMonto + CCur(.TextMatrix(I, 3))
            End If
            I = I + 1
        Wend
    End With
    '
    End Function
    
    '---------------------------------------------------------------------------------------------
    Private Sub rtnDif(curCelda As Currency, intRow As Integer) '
    '---------------------------------------------------------------------------------------------
    '
    If CCur(curCelda) <= CurDiferencia(intRow) Then  'EL MONTO A INTRODUCIR 'DEBE SER < AL RESTO
    '
        MshgAPago(0).TextMatrix(intRow, 4) = _
        Format(CurDiferencia(intRow) - curCelda, "#,##0.00")
        txtAPago(5) = Format(ftnMonto, "#,##0.00")
        LblAPago(32) = Format(ftnDif, "#,##0.00")
    Else    'SI MONTO > AL RESTO
    '
        MsgBox "Trata de introducir un monto mayor" _
        & vbCr & "al que fue registrado el gasto", vbExclamation, App.ProductName
        MshgAPago(1).Text = CurDiferencia(intRow)
        MshgAPago(0).TextMatrix(intRow, 4) = _
        Format(MshgAPago(1).Text, "#,##0.00")
        Exit Sub
    End If
    '
    End Sub
    
    '---------------------------------------------------------------------------------------------
    Private Function ftnDif() As Currency   '
    '---------------------------------------------------------------------------------------------
    On Error Resume Next
    With MshgAPago(0)
        For I = 1 To .Rows - 1
            ftnDif = ftnDif + CCur(.TextMatrix(I, 4))
        Next
    End With
    '
    End Function
    
'    '---------------------------------------------------------------------------------------------
'    Private Sub rtnConfig(intInd As Integer)    ' Configura la presentación del Grid
'    '---------------------------------------------------------------------------------------------
'    '
'    Call RtnCentraTitulo(intInd)
'    Select Case intInd
'    '
'        Case 0  'Distribución del Gastos Factura(s) seleccionada(s)
'    '   ---------------------
'            Call rtnCGrid(MshgAPago(0), 200)
'            MshgAPago(0).ColAlignment(0) = flexAlignCenterCenter
'            MshgAPago(0).ColAlignment(2) = flexAlignCenterCenter
'
'        Case 1  'Detalle cheque en emisión
'    '   ---------------------
'            MshgAPago(1).Width = fraApago(1).Width - (MshgAPago(1).Left * 2)
'            Call rtnCGrid(MshgAPago(1), 210)
'            MshgAPago(1).ColAlignment(1) = flexAlignLeftCenter
'            MshgAPago(1).ColAlignment(0) = flexAlignCenterCenter
'            MshgAPago(1).ColAlignment(2) = flexAlignCenterCenter
'
'        Case 2  'Selección de Factura(s)
'    '   ---------------------
'            Call rtnCGrid(MshgAPago(2), 300)
'            MshgAPago(2).ColAlignment(0) = flexAlignCenterCenter
'            MshgAPago(2).ColAlignment(1) = flexAlignCenterCenter
'
'        Case 4  'Grid Detalle Cheque emitido
'    '   ---------------------
'                Call rtnCGrid(MshgAPago(4), 500)
'                MshgAPago(4).ColAlignment(0) = flexAlignCenterCenter
'                MshgAPago(4).ColAlignment(1) = flexAlignCenterCenter
'    '
'    End Select
'
'    '
'    End Sub

    '---------------------------------------------------------------------------------------------
    Private Function ftnValida() As Boolean 'Verifica los datos necesarios para salvar un registro
    '---------------------------------------------------------------------------------------------
    '
    Dim strMsg$
    strMsg = "Imposible Guardar falta "
    For I = 1 To 3
        If DtcAPago(I) = "" Then ftnValida = MsgBox(strMsg & DtcAPago(I).ToolTipText, _
        vbExclamation, App.ProductName)
    Next
    If LblAPago(34) = "" Then ftnValida = MsgBox(strMsg & LblAPago(34).ToolTipText, _
        vbExclamation, App.ProductName)
    If txtAPago(4) = "" Then ftnValida = MsgBox(trmsg & txtAPago(4).ToolTipText, _
        vbExclamation, App.ProductName)
    If txtAPago(5) = "" Or Not txtAPago(5) > 0 Then ftnValida = MsgBox("Monto del Cheque es Nul" _
    & "o o igual a cero", vbExclamation, App.ProductName)
    '
    End Function

    '---------------------------------------------------------------------------------------------
    Private Sub rtnNewCan(blnYESNO As Boolean)  'Rutina Agregar / Cancelar Registro
    '---------------------------------------------------------------------------------------------
    '
    fraApago(1).Enabled = blnYESNO
    MshgAPago(0).Enabled = blnYESNO
    LblAPago(32) = IIf(blnYESNO = True, Format(ftnDif, "#,##0.00"), "0,00")
    txtAPago(4) = IIf(blnYESNO = True, "Doc. ", "")
    txtAPago(5) = "0,00"
    STabAPago.TabEnabled(0) = blnYESNO
    STabAPago.TabEnabled(1) = Not blnYESNO
    STabAPago.TabEnabled(2) = Not blnYESNO
    STabAPago.TabEnabled(3) = Not blnYESNO
    If blnYESNO = True Then
        cnnConexion.BeginTrans
        With MshgAPago(2)
            I = 1
            Do Until .TextMatrix(I, 0) = ""
                txtAPago(4) = txtAPago(4) + IIf(I = 1, .TextMatrix(I, 0) & " Fact. " & .TextMatrix(I, 1) & " " _
                & .TextMatrix(I, 2) & " / ", .TextMatrix(I, 0) & " Fact. " & .TextMatrix(I, 1) & " " & .TextMatrix(I, 2))
                I = I + 1
            Loop
            intRow = 0
            ReDim VecVinculo(MshgAPago(0).Rows, 2)
        End With
        STabAPago.tab = 0
        'MskFecha = Date
    Else
        cnnConexion.RollbackTrans
        STabAPago.tab = 1
    End If
    '
    End Sub
    
    '---------------------------------------------------------------------------------------------
    '
    '   Rutina:     rtnCheqE
    '
    '   Entrada:    variable entera intCase
    '
    '   Busca los cheques emitidos según parametros de usuario
    '---------------------------------------------------------------------------------------------

    Private Sub rtnCheqE(ByVal intCase%)
    '
    Dim strFecha$, StrBanco$, strCuenta$, strSQL$
    '
    If Not IsNull(dtpApago(1)) Then
        strFecha = " AND Cheque.FechaCheque=#" & Format(dtpApago(1), "mm-dd-yy") & "#"
    Else
        strFecha = ""
    End If
    '
    If DtcAPago(5) = "" Then
        StrBanco = ""
    Else
        StrBanco = " AND Cheque.Banco='" & DtcAPago(5) & "'"
    End If
    '
    If DtcAPago(6) = "" Then
        strCuenta = ""
    Else
        strCuenta = " AND Cheque.Cuenta='" & DtcAPago(6) & "'"
    End If
    If intCase <> 2 Then
        strSQL = "SELECT Format(Cheque.IDcheque,'000000') as ID, Format(Sum(ChequeDetalle.Monto" _
        & "),'#,##0.00') AS Monto, format(Cheque.FechaCheque,'dd/mm/yy') as Fecha,Cheque.Benefi" _
        & "ciario, Cheque.Banco, Cheque.Cuenta, Cheque.Concepto,Cheque.Impreso FROM Cheque INNE" _
        & "R JOIN ChequeDetalle ON Cheque.Clave = ChequeDetalle.Clave WHERE Cheque.Impres" _
        & "o=" & IIf(optApago(0) = True, "True", "False") & strFecha & StrBanco & strCuenta & _
        " GROUP BY Cheque.IDCheque, Cheque.FechaCheque, Cheque.Beneficiario, Cheque.Banco,Chequ" _
        & "e.Cuenta,Cheque.Concepto,Cheque.Impreso;"
    Else
        strSQL = "SELECT Format(IDcheque,'000000') as ID, format(Monto,'#,##0.00'), Format(Fech" _
        & "aCheque,'dd/mm/yy') as Fecha, Beneficiario, Banco, Cuenta, Concepto,0 FROM ChequeAnula" _
        & "do AS Cheque WHERE CodInm='" & gcCodInm & "'" & strFecha & StrBanco & strCuenta
        '
    End If
    Call RtnMuestra_Cheque(strSQL)
    '
    End Sub

    '---------------------------------------------------------------------------------------------
    Private Sub RtnCentraTitulo(intIndice As Integer)   'Centra Titulos de Columnas
    '---------------------------------------------------------------------------------------------
    '
    With MshgAPago(intIndice)
        .Col = 0
        .ColSel = .Cols - 1
        .FillStyle = flexFillRepeat
        .ColAlignmentFixed = flexAlignCenterCenter
        .FillStyle = flexFillSingle
    End With
    '
    End Sub
    
    '---------------------------------------------------------------------------------------------
    Private Function ftnAnterior(ByVal K As Integer) As String    '
    '---------------------------------------------------------------------------------------------
    '
    With RstFacturasA
        '.MoveFirst
        txtAPago(1) = !NombProv
        txtAPago(2) = !Benef
        MshgAPago(2).AddItem ("")
        'datos de la factura
        MshgAPago(2).TextMatrix(K, 0) = IIf(IsNull(!NDoc), "", !NDoc)
        MshgAPago(2).TextMatrix(K, 1) = IIf(IsNull(!Fact), "", !Fact)
        MshgAPago(2).TextMatrix(K, 2) = IIf(IsNull(!detalle), "", !detalle)
        MshgAPago(2).TextMatrix(K, 3) = Format(!Total, "#,##0.00")
        txtAPago(3) = Format(IIf(txtAPago(3) = "", 0, txtAPago(3)) + CCur(!Total), "#,##0.00")
        LblAPago(22) = txtAPago(3)
        MskFecha = .Fields("Fven")
        ftnAnterior = !codProv
        
        '
    End With
    '
    End Function

    '---------------------------------------------------------------------------------------------
    Private Sub rtnCGrid(grid As MSHFlexGrid, lngID As Long)
    '---------------------------------------------------------------------------------------------
    'configura la presentación del FlexGrid
    Dim I%, X%, j%
    Dim lngAncho&
    If grid.Index = 0 Or grid.Index = 1 Then
        X = grid.Cols - 2
        j = grid.Cols - 1
        grid.ColWidth(j) = 0
    Else
        X = grid.Cols - 1
        j = grid.Cols
    End If
    With grid
        For I = 0 To X
            lngAncho = lngID + I + j
            .TextArray(I) = LoadResString(lngID + I)
            .ColWidth(I) = CLng(LoadResString(lngAncho))
        Next
    End With
    End Sub
    
    Private Sub txtAPago_Change(Index As Integer)
    If Index = 12 Then
        If DtcAPago(8) <> "" Or DtcAPago(9) <> "" Then Call rtnConsecutivos
    End If
    End Sub

    '---------------------------------------------------------------------------------------------
    Private Sub txtAPago_KeyPress(Index%, KeyAscii%)    '
    '---------------------------------------------------------------------------------------------
    '
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    Select Case Index
        Case 6  'Número de Cheque Busqueda
    '   --------------------------
            Call Validacion(KeyAscii, "1234567890")
            
        Case 9
    '   --------------------------
    
            If KeyAscii = 46 Then KeyAscii = 44 'Convierte el punto en coma
            Call Validacion(KeyAscii, "1230456789.")
            If KeyAscii = 13 Then
                txtAPago(9) = Format(txtAPago(9), "#,##0.00")
                txtAPago(10).SetFocus
            End If

        Case 10
    '   --------------------------
            If KeyAscii = 46 Then KeyAscii = 44 'Convierte el punto en coma
            Call Validacion(KeyAscii, "1230456789.")
            If KeyAscii = 13 Then
                With txtAPago(10)
                    .Text = Format(.Text, "#,##0.00")
                    .SelStart = 0
                    .SelLength = Len(.Text)
                    .SetFocus
                End With
            End If
        
        Case 13
    '   --------------------------
        '
            If KeyAscii = 46 Then KeyAscii = 44 'Convierte el punto en coma
            Call Validacion(KeyAscii, "1234567890,")
            If KeyAscii = 13 Then
                If DtcAPago(8) <> "" And DtcAPago(9) <> "" Then Call rtnConsecutivos
                txtAPago(13) = Format(txtAPago(13), "#,##0.00")
            End If
        
        Case 14, 15, 16, 17
    '   --------------------------
            Call Validacion(KeyAscii, "1234567890")
            If KeyAscii = 13 Then If DtcAPago(8) <> "" And DtcAPago(9) <> "" Then _
            Call rtnConsecutivos
    '
    End Select
    '
    End Sub

    '---------------------------------------------------------------------------------------------
    Private Sub RtnImprimir_Cheque(IntOp%)    '
    '---------------------------------------------------------------------------------------------
    Dim strCheqI$, strCheqA$
    '
    With MshgAPago(3)
        '
        If IntOp = 0 Then
            I = 1
            While I < .Rows
                '
                If Option1(2).Value = True Then
                '
                    If .TextMatrix(I, 0) <> "" Then
                        strCheqA = ftnPrintCheque(CLng(.TextMatrix(I, 0)) & .TextMatrix(I, 5), _
                        1, CCur(.TextMatrix(I, 1)))
                        strCheqI = strCheqI & IIf(strCheqI = "", "'" & strCheqA & "'", " or Cla" _
                        & "ve='" & strCheqA & "'")
                    End If
                    '
                ElseIf Option1(3).Value = True Then
                    .Col = 0
                    .Row = I
                    If .CellPicture = imgPago Then
                        strCheqA = ftnPrintCheque(CLng(.TextMatrix(I, 0)) & .TextMatrix(I, 5), _
                        1, CCur(.TextMatrix(I, 1)))
                        strCheqI = strCheqI & IIf(strCheqI = "", "'" & strCheqA & "'", " or Cla" _
                        & "ve='" & strCheqA & "'")
                        Set .CellPicture = Nothing
                    End If
                    '
                End If
                I = I + 1
                '
            Wend
        
        If strCheqI <> "" Then
            'Actualiza el estatus del cheque
            cnnConexion.Execute "UPDATE Cheque SET Impreso=True,IDEstado=0 WHERE Clave = " & _
            strCheqI
            Adodc1(0).Refresh
            'Retardo para actualizar el ADODB.Recordset
            For I = 0 To 100000
            Next
            Call rtnDistribuye
        '
        End If
        '
    Else
        '
        If .TextMatrix(.RowSel, 0) <> "" Then
            If Respuesta("¿Confirma imprimir voucher cheque Nº " & .TextMatrix(.RowSel, 0) & _
            "?") Then strCheqA = ftnPrintCheque(CLng(.TextMatrix(.RowSel, 0)) & _
            .TextMatrix(.RowSel, 5), 0, CCur(.TextMatrix(.RowSel, 1)))
        End If
        '
    End If
    '
    End With
    '
    End Sub
    
    '---------------------------------------------------------------------------------------------
    Private Sub rtnBusca_Cheque()   '
    '---------------------------------------------------------------------------------------------
    '
    Dim strBusc$, strMont$, strSQL$
    '---------
    If Not txtAPago(6) = "" Then
        strBusc = "Cheque.IDcheque=" & txtAPago(6)
    Else
        strBusc = ""
    End If
    '---------
    If Not txtAPago(7) = "" Then
        strBusc = strBusc & IIf(strBusc = "", "", " AND ") & "Banco='" & txtAPago(7) & "'"
    Else
        strBusc = strBusc
    End If
    '---------
    If Not txtAPago(8) = "" Then
        strBusc = strBusc & IIf(strBusc = "", "", " AND ") & "Beneficiario='" & txtAPago(8) & "'"
    Else
        strBusc = strBusc
    End If
    '---------
    If chkPago(0) = 1 Then
        strBusc = strBusc & IIf(strBusc = "", "", " AND ") & "FechaCheque Between #" _
        & Format(dtpApago(2), "mm/dd/yy") & "# AND #" _
        & Format(dtpApago(3), "mm/dd/yy") & "#"
    Else
        strBusc = strBusc
    End If
    '---------
    If strBusc = "" Then Exit Sub
    If chkPago(1) = 1 Then
        strMont = "HAVING Sum(ChequeDetalle.Monto) Between " & CCur(txtAPago(9)) & " And " _
        & CCur(txtAPago(10)) & ";"
    Else
        strMonto = ";"
    End If
    '---------
    '
    strSQL = "SELECT Format(Cheque.IDcheque,'000000') as ID, Format(Sum(ChequeDetalle.Monto),'#" _
    & ",##0.00') AS Monto, Format(Cheque.FechaCheque,'dd/mm/yy') as Fecha,Cheque.Beneficiario, " _
    & "Cheque.Banco, Cheque.Cuenta, Cheque.Concepto,Cheque.Impreso FROM Cheque INNER JOIN Chequ" _
    & "eDetalle ON Cheque.Clave = ChequeDetalle.Clave WHERE " & strBusc & " GROUP BY Cheq" _
    & "ue.IDCheque, Cheque.FechaCheque, Cheque.Beneficiario, Cheque.Banco,Cheque.Cuenta,Cheque." _
    & "Concepto,cheque.Impreso " & strMont
    '
    Call RtnMuestra_Cheque(strSQL)
    '
    End Sub
    
    '---------------------------------------------------------------------------------------------
    Private Sub RtnMuestra_Cheque(strSour As String)
    '---------------------------------------------------------------------------------------------
    '
    Dim grid As MSHFlexGrid
    Adodc1(0).RecordSource = strSour
    Adodc1(0).Refresh
    '
    Call rtnDistribuye
    '
    End Sub
    
    '---------------------------------------------------------------------------------------------
    '   Rutina:     rtnBusca_Cuenta
    '
    '   Busca cuentas asociadas al inmueble o datos de las cuentas
    '   del la administradora en caso de que el inmueble se maneje
    '   a través de la cuenta pote
    '---------------------------------------------------------------------------------------------
    Private Sub rtnBusca_Cuenta(strData$)
    '
    Set RstBco = New ADODB.Recordset
    RstBco.ActiveConnection = cnnOLEDB & strData
    RstBco.Open "SELECT Cuentas.IDcuenta, Bancos.NombreBanco, Cuentas.NumCuenta FROM Bancos " _
    & "INNER JOIN Cuentas ON Bancos.IDBanco=Cuentas.IDBanco WHERE Cuentas.Inactiva=False ORDER " _
    & "BY Pred,Cuentas.IDCuenta", RstBco.ActiveConnection, adOpenKeyset, adLockReadOnly
    '
        If RstBco.EOF Then
            
            MsgBox "No existe información sobre las cuentas y chequeras" & vbCrLf & "asociadas a este condominio," _
            & vbCrLf & "Para poder continuar debe registrarlas." & vbCrLf & "Desea hacerlo ahora?", _
            vbYesNo + vbInformation, App.ProductName
        
        Else
        
            Dim I%
            For I = 1 To 2  'llena la lista de # Cuenta / Banco
                
                Set DtcAPago(I).RowSource = RstBco  'Ficha Datos genrales
                Set DtcAPago(I + 4).RowSource = RstBco  'Ficha Lista Pagos
                Set DtcAPago(I + 7).RowSource = RstBco  'Ficha Consecutivos
                '------------------------------------------------------------------------------------------------
                'estas lineas se pueden activar para futuros programas
                'buscan una cuenta predeterminada al comenzar el formulario
                '
                'DtcAPago(i).Text = RstBco.Fields(i)
                DtcAPago(I + 4).Text = RstBco.Fields(I)
            Next
            DtcAPago(3) = ""
            LblAPago(34) = ""
            'Call RtnChqrs   'Llama al procedimiento buscar chequera asignadas
        End If
    End Sub
    
    Private Sub rtnSelOPT()
    For I = 0 To 2
        If optApago(I) Then
            Call rtnCheqE(I)
            Exit For
        End If
    Next
    End Sub

    '---------------------------------------------------------------------------------------------
    '   Rutina:     finchequera
    '
    '   Entrada:    cadena de conexion al origen de datos
    '
    '   Actualiza el cheque utilizado, verifica el final de la chequera para
    '   avizar al ususario que debe registrar o asignar una nueva chequera
    '---------------------------------------------------------------------------------------------
    Private Sub finChequera(strLocal$)
    '
    cnnConexion.Execute "UPDATE Chequera IN '" & strLocal & "' SET Ultimo='" & LblAPago(34) _
    & "' WHERE IDchequera=" & DtcAPago(3) & ";"
    'Recordset Local
    Dim rstlocal As New ADODB.Recordset
    
    rstlocal.Open "SELECT * FROM Chequera WHERE IDchequera=" & Trim(DtcAPago(3)), _
    cnnOLEDB & strLocal, adOpenKeyset, adLockOptimistic
    
    LblAPago(34) = LblAPago(34) + 1
    
    If rstlocal!Hasta <= (rstlocal!Ultimo + 1) Then  'Se utilizó el último cheque
        
        MsgBox "Sr(a): " & gcUsuario & vbCrLf & vbCrLf & _
        "Chequera Terminada." & vbCrLf & "Registre o asigne una nueva chequera", _
        vbInformation, App.ProductName
        'Desactiva la chequera en uso
        cnnConexion.Execute "UPDATE AsignaChequera IN '" & strLocal & "' SET Activa=FALSE WHERE" _
        & " IDChequera=" & DtcAPago(3) & ";"
        Call rtnBusca_Cuenta(strLocal)
        Call rtnBitacora("FIN CHEQUERA N° " & DtcAPago(3) & " Inm: " & gcCodInm)
        DtcAPago(3) = "": LblAPago(34) = ""
        DtcAPago(1) = ""
        DtcAPago(2) = ""
    ElseIf rstlocal!Ultimo = rstlocal!Ultimo + 10 Then
        MsgBox "Quedan apróx. 10 Cheques de esta chequera", vbInformation, gcUsuario
    End If
    rstlocal.Close
    Set rstlocal = Nothing
    '
    End Sub
    
    '---------------------------------------------------------------------------------------------
    '
    '   Rutina:     rtnDistribuye
    '
    '
    '---------------------------------------------------------------------------------------------
    Private Sub rtnDistribuye()
    '
    Call rtnLimpiar_Grid(MshgAPago(4))
    Call rtnLimpiar_Grid(MshgAPago(3))
    With MshgAPago(3)
        If Adodc1(0).Recordset.RecordCount > 0 Then
            .Rows = Adodc1(0).Recordset.RecordCount + 1
            Adodc1(0).Recordset.MoveFirst
            I = 1
            Do Until Adodc1(0).Recordset.EOF
               For j = 0 To 7
                .TextMatrix(I, j) = Adodc1(0).Recordset.Fields(j)
               Next
               I = I + 1
               Adodc1(0).Recordset.MoveNext
            Loop
        Else
            .Rows = 2
        End If
        .Row = 1
        .Col = 0
        .ColSel = .Cols - 1
        With MshgAPago(4)
            .Row = 1
            .Col = 0
            .ColSel = .Cols - 1
        End With
        '
    End With
    '
    End Sub

    '---------------------------------------------------------------------------------------------
    '   Funcion:    ftnCargado
    '
    '   Buca información sobre el cargado del documento respecitvo,
    '   además de pagos emitidos a/c del mismo documento
    '   Devuelve un valor entero indicanco la cantidad de items que corresponde
    '   al cargado del documento seleccionado
    '---------------------------------------------------------------------------------------------
    Private Function ftnCargado(rst As ADODB.Recordset, strInm$, Z%, strDoc$) As Integer
    '
    Dim abono As Currency
    
    rst.MoveFirst
    Do Until rst.EOF  'Hacer mientras no sea fin de archivo
        Z = Z + 1
        If Z >= 1 Then
        
            If Z > (MshgAPago(0).Rows - 1) Then MshgAPago(0).AddItem ("")
            If MshgAPago(0).TextMatrix(Z, 0) = rst!codGasto Then
                If MshgAPago(0).TextMatrix(Z, 1) = rst!detalle Then
                    If MshgAPago(0).TextMatrix(Z, 2) = Format(rst!Periodo, "MM-YY") Then
                        abono = MshgAPago(0).TextMatrix(Z, 3): Z = Z - 1
                    End If
                Else
                    abono = 0
                End If
            Else
                abono = 0
            End If
        End If
    '      -----------------------llena el Grid de Cargados
        With MshgAPago(0)
        
            If Z > (.Rows - 1) Then .AddItem ("")
            .TextMatrix(Z, 0) = rst!codGasto
            .TextMatrix(Z, 1) = rst!detalle
            .TextMatrix(Z, 2) = Format(rst!Periodo, "MM-YY")
            .TextMatrix(Z, 3) = Format(rst!Monto + abono, "#,##0.00")
            .TextMatrix(Z, 4) = Format(Saldo(rst!codGasto, strDoc, rst!Monto + abono, _
            Format(rst!Periodo, "mm-dd-yy"), rst!detalle), "#,##0.00")
            .TextMatrix(Z, 5) = strInm
            .TextMatrix(Z, 6) = rst!NDoc
        End With
        'Aqui busca los abonos que tenga esta factura
        
        rst.MoveNext
    '
        Loop    'Punto de control
    '  -------------------------------------
    ftnCargado = Z
    MshgAPago(1).Rows = MshgAPago(0).Rows
    '
    End Function

    '---------------------------------------------------------------------------------------------
    '   Function:   Saldo
    '
    
    '   Devuelve el saldo de un determinado Gasto
    '---------------------------------------------------------------------------------------------
    Function Saldo(strC$, strD$, Monto1@, Periodo$, D$) As Currency
    '
    Dim rst2 As New ADODB.Recordset   'variables locales
    '
    rst2.Open "SELECT Sum(Monto) AS Total From CHequeDetalle " _
    & "WHERE IDCheque In (SELECT IDCheque FROM ChequeFactura WHERE Ndoc='" _
    & strD & "') AND CodGasto='" & strC & "' AND Cargado=#" & Periodo & "# AND Detalle='" & _
    D & "';", cnnConexion, adOpenStatic, adLockOptimistic, adCmdText
    
    If rst2.RecordCount > 0 Then
        Saldo = Monto1 - IIf(IsNull(rst2!Total), 0, rst2!Total)
    Else
        Saldo = Monto1
    End If
    rst2.Close
    Set rst2 = Nothing
    
    '
    End Function

    Private Sub Fin_Emitir()
    '
    For I = 0 To 2
        MshgAPago(I).Rows = 2
        Call rtnLimpiar_Grid(MshgAPago(I))
    Next
   With STabAPago
        .TabEnabled(0) = False
        .TabEnabled(1) = True
        .TabEnabled(2) = True
        .TabEnabled(3) = True
        .tab = 1
    End With
    blnOtro = False
    Call RtnEstado(6, Toolbar1)
    Toolbar1.Buttons(9).Enabled = False
    End Sub

    '---------------------------------------------------------------------------------------------
    '
    '   Rutina: rtnConcesutivos
    '
    '   Busca todos los cheques emitidos de la cuenta señalada,según paràmetos
    '   enviados por el usuario
    '
    '---------------------------------------------------------------------------------------------
    Private Sub rtnConsecutivos(Optional Criterio$)
    'Variables locales
    Dim rstCons As New ADODB.Recordset
    Dim strSQL$, strCriterio$ ', Criterio As String
    Dim I&
    '
    Call rtnLimpiar_Grid(MshgAPago(6))
    Call rtnLimpiar_Grid(MshgAPago(7))
    
    If Not DtcAPago(8) = "" Then strCriterio = "C.Banco='" & DtcAPago(8) & "'"
    
    If Not DtcAPago(9) = "" Then
        strCriterio = IIf(strCriterio = "", "", strCriterio & " AND ") & "C.Cuenta='" & DtcAPago(9) & "'"
    End If
    If Not txtAPago(16) = "" Then
        strCriterio = IIf(strCriterio = "", "", strCriterio & " AND ") & "C.IDCheque=" & txtAPago(16) & ""
    End If
    If Not IsNull(dtpApago(5)) And Not IsNull(dtpApago(6)) Then
        strCriterio = IIf(strCriterio = "", "", strCriterio & " AND ") & "C.FechaCheque Between #" & _
        Format(dtpApago(5), "mm/dd/yyyy") & "# AND #" & Format(dtpApago(6), "mm/dd/yyyy") & "#"
    ElseIf Not IsNull(dtpApago(5)) Then
        strCriterio = IIf(strCriterio = "", "", strCriterio & " AND ") & "C.FechaCheque>=#" & _
        Format(dtpApago(5), "mm/dd/yyyy") & "#"
    ElseIf Not IsNull(dtpApago(6)) Then
        strCriterio = IIf(strCriterio = "", "", strCriterio & " AND ") & "C.FechaCheque<=#" & _
        Format(dtpApago(6), "mm/dd/yyyy") & "#"
    End If
    If Not txtAPago(13) = "" Then
        strCriterio = IIf(strCriterio = "", "", strCriterio & " AND ") & "Monto=" & Replace(CCur(txtAPago(13)), ",", ".")
    End If
    If Not txtAPago(12) = "" Then
        strCriterio = IIf(strCriterio = "", "", strCriterio & " AND ") & "C.Beneficiario LIKE '%" & _
        txtAPago(12) & "%'"
    End If
    '
    If strCriterio = "" Then Exit Sub
    'Criterio adicional
    If DtcAPago(0) <> "" Then
        Criterio = "01/" & DtcAPago(0)
        Criterio = Format(Criterio, "mm/dd/yyyy")
        Criterio = " AND CD.Cargado=#" & Criterio & "#"
    End If
    If txtAPago(15) <> "" Then
        Criterio = Criterio & " AND CD.CodGasto='" & Trim(txtAPago(15)) & "'"
    End If
    If txtAPago(14) <> "" Then
        Criterio = Criterio & " AND CD.CodInm='" & Trim(txtAPago(14)) & "'"
    End If
    If txtAPago(17) <> "" Then
        Criterio = Criterio & " AND C.Clave IN (SELECT Clave FROM ChequeFactura WHERE Ndoc ='" & Trim(txtAPago(17)) & "')"
    End If
    '
    If Criterio = "" Then
        '
        strSQL = "SELECT Format(C.IDcheque,'000000') as ID,Format(Sum(CD.Monto),'#,##0.00 ') AS" _
        & " Monto, Format(C.FechaCheque,'DD/MM/yy') as Fecha,C.Beneficiario, C.Banco, C.Cuenta," _
        & " C.Concepto FROM Cheque AS C INNER JOIN ChequeDetalle AS CD ON C.Clave = CD.Clave" _
        & " WHERE " & strCriterio & Criterio & " GROUP BY C.IDCheque, C.FechaCheque, C.Benef" _
        & "iciario, C.Banco,C.Cuenta,C.Concepto UNION SELECT Format(IDCheque,'000000'),Format(M" _
        & "onto,'#,##0.00 '),Format(FechaCheque,'dd/mm/yy'),Beneficiario,Banco,Cuenta,'ANULADO' F" _
        & "ROM ChequeAnulado as C WHERE " & strCriterio & " ORDER BY ID DESC;"
        '
        Dim intPos%
        intPos = InStr(strCriterio, "%")
        If intPos > 0 Then
            Do
                strCriterio = Left(strCriterio, intPos - 1) & "*" & Right(strCriterio, Len(strCriterio) - intPos)
                intPos = InStr(strCriterio, "%")
            Loop Until intPos = 0
        End If
        strCtvo = "SELECT C.IDcheque,Sum(CD.Monto) AS Monto, C.FechaCheque,C.Beneficiari" _
        & "o, C.Banco, C.Cuenta,C.Concepto FROM Cheque AS C INNER JOIN ChequeDetalle AS CD ON C" _
        & ".Clave = CD.Clave WHERE " & strCriterio & Criterio & " GROUP BY C.IDCheque, C." _
        & "FechaCheque, C.Beneficiario, C.Banco,C.Cuenta,C.Concepto UNION SELECT IDCheque,Monto" _
        & ",FechaCheque,Beneficiario,Banco,Cuenta,'ANULADO' FROM ChequeAnulado as C WHERE " & _
        strCriterio & " ORDER BY C.IDcheque DESC;"
        
    Else
        strSQL = "SELECT Format(C.IDcheque,'000000') as ID,Format(Sum(CD.Monto),'#,##0.00 ') AS" _
        & " Monto, Format(C.FechaCheque,'DD/MM/yy') as Fecha,C.Beneficiario, C.Banco, C.Cuenta," _
        & "C.Concepto FROM Cheque AS C INNER JOIN ChequeDetalle AS CD ON C.Clave = CD.Clave" _
        & " WHERE " & strCriterio & Criterio & " GROUP BY C.IDCheque, C.FechaCheque, C.Benefi" _
        & "ciario, C.Banco,C.Cuenta,C.Concepto ORDER BY C.IDCheque DESC;"
        
        strCtvo = "SELECT C.IDcheque ,Sum(CD.Monto) AS Monto,C.FechaCheque ,C.Bene" _
        & "ficiario, C.Banco, C.Cuenta,C.Concepto FROM Cheque AS C INNER JOIN ChequeDetalle AS " _
        & "CD ON C.Clave = CD.Clave WHERE " & strCriterio & Criterio & " GROUP BY C.IDChe" _
        & "que, C.FechaCheque, C.Beneficiario, C.Banco,C.Cuenta,C.Concepto ORDER BY C.IDCheque DESC;"
        
    End If
    '
    rstCons.Open strSQL, cnnConexion, adOpenKeyset, adLockOptimistic, adCmdText
    'Set MshgAPago(6).DataSource = rstCons
    With rstCons
        If Not .EOF And Not .BOF Then
            .MoveFirst
            MshgAPago(6).Rows = .RecordCount + 1
            I = 1
            Do
                'MshgAPago(6).TextMatrix(I, 0) = !ID
                'MshgAPago(6).TextMatrix(I, 1) = !Monto
                For j = 0 To .Fields.Count - 1
                    MshgAPago(6).TextMatrix(I, j) = .Fields(j)
                Next
                I = I + 1
                .MoveNext
            Loop Until .EOF
        End If
    End With
    'MshgAPago(6).Refresh
    If rstCons.RecordCount > 0 Then
        Dim rstLis As New ADODB.Recordset
        'Configura la alineación del texto
        MshgAPago(6).ColAlignment(0) = flexAlignCenterCenter
        MshgAPago(6).ColAlignment(1) = flexAlignRightCenter
        MshgAPago(6).ColAlignment(2) = flexAlignCenterCenter
        
        strSQL = "SELECT DISTINCT UCASE(Format(Cargado,'MMM / YYYY')) as P FROM ChequeDetalle W" _
        & "HERE Clave IN (SELECT Clave FROM Cheque as C WHERE (((" & strCriterio & "))));"
        
        rstLis.Open strSQL, cnnConexion, adOpenKeyset, adLockOptimistic, adCmdText
        Set DtcAPago(0).RowSource = rstLis
        If IsNumeric(MshgAPago(6).TextMatrix(MshgAPago(6).RowSel, 0)) Then
            Call busca_detalle(MshgAPago(6).TextMatrix(MshgAPago(6).RowSel, 0))
        End If
        'Call busca_detalle(rstCons.Fields(0))
    End If
    rstCons.Close
    Set rstCons = Nothing
    'If Not MshgAPago(6).RowSel = 0 Or Not MshgAPago(6).TextMatrix(MshgAPago(6).RowSel, 0) = "" _
    Then Call busca_detalle(MshgAPago(6).TextMatrix(MshgAPago(6).RowSel, 0))
    '
    End Sub


    '---------------------------------------------------------------------------------------------
    '   Rutina: busca_detalle
    '
    '   Busaca el detalle del cargado del cheque señalado en el paràmetro
    '---------------------------------------------------------------------------------------------
    Private Sub busca_detalle(Cheque As Long)
    '
    'Variables locales
    Dim rstDet As New ADODB.Recordset
    Dim strSQL$, Linea&
    strSQL = "SELECT CodInm & ' - ' &  CodGasto as Cuenta,Cargado,Detalle,Monto FROM ChequeDeta" _
    & "lle WHERE IDCheque=" & Cheque & ";"
    '
    Call rtnLimpiar_Grid(MshgAPago(7))
    With rstDet
        .Open strSQL, cnnConexion, adOpenKeyset, adLockOptimistic, adCmdText
        
        If Not .EOF Or Not .BOF Then
            MshgAPago(7).Rows = .RecordCount + 1
            .MoveFirst
            For Linea = 1 To .RecordCount
                MshgAPago(7).TextMatrix(Linea, 0) = !Cuenta
                MshgAPago(7).TextMatrix(Linea, 1) = UCase(Format(!cargado, "mmm-yyyy"))
                MshgAPago(7).TextMatrix(Linea, 2) = !detalle
                MshgAPago(7).TextMatrix(Linea, 3) = Format(!Monto, "#,##0.00 ")
                .MoveNext
            Next Linea
        Else
            MshgAPago(7).Rows = 2
        End If
    .Close
    End With
    Set rstDet = Nothing
    MshgAPago(7).ColAlignment(0) = flexAlignCenterCenter
    MshgAPago(7).ColAlignment(1) = flexAlignCenterCenter
    MshgAPago(7).ColAlignment(2) = flexAlignLeftCenter
    MshgAPago(7).ColAlignment(3) = flexAlignRightCenter
    End Sub

    '---------------------------------------------------------------------------------------------
    '   Function:   Filtrar
    '
    '   Devuelve una variable tipo cadena, que representa el titulo del informe
    '---------------------------------------------------------------------------------------------
    Function Filtrar() As String
    '
    If Not IsNull(dtpApago(5)) And Not IsNull(dtpApago(6)) Then
        Filtrar = "Desde: " & dtpApago(5) & " Hasta: " & dtpApago(6)
    ElseIf Not IsNull(dtpApago(5)) Then
        Filtrar = "Desde: " & dtpApago(5)
    ElseIf Not IsNull(dtpApago(6)) Then
        filtar = "Hasta: " & dtpApago(6)
    End If
    '
    If Not txtAPago(12) = "" Then
        Filtrar = IIf(Filtrar = "", "", ";") & "Proveedor como= " & txtAPago(12)
    End If
    If Not DtcAPago(0) = "" Then
        Filtrar = IIf(Filtrar = "", "", ";") & "Cagado a: " & DtcAPago(0)
    End If
    If Not txtAPago(13) = "" Then
        Filtrar = IIf(Filtrar = "", "", ";") & "Monto= " & txtAPago(13)
    End If
    If Not txtAPago(14) = "" Then
        Filtrar = IIf(Filtrar = "", "", ";") & "Cod.Inm= " & txtAPago(14)
    End If
    If Not txtAPago(14) = "" Then
        Filtrar = IIf(Filtrar = "", "", ";") & "Cod.Gasto= " & txtAPago(15)
    End If
    '
    End Function

    Private Function LoadResCustom(ID, Optional Tipo) As Variant
    'variables locales
    Dim bytes() As Byte, idf As Integer
    
    idf = FreeFile
    Open App.Path & "\ressac.tmp" For Binary As #idf
    bytes = LoadResData(ID, Tipo)
    Put #idf, , bytes
    Close #idf
    Set LoadResCustom = LoadPicture(App.Path & "\ressac.tmp")
    
    End Function
    
