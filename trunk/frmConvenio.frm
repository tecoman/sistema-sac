VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmConvenio 
   Caption         =   "..::::Convenimiento de Pago:::::.."
   ClientHeight    =   8490
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11400
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8490
   ScaleWidth      =   11400
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmd 
      Caption         =   "Guardar"
      Height          =   975
      Index           =   2
      Left            =   10920
      Style           =   1  'Graphical
      TabIndex        =   63
      Top             =   8895
      Visible         =   0   'False
      Width           =   1215
   End
   Begin MSAdodcLib.Adodc adoCon 
      Height          =   330
      Index           =   2
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   2850
      _ExtentX        =   5027
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
      Caption         =   "Estatus Cobranza"
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
   Begin VB.CommandButton cmd 
      Caption         =   "Imprimir"
      Height          =   975
      Index           =   1
      Left            =   12165
      Style           =   1  'Graphical
      TabIndex        =   43
      Top             =   8895
      Width           =   1215
   End
   Begin VB.CommandButton cmd 
      Caption         =   "Salir"
      Height          =   975
      Index           =   0
      Left            =   13425
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   8895
      Width           =   1215
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   10065
      Left            =   135
      TabIndex        =   7
      Top             =   165
      Width           =   14760
      _ExtentX        =   26035
      _ExtentY        =   17754
      _Version        =   393216
      Tabs            =   2
      TabHeight       =   520
      ShowFocusRect   =   0   'False
      TabCaption(0)   =   "Datos Generales"
      TabPicture(0)   =   "frmConvenio.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "adoCon(1)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "adoCon(0)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "fra(0)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "fra(1)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "fra(2)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "fra(3)"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).ControlCount=   6
      TabCaption(1)   =   "Convenio de Pago"
      TabPicture(1)   =   "frmConvenio.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "lbl(6)"
      Tab(1).Control(1)=   "Shape1(0)"
      Tab(1).Control(2)=   "Shape1(1)"
      Tab(1).Control(3)=   "lbl(7)"
      Tab(1).Control(4)=   "lbl(8)"
      Tab(1).Control(5)=   "LIN(0)"
      Tab(1).Control(6)=   "LIN(1)"
      Tab(1).Control(7)=   "lbl(9)"
      Tab(1).Control(8)=   "lbl(10)"
      Tab(1).Control(9)=   "LIN(2)"
      Tab(1).Control(10)=   "lbl(11)"
      Tab(1).Control(11)=   "Shape1(2)"
      Tab(1).Control(12)=   "Shape1(3)"
      Tab(1).Control(13)=   "lbl(12)"
      Tab(1).Control(14)=   "lbl(13)"
      Tab(1).Control(15)=   "lbl(14)"
      Tab(1).Control(16)=   "lbl(15)"
      Tab(1).Control(17)=   "lbl(16)"
      Tab(1).Control(18)=   "lbl(17)"
      Tab(1).Control(19)=   "lbl(18)"
      Tab(1).Control(20)=   "lbl(19)"
      Tab(1).Control(21)=   "lbl(27)"
      Tab(1).Control(22)=   "lbl(28)"
      Tab(1).Control(23)=   "LIN(3)"
      Tab(1).Control(24)=   "lbl(29)"
      Tab(1).Control(25)=   "lbl(30)"
      Tab(1).Control(26)=   "Shape1(4)"
      Tab(1).Control(27)=   "lbl(31)"
      Tab(1).Control(28)=   "Dat(5)"
      Tab(1).Control(29)=   "Dat(4)"
      Tab(1).Control(30)=   "GRID"
      Tab(1).Control(31)=   "fra(4)"
      Tab(1).Control(32)=   "TXT(10)"
      Tab(1).Control(33)=   "TXT(11)"
      Tab(1).ControlCount=   34
      Begin VB.TextBox TXT 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   11
         Left            =   -65460
         Locked          =   -1  'True
         TabIndex        =   59
         Text            =   "0,00"
         ToolTipText     =   "Honorarios Extrajudiciales"
         Top             =   8745
         Visible         =   0   'False
         Width           =   1785
      End
      Begin VB.TextBox TXT 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   10
         Left            =   -65460
         Locked          =   -1  'True
         TabIndex        =   58
         Text            =   "0,00"
         ToolTipText     =   "Gastos de Cobranza"
         Top             =   8370
         Visible         =   0   'False
         Width           =   1785
      End
      Begin VB.Frame fra 
         Caption         =   "Calculadora:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   6090
         Index           =   4
         Left            =   -64005
         TabIndex        =   44
         Top             =   2040
         Width           =   3480
         Begin VB.TextBox TXT 
            Alignment       =   1  'Right Justify
            Height          =   315
            Index           =   9
            Left            =   1395
            Locked          =   -1  'True
            TabIndex        =   52
            Text            =   "0,00"
            ToolTipText     =   "Deuda del Convenio"
            Top             =   255
            Width           =   1785
         End
         Begin VB.TextBox TXT 
            Alignment       =   1  'Right Justify
            Height          =   315
            Index           =   8
            Left            =   2040
            TabIndex        =   50
            Text            =   "0"
            ToolTipText     =   "Nº Cuotas"
            Top             =   1560
            Width           =   1140
         End
         Begin VB.TextBox TXT 
            Alignment       =   1  'Right Justify
            Height          =   315
            Index           =   7
            Left            =   1410
            Locked          =   -1  'True
            TabIndex        =   49
            Text            =   "0,00"
            Top             =   1125
            Width           =   1770
         End
         Begin VB.TextBox TXT 
            Alignment       =   1  'Right Justify
            Height          =   315
            Index           =   6
            Left            =   1395
            TabIndex        =   48
            Text            =   "0,00"
            ToolTipText     =   "Monto Inicial"
            Top             =   690
            Width           =   1785
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid GRID1 
            Height          =   3690
            Left            =   225
            TabIndex        =   51
            Tag             =   "300|1000|1200"
            Top             =   2160
            Width           =   3075
            _ExtentX        =   5424
            _ExtentY        =   6509
            _Version        =   393216
            Cols            =   3
            FixedCols       =   0
            BackColorFixed  =   -2147483646
            ForeColorFixed  =   -2147483639
            BackColorBkg    =   -2147483636
            GridColor       =   -2147483637
            MergeCells      =   2
            FormatString    =   "^Nº |^FECHA |>MONTO"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial Narrow"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            _NumberOfBands  =   1
            _Band(0).Cols   =   3
         End
         Begin VB.Label lbl 
            Caption         =   "DEUDA:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   26
            Left            =   285
            TabIndex        =   53
            Top             =   330
            Width           =   795
         End
         Begin VB.Label lbl 
            Caption         =   "INICIAL:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   20
            Left            =   285
            TabIndex        =   47
            Top             =   765
            Width           =   795
         End
         Begin VB.Label lbl 
            Caption         =   "Nº CUOTAS:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   21
            Left            =   285
            TabIndex        =   46
            Top             =   1635
            Width           =   930
         End
         Begin VB.Label lbl 
            Caption         =   "RESTA:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   25
            Left            =   285
            TabIndex        =   45
            Top             =   1200
            Width           =   795
         End
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid GRID 
         Height          =   5475
         Left            =   -74625
         TabIndex        =   29
         Tag             =   "1200|1200|1700|1700|1700|1700"
         Top             =   2670
         Width           =   10500
         _ExtentX        =   18521
         _ExtentY        =   9657
         _Version        =   393216
         Cols            =   6
         FixedCols       =   0
         BackColorFixed  =   -2147483646
         ForeColorFixed  =   -2147483639
         BackColorBkg    =   -2147483636
         GridColor       =   -2147483637
         MergeCells      =   2
         FormatString    =   "^Nº MES |^PERIODO |>MONTO |>G. COBRANZA |>HONORARIOS |>TOTAL"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _NumberOfBands  =   1
         _Band(0).Cols   =   6
      End
      Begin VB.Frame fra 
         Caption         =   "Opciones:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3600
         Index           =   3
         Left            =   10350
         TabIndex        =   20
         Top             =   4560
         Width           =   4125
         Begin VB.TextBox TXT 
            Alignment       =   1  'Right Justify
            Height          =   315
            Index           =   5
            Left            =   1245
            TabIndex        =   42
            Text            =   "0,00"
            ToolTipText     =   "Deducciones Honorarios"
            Top             =   2850
            Width           =   2475
         End
         Begin VB.TextBox TXT 
            Alignment       =   1  'Right Justify
            Height          =   315
            Index           =   4
            Left            =   1245
            TabIndex        =   41
            Text            =   "0,00"
            ToolTipText     =   "Deducciones Gastos"
            Top             =   2025
            Width           =   2475
         End
         Begin VB.CheckBox chk 
            Caption         =   "Honorarios Extrajudiciales"
            Height          =   385
            Index           =   1
            Left            =   240
            TabIndex        =   22
            Top             =   795
            Width           =   3120
         End
         Begin VB.CheckBox chk 
            Caption         =   "Gastos de Cobranza"
            Height          =   385
            Index           =   0
            Left            =   240
            TabIndex        =   21
            Top             =   405
            Width           =   3120
         End
         Begin VB.Label lbl 
            Caption         =   "HONORARIOS EXTRAJUDICIALES"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   24
            Left            =   600
            TabIndex        =   40
            Top             =   2565
            Width           =   2490
         End
         Begin VB.Label lbl 
            Caption         =   "GASTOS DE COBRANZA"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   23
            Left            =   600
            TabIndex        =   39
            Top             =   1695
            Width           =   2085
         End
         Begin VB.Label lbl 
            Caption         =   " DEDUCCIONES:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   22
            Left            =   345
            TabIndex        =   38
            Top             =   1260
            Width           =   1305
         End
         Begin VB.Shape Shape1 
            BorderColor     =   &H8000000C&
            Height          =   2055
            Index           =   5
            Left            =   195
            Top             =   1350
            Width           =   3720
         End
      End
      Begin VB.Frame fra 
         Enabled         =   0   'False
         Height          =   2760
         Index           =   2
         Left            =   10350
         TabIndex        =   10
         Top             =   1635
         Width           =   4125
         Begin VB.TextBox TXT 
            Alignment       =   1  'Right Justify
            DataField       =   "Recibos"
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   8202
               SubFormatType   =   1
            EndProperty
            DataSource      =   "adoCon(1)"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   420
            Index           =   3
            Left            =   1530
            Locked          =   -1  'True
            TabIndex        =   18
            Text            =   "0"
            Top             =   390
            Width           =   2235
         End
         Begin VB.TextBox TXT 
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
            DataSource      =   "adoCon(1)"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   420
            Index           =   0
            Left            =   1530
            Locked          =   -1  'True
            TabIndex        =   13
            Text            =   " 0,00"
            ToolTipText     =   "Deuda del Propietario"
            Top             =   900
            Width           =   2235
         End
         Begin VB.TextBox TXT 
            Alignment       =   1  'Right Justify
            BeginProperty DataFormat 
               Type            =   1
               Format          =   """Bs"" #.##0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   8202
               SubFormatType   =   2
            EndProperty
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000002&
            Height          =   420
            Index           =   2
            Left            =   1530
            Locked          =   -1  'True
            TabIndex        =   12
            Text            =   " 0,00"
            Top             =   2010
            Width           =   2235
         End
         Begin VB.TextBox TXT 
            Alignment       =   1  'Right Justify
            DataField       =   "HonorariosMovimientoCaja"
            BeginProperty DataFormat 
               Type            =   1
               Format          =   """Bs"" #,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   8202
               SubFormatType   =   2
            EndProperty
            DataSource      =   "ADOcontrol(0)"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   420
            Index           =   1
            Left            =   1530
            Locked          =   -1  'True
            TabIndex        =   11
            Text            =   " 0,00"
            Top             =   1410
            Width           =   2235
         End
         Begin VB.Label lbl 
            Caption         =   "Recibos Pendientes:"
            Height          =   390
            Index           =   5
            Left            =   270
            TabIndex        =   17
            Top             =   420
            Width           =   1170
         End
         Begin VB.Label lbl 
            Caption         =   "Deuda:"
            Height          =   300
            Index           =   2
            Left            =   270
            TabIndex        =   16
            Top             =   945
            Width           =   1170
         End
         Begin VB.Label lbl 
            Caption         =   "Saldo:"
            Height          =   285
            Index           =   4
            Left            =   270
            TabIndex        =   15
            Top             =   2040
            Width           =   1170
         End
         Begin VB.Label lbl 
            Caption         =   "Honorarios:"
            Height          =   210
            Index           =   3
            Left            =   270
            TabIndex        =   14
            Top             =   1500
            Width           =   1170
         End
         Begin VB.Line Line1 
            BorderColor     =   &H8000000C&
            Index           =   0
            X1              =   1530
            X2              =   3780
            Y1              =   1905
            Y2              =   1905
         End
      End
      Begin VB.Frame fra 
         Caption         =   "Deuda"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   7530
         Index           =   1
         Left            =   255
         TabIndex        =   9
         Top             =   1650
         Width           =   9930
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid Flex 
            Height          =   6870
            Left            =   225
            TabIndex        =   6
            Tag             =   "1300|800|1500|1500|1500|1500"
            Top             =   420
            Width           =   9375
            _ExtentX        =   16536
            _ExtentY        =   12118
            _Version        =   393216
            Cols            =   6
            FixedCols       =   0
            BackColorFixed  =   -2147483646
            ForeColorFixed  =   -2147483639
            BackColorSel    =   65280
            ForeColorSel    =   0
            BackColorBkg    =   -2147483636
            GridColor       =   -2147483637
            SelectionMode   =   1
            FormatString    =   "^Factura |Período |Facturado |Abonado |Saldo |Acumulado"
            _NumberOfBands  =   1
            _Band(0).Cols   =   6
         End
      End
      Begin VB.Frame fra 
         Caption         =   "Datos del Propietario:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1080
         Index           =   0
         Left            =   255
         TabIndex        =   8
         Top             =   480
         Width           =   14280
         Begin MSDataListLib.DataCombo Dat 
            Bindings        =   "frmConvenio.frx":0038
            Height          =   315
            Index           =   2
            Left            =   7965
            TabIndex        =   4
            Top             =   450
            Width           =   1050
            _ExtentX        =   1852
            _ExtentY        =   556
            _Version        =   393216
            ListField       =   "Codigo"
            BoundColumn     =   "Nombre"
            Text            =   " "
         End
         Begin MSDataListLib.DataCombo Dat 
            DataField       =   "InmuebleMovimientoCaja"
            Height          =   315
            Index           =   0
            Left            =   1335
            TabIndex        =   1
            Top             =   450
            Width           =   1050
            _ExtentX        =   1852
            _ExtentY        =   556
            _Version        =   393216
            ListField       =   "CodInm"
            BoundColumn     =   "Nombre"
            Text            =   " "
            Object.DataMember      =   ""
         End
         Begin MSDataListLib.DataCombo Dat 
            Height          =   315
            Index           =   1
            Left            =   2400
            TabIndex        =   2
            Top             =   450
            Width           =   4050
            _ExtentX        =   7144
            _ExtentY        =   556
            _Version        =   393216
            ListField       =   "Nombre"
            BoundColumn     =   "CodInm"
            Text            =   " "
            Object.DataMember      =   ""
         End
         Begin MSDataListLib.DataCombo Dat 
            Bindings        =   "frmConvenio.frx":0050
            Height          =   315
            Index           =   3
            Left            =   9030
            TabIndex        =   5
            Top             =   450
            Width           =   4050
            _ExtentX        =   7144
            _ExtentY        =   556
            _Version        =   393216
            ListField       =   "Nombre"
            BoundColumn     =   "Codigo"
            Text            =   " "
         End
         Begin VB.Label lbl 
            Caption         =   "&Propietario"
            Height          =   285
            Index           =   1
            Left            =   6885
            TabIndex        =   3
            Top             =   465
            Width           =   915
         End
         Begin VB.Label lbl 
            Caption         =   "&Inmueble:"
            Height          =   285
            Index           =   0
            Left            =   270
            TabIndex        =   0
            Top             =   465
            Width           =   1215
         End
      End
      Begin MSAdodcLib.Adodc adoCon 
         Height          =   330
         Index           =   0
         Left            =   255
         Top             =   1695
         Visible         =   0   'False
         Width           =   2850
         _ExtentX        =   5027
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
      Begin MSAdodcLib.Adodc adoCon 
         Height          =   330
         Index           =   1
         Left            =   255
         Top             =   2070
         Visible         =   0   'False
         Width           =   2850
         _ExtentX        =   5027
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
         Caption         =   "Propietarios"
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
      Begin MSDataListLib.DataCombo Dat 
         Bindings        =   "frmConvenio.frx":0068
         Height          =   315
         Index           =   4
         Left            =   -71805
         TabIndex        =   61
         Top             =   2115
         Width           =   7530
         _ExtentX        =   13282
         _ExtentY        =   556
         _Version        =   393216
         MatchEntry      =   -1  'True
         Style           =   2
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         ListField       =   "Descripcion"
         BoundColumn     =   ""
         Text            =   " "
         Object.DataMember      =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDataListLib.DataCombo Dat 
         Bindings        =   "frmConvenio.frx":0081
         Height          =   315
         Index           =   5
         Left            =   -74325
         TabIndex        =   62
         Top             =   8745
         Width           =   4050
         _ExtentX        =   7144
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         ListField       =   "Nombre"
         BoundColumn     =   ""
         Text            =   " "
         Object.DataMember      =   ""
      End
      Begin VB.Label lbl 
         BackStyle       =   0  'Transparent
         Caption         =   "ESTATUS DE COBRANZA:"
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
         Height          =   285
         Index           =   31
         Left            =   -74310
         TabIndex        =   60
         Top             =   2175
         Width           =   2400
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H80000002&
         FillColor       =   &H80000002&
         FillStyle       =   0  'Solid
         Height          =   340
         Index           =   4
         Left            =   -74625
         Shape           =   4  'Rounded Rectangle
         Top             =   2100
         Width           =   10425
      End
      Begin VB.Label lbl 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         ForeColor       =   &H80000009&
         Height          =   285
         Index           =   30
         Left            =   -62190
         TabIndex        =   57
         Top             =   1725
         Width           =   1290
      End
      Begin VB.Label lbl 
         BackStyle       =   0  'Transparent
         Caption         =   "RECIBOS VENC.:"
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
         Height          =   285
         Index           =   29
         Left            =   -63735
         TabIndex        =   56
         Top             =   1725
         Width           =   1710
      End
      Begin VB.Line LIN 
         BorderColor     =   &H80000009&
         Index           =   3
         X1              =   -64215
         X2              =   -64215
         Y1              =   1680
         Y2              =   1980
      End
      Begin VB.Label lbl 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   28
         Left            =   -69645
         TabIndex        =   55
         Top             =   8790
         Width           =   3090
      End
      Begin VB.Label lbl 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   27
         Left            =   -73890
         TabIndex        =   54
         Top             =   8790
         Width           =   3345
      End
      Begin VB.Label lbl 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         ForeColor       =   &H80000009&
         Height          =   285
         Index           =   19
         Left            =   -62760
         TabIndex        =   37
         Top             =   1260
         Width           =   2025
      End
      Begin VB.Label lbl 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         ForeColor       =   &H80000009&
         Height          =   285
         Index           =   18
         Left            =   -66285
         TabIndex        =   36
         Top             =   1740
         Width           =   2025
      End
      Begin VB.Label lbl 
         BackStyle       =   0  'Transparent
         ForeColor       =   &H80000009&
         Height          =   285
         Index           =   17
         Left            =   -72855
         TabIndex        =   35
         Top             =   1725
         Width           =   4950
      End
      Begin VB.Label lbl 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         ForeColor       =   &H80000009&
         Height          =   285
         Index           =   16
         Left            =   -66300
         TabIndex        =   34
         Top             =   1245
         Width           =   2025
      End
      Begin VB.Label lbl 
         BackStyle       =   0  'Transparent
         ForeColor       =   &H80000009&
         Height          =   285
         Index           =   15
         Left            =   -72840
         TabIndex        =   33
         Top             =   1245
         Width           =   4950
      End
      Begin VB.Label lbl 
         Caption         =   "ELABORADO POR:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   14
         Left            =   -69675
         TabIndex        =   32
         Top             =   8490
         Width           =   3135
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         Caption         =   "DEPARTAMENTO LEGAL"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   165
         Index           =   13
         Left            =   -74265
         TabIndex        =   31
         Top             =   9120
         Width           =   4020
      End
      Begin VB.Label lbl 
         Caption         =   "AUTORIZADO POR:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   12
         Left            =   -74460
         TabIndex        =   30
         Top             =   8490
         Width           =   1785
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H80000002&
         Height          =   960
         Index           =   3
         Left            =   -69840
         Top             =   8445
         Width           =   3495
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H80000002&
         Height          =   960
         Index           =   2
         Left            =   -74655
         Top             =   8430
         Width           =   4650
      End
      Begin VB.Label lbl 
         BackStyle       =   0  'Transparent
         Caption         =   "FECHA:"
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
         Height          =   285
         Index           =   11
         Left            =   -63735
         TabIndex        =   28
         Top             =   1245
         Width           =   765
      End
      Begin VB.Line LIN 
         BorderColor     =   &H80000009&
         Index           =   2
         X1              =   -64215
         X2              =   -64215
         Y1              =   1200
         Y2              =   1500
      End
      Begin VB.Label lbl 
         BackStyle       =   0  'Transparent
         Caption         =   "CODIGO:"
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
         Height          =   285
         Index           =   10
         Left            =   -67560
         TabIndex        =   27
         Top             =   1725
         Width           =   1335
      End
      Begin VB.Label lbl 
         BackStyle       =   0  'Transparent
         Caption         =   "APTO. Nº:"
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
         Height          =   285
         Index           =   9
         Left            =   -67560
         TabIndex        =   26
         Top             =   1245
         Width           =   1335
      End
      Begin VB.Line LIN 
         BorderColor     =   &H80000009&
         Index           =   1
         X1              =   -67680
         X2              =   -67680
         Y1              =   1680
         Y2              =   1980
      End
      Begin VB.Line LIN 
         BorderColor     =   &H80000009&
         Index           =   0
         X1              =   -67695
         X2              =   -67695
         Y1              =   1200
         Y2              =   1500
      End
      Begin VB.Label lbl 
         BackStyle       =   0  'Transparent
         Caption         =   "RESIDENCIAS:"
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
         Height          =   285
         Index           =   8
         Left            =   -74325
         TabIndex        =   25
         Top             =   1725
         Width           =   1335
      End
      Begin VB.Label lbl 
         BackStyle       =   0  'Transparent
         Caption         =   "PROPIETARIO:"
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
         Height          =   285
         Index           =   7
         Left            =   -74325
         TabIndex        =   24
         Top             =   1245
         Width           =   1335
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H80000002&
         FillColor       =   &H80000002&
         FillStyle       =   0  'Solid
         Height          =   315
         Index           =   1
         Left            =   -74625
         Shape           =   4  'Rounded Rectangle
         Top             =   1665
         Width           =   14070
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H80000002&
         FillColor       =   &H80000002&
         FillStyle       =   0  'Solid
         Height          =   315
         Index           =   0
         Left            =   -74625
         Shape           =   4  'Rounded Rectangle
         Top             =   1200
         Width           =   14070
      End
      Begin VB.Label lbl 
         Caption         =   "RELACION DE DEUDA DEPARTAMENTO LEGAL"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   6
         Left            =   -74535
         TabIndex        =   23
         Top             =   765
         Width           =   5475
      End
   End
   Begin MSAdodcLib.Adodc adoCon 
      Height          =   330
      Index           =   3
      Left            =   2970
      Top             =   0
      Visible         =   0   'False
      Width           =   2850
      _ExtentX        =   5027
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
      Caption         =   "Abogados"
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
Attribute VB_Name = "frmConvenio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'varibales públicas a nivel de módulo
Dim InmTemp As String
Dim IntMMora As Integer
Dim IntHonoMorosidad As Integer
Dim strPropietario As String
'
Private Sub cmd_Click(Index%)
'variables locales
Dim errLocal As Long, rpReporte As ctlReport
Dim StrRutaInmueble
'
Select Case Index
'
    Case 0
        Unload Me
        Set frmConvenio = Nothing
    '
    Case 1
        '
        If SSTab1.Tab = 0 Then  'estado de cuenta
        
            'Call clear_Crystal(FrmAdmin.rptReporte)
            StrRutaInmueble = "\" & Dat(0) & "\"
            If Dir(gcPath & StrRutaInmueble & "inm.mdb") = "" Then
                MsgBox "Imposible emitir el reporte. No encuentra el archivo '" & gcPath & _
                StrRutaInmueble & "inm.mdb", vbExclamation, App.ProductName
                Exit Sub
            End If
            '
            Set rpReporte = New ctlReport
            With rpReporte
                '
                    Call Porcentajes
                    .Reporte = gcReport + "EdoCtaPro.rpt"
                    .OrigenDatos(0) = gcPath & StrRutaInmueble & "inm.mdb"
                    .OrigenDatos(1) = gcPath & StrRutaInmueble & "inm.mdb"
                    '.SortFields(0) = "+{Factura.periodo}"
                    .Formulas(0) = "Inmueble ='" & Dat(1) & "'"
                    .Formulas(1) = "MesMora =" & IntMMora
                    .Formulas(2) = "Morosidad =" & IntHonoMorosidad
                    .FormuladeSeleccion = "{Factura.Saldo} > 0 AND {Propietarios.Codigo} = '" & Dat(2) & "'"
                    .Salida = crPantalla
                    .TituloVentana = "Estado de Cuenta " & Dat(0) & " / " & Dat(2)
                    errLocal = .Imprimir
                    Call rtnBitacora("Impresión " & .TituloVentana)
                    If errLocal <> 0 Then MsgBox Err.Description, vbCritical, Err
                '
            End With
            Set rpReporte = Nothing
    
        Else    'relacion de deuda
            
            PopupMenu FrmAdmin.mnuPrint, vbPopupMenuRightAlign, cmd(1).Left, cmd(1).Top
                '
        End If
    
    Case 2
        Dim V As Boolean
    
        If lbl(18) = "" Then V = MsgBox("Falta el Código del Inmueble", vbCritical, App.ProductName)
        If lbl(16) = "" Then V = MsgBox("Falta el Nª del Apartamento", vbCritical, App.ProductName)
        If lbl(15) = "" Then V = MsgBox("Falta el nombre del propietario", vbCritical, App.ProductName)
    '
        For I = 0 To 11
            If InStr("0,9,10,11,4,5,6,8", I & ",") > 0 Then
                If txt(I) = "" Or Not IsNumeric(txt(I)) Then V = MsgBox("Falta o no és válido el " _
                & "valor del campo '" & txt(I).ToolTipText & "'", vbCritical, App.ProductName)
            End If
        Next
        
    If Dat(4) = "" Then
        V = MsgBox("Debe seleccionar un estatus de la cobranza", vbCritical, App.ProductName)
    ElseIf Dat(5) = "" Then
        V = MsgBox("Debe seleccionar a l abogado que autoriza el convenio", vbCritical, _
        App.ProductName)
    End If
    If V Then Exit Sub
    '
    If Respuesta("¿Desea guardar este convenio?") Then
    
        errLocal = ID
        'guarda el encabezado del convenio
        cnnConexion.Execute "INSERT INTO Convenio (IDconvenio,CodInm,CodProp,Propietario,Deuda," _
        & "DuedaCon,Gastos,CanceladoG,Honorarios,CanceladoH,DedGasto,DedHono,Inicial,NCuotas,ID" _
        & "Status,CobranzaEstatus,IDAbogado,Usuario,  Fecha) VALUES (" & errLocal & ",'" & _
        lbl(18) & "','" & lbl(16) & "','" & lbl(15) & "','" & CCur(txt(0)) & "','" & _
        CCur(txt(9)) & "','" & CCur(txt(10)) & "',False,'" & CCur(txt(11)) & "',False,'" & _
        CCur(txt(4)) & "','" & CCur(txt(5)) & "','" & CCur(txt(6)) & "','" & CCur(txt(8)) & _
        "',1," & Estatus & "," & Abogado & ",'" & gcUsuario & "',DATE())"
        '
        'agrega el detale del convenio
        With GRID1
            '.Row = 1
            I = 0
            Do
               I = I + 1
               .Row = I
                cnnConexion.Execute "INSERT INTO Convenio_Detalle (IDConvenio,Fecha,Monto,Cance" _
                & "lada) VALUES (" & errLocal & ",'" & .TextMatrix(.Row, 1) & "','" & _
                CCur(.TextMatrix(.Row, 2)) & "',False)"
                '.Row = .Row + 1
            Loop Until .Row = .Rows - 1
            
            
        End With
        'call rtnbitacora("Guardado convenio " & errlocal & " con éxito")
        'actualiza el campo convenio de la ficha del propietario
        cnnConexion.Execute "UPDATE Propietarios IN '" & gcPath & "\" & Me.Dat(0) & "\inm.mdb' SET" _
        & " Convenio=True WHERE Codigo ='" & Me.Dat(2) & "'"
        Call Guardar_Copia(Dat(0), errLocal)
        MsgBox "El convenio a sido guardado con éxito", vbInformation, App.ProductName
    End If
    
End Select
'
End Sub


Private Sub Dat_Click(Index%, Area%)
'variables locaels
Dim cnnCon As ADODB.Connection
'
If Area = 2 Then

    Call rtnLimpiar_Grid(FLEX)
    '
    Select Case Index
        
        Case 0
            Dat(1) = Dat(0).BoundText
            Call ConfigAdo
            
        Case 1
            Dat(0) = Dat(1).BoundText
            Call ConfigAdo
        
        Case 2
            Dat(3) = Dat(2).BoundText
            txt(1) = "0,00"
            txt(2) = "0,00"
            Set cnnCon = New ADODB.Connection
            cnnCon.Open adoCon(1).ConnectionString
            
            Call RtnFlex(Dat(2).Text, FLEX, 3, 30, 6, txt(1), cnnCon)
            cnnCon.Close
            Set cnnCon = Nothing
            
            With adoCon(1).Recordset
                .MoveFirst
                .Find "Codigo ='" & Dat(2) & "'"
                If Not .EOF And Not .BOF Then strPropietario = .Fields("Nombre"): Me.Tag = Format(.Fields("Cedula"), "#,##0")
                
            End With
        
        Case 3
        
            Dat(2) = Dat(3).BoundText
            txt(1) = "0,00"
            txt(2) = "0,00"
            Set cnnCon = New ADODB.Connection
            cnnCon.Open adoCon(1).ConnectionString
            
            Call RtnFlex(Dat(2).Text, FLEX, 3, 30, 6, txt(1), cnnCon)
            cnnCon.Close
            Set cnnCon = Nothing
            
            With adoCon(1).Recordset
                .MoveFirst
                .Find "Codigo ='" & Dat(2) & "'"
                If Not .EOF And Not .BOF Then strPropietario = .Fields("Nombre"): Me.Tag = Format(.Fields("Cedula"), "#,##0")
                
            End With
            
            
    End Select
    
    '
End If
'
End Sub

Private Sub Dat_GotFocus(Index As Integer)
'variables locales
'
If Index > 1 And Dir(gcPath & "\" & Dat(0) & "\inm.mdb") <> "" Then
    If InmTemp <> Dat(0) Then
        adoCon(1).Refresh
        InmTemp = Dat(0)
        adoCon(1).Recordset.MoveLast
    End If
End If
'
End Sub

Private Sub Dat_KeyPress(Index%, KeyAscii%): If KeyAscii = 13 Then Call Dat_Click(Index, 2)
End Sub

Private Sub Form_Load()
'


Call centra_titulo(FLEX, True)
adoCon(0).ConnectionString = cnnOLEDB & gcPath & "\sac.mdb"
adoCon(0).CommandType = adCmdTable
adoCon(0).RecordSource = "ParamConvenio"
adoCon(0).Refresh

With adoCon(2)

    .ConnectionString = cnnConexion.ConnectionString
    .CommandType = adCmdTable
    .CursorLocation = adUseClient
    .RecordSource = "Cobranza_Status"
    .Refresh
    .Recordset.Sort = "Descripcion"
    
End With

With adoCon(3)

    .ConnectionString = cnnConexion.ConnectionString
    .CommandType = adCmdTable
    .CursorLocation = adUseClient
    .RecordSource = "Convenio_Abogado"
    .Refresh
    .Recordset.Sort = "Nombre"
    
End With
'
lbl(27) = adoCon(0).Recordset("Abogado")
lbl(28) = gcNombreCompleto
'
Call centra_titulo(Grid, True)
Call centra_titulo(GRID1, True)
cmd(0).Picture = LoadResPicture("SALIR", vbResIcon)
cmd(1).Picture = LoadResPicture("PRINT", vbResIcon)
cmd(2).Picture = LoadResPicture("GUARDAR", vbResIcon)
'
Set Dat(0).RowSource = FrmAdmin.objRst
Set Dat(1).RowSource = FrmAdmin.ObjRstNom
'
End Sub

Private Sub ConfigAdo()
'variables locales
On Error GoTo salir
Dat(2) = ""
Dat(3) = ""
adoCon(1).ConnectionString = cnnOLEDB & gcPath & "\" & Dat(0) & "\inm.mdb"
adoCon(1).CommandType = adCmdTable
adoCon(1).RecordSource = "Propietarios"
adoCon(1).Refresh
adoCon(1).Recordset.MoveLast
For I = 0 To txt.UBound: txt(I) = 0
Next
'
Dat(2).SetFocus
salir:
If Err.Number = -2147467259 Then
    MsgBox "Datos del inmueble invalidos", vbCritical, App.ProductName
    Dat(0).SetFocus
ElseIf Err.Number <> 0 Then
    MsgBox Err.Description, vbCritical, App.ProductName
End If

'
End Sub

Private Sub Form_Resize()
'
On Error Resume Next
'
With SSTab1
    '
    .Height = Me.ScaleHeight - .Top - 200
    cmd(0).Top = SSTab1.Height - cmd(0).Height - 50
    cmd(1).Top = cmd(0).Top
    cmd(2).Top = cmd(0).Top
    fra(1).Height = SSTab1.Height - 200 - fra(1).Top
    FLEX.Height = fra(1).Height - 200 - FLEX.Top
    '
End With
'
End Sub

Public Sub SSTab1_Click(PreviousTab As Integer)
'variables locales
Dim curMonto@, curGastos@, curHonorarios@
Dim Dec As Long
'
Grid.Rows = 2
Call rtnLimpiar_Grid(Grid)

If SSTab1.Tab = 1 And Val(txt(3)) > 0 Then
'
    Grid.Rows = FLEX.Rows + 3
    '
    lbl(15) = Dat(3)
    lbl(16) = Dat(2)
    lbl(17) = Dat(1)
    lbl(18) = Dat(0)
    lbl(19) = Date
    lbl(30) = txt(3)
    '
    adoCon(0).Refresh
    Dec = FLEX.Rows - 1
    
    For I = 1 To Dec
    
        Grid.TextMatrix(I, 0) = I
        Grid.TextMatrix(I, 1) = FLEX.TextMatrix(I, 1)
        Grid.TextMatrix(I, 2) = FLEX.TextMatrix(I, 4)
        Grid.TextMatrix(I, 3) = IIf(chk(0).Value = vbChecked, Format((FLEX.TextMatrix(I, 4) * _
        adoCon(0).Recordset("GastosE") / 100) * (Dec - I), "#,##0.00"), "0,00")
        Grid.TextMatrix(I, 4) = IIf(chk(1).Value = vbChecked, Format(FLEX.TextMatrix(I, 4) * _
        adoCon(0).Recordset("Honorarios") / 100, "#,##0.00"), "0,00")
        Grid.TextMatrix(I, 5) = Format(CCur(Grid.TextMatrix(I, 2)) + CCur(Grid.TextMatrix(I, 3)) _
        + CCur(Grid.TextMatrix(I, 4)), "#,##0.00")
        'acumuladores
        curMonto = curMonto + CCur(Grid.TextMatrix(I, 2))
        curGastos = curGastos + CCur(Grid.TextMatrix(I, 3))
        curHonorarios = curHonorarios + CCur(Grid.TextMatrix(I, 4))
        
    Next
    'subtotales
    '
    Grid.RowHeight(I) = 315
    Grid.TextMatrix(I, 1) = "SUB-TOTAL"
    Grid.Row = I
    Grid.Col = 1
    Grid.CellAlignment = flexAlignRightCenter
    Grid.TextMatrix(I, 2) = Format(curMonto, "#,##0.00")
    Grid.TextMatrix(I, 3) = Format(curGastos, "#,##0.00")
    Grid.TextMatrix(I, 4) = Format(curHonorarios, "#,##0.00")
    Grid.TextMatrix(I, 5) = Format(curHonorarios + curGastos + curMonto, "#,##0.00")
    Grid.ColSel = Grid.Cols - 1
    Grid.FillStyle = flexFillRepeat
    'GRID.CellFontSize = 7
    Grid.CellFontBold = True
    Grid.CellTextStyle = flexTextRaised
    Grid.FillStyle = flexFillSingle
    '
    'deducciones
    I = I + 1
    Grid.TextMatrix(I, 1) = "DED."
    Grid.Row = I
    Grid.Col = 1
    Grid.CellAlignment = flexAlignRightCenter
    If Not IsNumeric(txt(4)) Then txt(4) = 0
    curGastos = curGastos - CCur(txt(4))
    curHonorarios = curHonorarios - CCur(txt(5))
    Grid.TextMatrix(I, 3) = Format(CCur(txt(4)), "#,##0.00")
    Grid.TextMatrix(I, 4) = Format(CCur(txt(5)), "#,##0.00")
    '
    'totales
    I = I + 1
    Grid.RowHeight(I) = 315
    Grid.TextMatrix(I, 1) = "TOTALES"
    Grid.Row = I
    Grid.Col = 1
    Grid.CellAlignment = flexAlignRightCenter
    Grid.TextMatrix(I, 2) = Format(curMonto, "#,##0.00")
    Grid.TextMatrix(I, 3) = Format(curGastos, "#,##0.00")
    Grid.TextMatrix(I, 4) = Format(curHonorarios, "#,##0.00")
    Grid.TextMatrix(I, 5) = Format(curHonorarios + curGastos + curMonto, "#,##0.00")
    Grid.ColSel = Grid.Cols - 1
    Grid.FillStyle = flexFillRepeat
    'GRID.CellFontSize = 7
    Grid.CellFontBold = True
    Grid.CellTextStyle = flexTextRaised
    Grid.FillStyle = flexFillSingle
    txt(9) = Grid.TextMatrix(I, 5)
    txt(10) = curGastos
    txt(11) = curHonorarios
    cmd(1).Enabled = True
    cmd(2).Visible = True
Else
    cmd(1).Enabled = False
    cmd(2).Visible = False
    '
End If
If SSTab1.Tab = 0 Then cmd(1).Enabled = True: cmd(2).Visible = False
'
End Sub

Private Sub SSTab1_DblClick()
'variables locales
Call rtnLimpiar_Grid(Grid)
'
If SSTab1.Tab = 1 And Val(txt(3)) > 0 Then
'
    For I = 1 To FLEX.Rows - 1
        Grid.TextMatrix(I, 0) = I
        Grid.TextMatrix(I, 1) = FLEX.TextMatrix(I, 1)
        Grid.TextMatrix(I, 2) = FLEX.TextMatrix(I, 2)
    Next
    '
End If
'
End Sub

Private Sub txt_Change(Index As Integer)
'VARIABLES LOCALES
Dim m@
'
If Index = 0 And txt(0) <> "" Then
    txt(2) = Format(CCur(txt(0)) + CCur(txt(1)), "#,##0.00")
ElseIf Index = 9 Then   'deuda
    If txt(6) > 0 Then GoTo 5
ElseIf Index = 6 Then 'INICIAL
5   If txt(6) = "" Then txt(6) = 0
    txt(7) = Format(CCur(txt(9)) - CCur(txt(6)), "#,##0.00")
    If txt(8) > 0 Then GoTo 10

ElseIf Index = 8 And CCur(txt(7)) > 0 Then 'CUOTAS
10  GRID1.Rows = 2
    Call rtnLimpiar_Grid(GRID1)
    If IsNumeric(txt(8)) Then
        m = txt(7) / txt(8)
        
        With GRID1
            .Rows = txt(8) + 1
            For I = 1 To txt(8)
                .TextMatrix(I, 0) = I
                .TextMatrix(I, 1) = DateAdd("M", I, Date)
                .TextMatrix(I, 2) = Format(m, "#,##0.00")
            Next
            '
        End With
        '
    End If
    '
End If
'
End Sub

Private Sub Porcentajes()
'variable locales
Dim rst As New ADODB.Recordset

rst.Open "Inmueble", cnnConexion, adOpenKeyset, adLockOptimistic, adCmdTable
rst.Filter = "CodInm='" & Dat(0) & "'"
If Not rst.EOF And Not rst.BOF Then
    IntMMora = rst("MesesMora")
    IntHonoMorosidad = rst("HonoMorosidad")
End If
rst.Close
Set rst = Nothing
End Sub

Private Sub txt_GotFocus(Index%)
txt(Index) = CCur(txt(Index))
txt(Index).SelStart = 0
txt(Index).SelLength = Len(txt(Index))
End Sub

Private Sub txt_KeyPress(Index%, KeyAscii%)
If KeyAscii = Asc(".") Then KeyAscii = Asc(",")
Call Validacion(KeyAscii, "0123456789,")
End Sub

Private Sub txt_LostFocus(Index%)
txt(Index) = Format(txt(Index), IIf(Index = 8, "#,##0", "#,##0.00"))
End Sub

Private Function ID() As Long
'variables locales
Dim rstlocal As New ADODB.Recordset
'
rstlocal.Open "SELECT Max(IDConvenio) FROM Convenio", cnnConexion, adOpenKeyset, _
adLockOptimistic, adCmdText
ID = IIf(IsNull(rstlocal.Fields(0)), 0, rstlocal(0))
ID = ID + 1
'
rstlocal.Close
Set rsrlocal = Nothing
'
End Function

Function Estatus() As Integer
'variables locales
Dim rstlocal As New ADODB.Recordset
'
rstlocal.Open "SELECT * FROM cobranza_status WHERE Descripcion ='" & Dat(4) & "'", cnnConexion, _
adOpenKeyset, adLockOptimistic, adCmdText
'
If Not rstlocal.EOF And Not rstlocal.BOF Then Estatus = rstlocal("IDEStatus")
'
rstlocal.Close
Set rstlocal = Nothing
'
End Function

Function Abogado() As Integer
'variables locales
Dim rstlocal As New ADODB.Recordset
'
rstlocal.Open "SELECT * FROM Convenio_abogado WHERE Nombre='" & Dat(5) & "'", cnnConexion, _
adOpenKeyset, adLockOptimistic, adCmdText

If Not rstlocal.EOF And Not rstlocal.BOF Then Abogado = rstlocal("IDabogado")
'
rstlocal.Close
Set rstlocal = Nothing
End Function

Private Sub Guardar_Copia(Inm As String, Conve As Long)
'imprime el convenio de pago
Dim cAletra As clsNum2Let
Dim rpReporte As ctlReport
Dim Spacio As Integer, INI As Integer
Dim Cantidad As String
'
Set cAletra = New clsNum2Let
cAletra.Moneda = "Bs."
'
Set rpReporte = New ctlReport
With rpReporte
    '
    .Reporte = gcReport & "convenio.rpt"
    .FormuladeSeleccion = "{Convenio.IDConvenio}=" & Conve
    cAletra.Numero = CCur(frmConvenio.txt(9))
    Cantidad = UCase(cAletra.ALetra) & "(" & frmConvenio.txt(9) & ")"
    '
    Do
        INI = Spacio + 1
        Spacio = InStr(INI, Cantidad, " ", vbTextCompare)
    Loop Until Spacio > 15 Or INI = Spacio
    
    Spacio = 0
    .Formulas(0) = "deuda='" & String(15 - INI, "x") & Left(Cantidad, INI - 1) & "'"
    .Formulas(1) = "deuda1='" & Mid(Cantidad, INI, Len(Cantidad)) & String(10, "x") & "'"
    '
    .Formulas(2) = "empresa='" & sysEmpresa & "'"
    
    cAletra.Numero = CCur(frmConvenio.txt(6))
    Cantidad = UCase(cAletra.ALetra) & "(" & frmConvenio.txt(6) & ")"
    '
    Do
        INI = Spacio + 1
        Spacio = InStr(INI, Cantidad, " ", vbTextCompare)
    Loop Until Spacio > 28 Or INI > Spacio
    
    Spacio = 0
    .Formulas(3) = "inicial='" & String(28 - INI, "x") & Left(Cantidad, INI - 1) & "'"
    .Formulas(4) = "inicial1='" & Mid(Cantidad, INI, Len(Cantidad)) & String(10, "x") & "'"
    cAletra.Moneda = ""
    cAletra.Numero = CLng(frmConvenio.txt(8))
    Cantidad = cAletra.ALetra
    .Formulas(5) = "cuotas1='" & Left(UCase(Cantidad), InStr(Cantidad, " ")) & "(" & frmConvenio.txt(8) & ")'"
    '
    cAletra.Numero = CCur(CCur(frmConvenio.txt(7)) / CCur(frmConvenio.txt(8)))
    cAletra.Moneda = "Bs."
    Cantidad = UCase(cAletra.ALetra) & "(" & Format(cAletra.Numero, "#,##0.00") & ")"
    '
    Do
        INI = Spacio + 1
        Spacio = InStr(INI, Cantidad, " ", vbTextCompare)
    Loop Until Spacio > 70 Or INI > Spacio
    '
    Spacio = 0
    .Formulas(6) = "bscuota='" & String(70 - INI, "x") & Left(Cantidad, INI - 1) & "'"
    .Formulas(7) = "bscuota1='" & Mid(Cantidad, INI, Len(Cantidad)) & String(10, "x") & "'"
    '
    cAletra.Numero = CCur(frmConvenio.txt(7))
    Cantidad = UCase(cAletra.ALetra) & "(" & frmConvenio.txt(7) & ")"
    Do
        INI = Spacio + 1
        Spacio = InStr(INI, Cantidad, " ", vbTextCompare)
        
    Loop Until Spacio > 60 Or INI > Spacio
    .Formulas(8) = "finan='" & String(70 - INI, "x") & Left(Cantidad, INI - 1) & "'"
    .Formulas(9) = "finan1='" & Mid(Cantidad, INI, Len(Cantidad)) & String(10, "x") & "'"
    .Formulas(10) = "CI='" & IIf(IsNull(frmConvenio.Tag), "", frmConvenio.Tag) & "'"
    .Salida = crArchivoDisco
    .ArchivoSalida = gcPath & "\" & Inm & "\reportes\CON" & Conve & ".rpt"
    .Imprimir
    
    
End With
Set rpReporte = Nothing
End Sub
