VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form FrmOperativo 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Parámetros Operativos"
   ClientHeight    =   5790
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5670
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5790
   ScaleWidth      =   5670
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdBackUp 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   405
      Index           =   1
      Left            =   2625
      TabIndex        =   22
      Top             =   5040
      Width           =   1170
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   5775
      Left            =   15
      TabIndex        =   0
      Top             =   15
      Width           =   5655
      _ExtentX        =   9975
      _ExtentY        =   10186
      _Version        =   393216
      Tabs            =   6
      Tab             =   3
      TabHeight       =   847
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
      TabCaption(0)   =   "Back-Up"
      TabPicture(0)   =   "FrmOperativo.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "lblBackUp(0)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "fraOperativo(0)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "fraOperativo(1)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "cmdBackUp(0)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "cmdBackUp(2)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "chkBackup"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).ControlCount=   6
      TabCaption(1)   =   "Facturación - (IDB)"
      TabPicture(1)   =   "FrmOperativo.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "lblBackUp(5)"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "lblBackUp(6)"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Text1(0)"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "List1(0)"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "cmdBackUp(3)"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "cmdBackUp(4)"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).ControlCount=   6
      TabCaption(2)   =   "Facturación Presupuestados"
      TabPicture(2)   =   "FrmOperativo.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "lblBackUp(7)"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "lblBackUp(8)"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "Text1(1)"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "List1(1)"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).Control(4)=   "cmdBackUp(5)"
      Tab(2).Control(4).Enabled=   0   'False
      Tab(2).Control(5)=   "cmdBackUp(6)"
      Tab(2).Control(5).Enabled=   0   'False
      Tab(2).ControlCount=   6
      TabCaption(3)   =   "Servicios Básicos"
      TabPicture(3)   =   "FrmOperativo.frx":0054
      Tab(3).ControlEnabled=   -1  'True
      Tab(3).Control(0)=   "lblBackUp(9)"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).Control(1)=   "DataGrid1"
      Tab(3).Control(1).Enabled=   0   'False
      Tab(3).Control(2)=   "cmdBackUp(7)"
      Tab(3).Control(2).Enabled=   0   'False
      Tab(3).Control(3)=   "cmdBackUp(8)"
      Tab(3).Control(3).Enabled=   0   'False
      Tab(3).ControlCount=   4
      TabCaption(4)   =   "Nómina"
      TabPicture(4)   =   "FrmOperativo.frx":0070
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "fraOperativo(2)"
      Tab(4).Control(0).Enabled=   0   'False
      Tab(4).Control(1)=   "cmdBackUp(9)"
      Tab(4).Control(1).Enabled=   0   'False
      Tab(4).Control(2)=   "cmdBackUp(10)"
      Tab(4).Control(2).Enabled=   0   'False
      Tab(4).ControlCount=   3
      TabCaption(5)   =   "Convenios"
      TabPicture(5)   =   "FrmOperativo.frx":008C
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "AdoAbogado"
      Tab(5).Control(0).Enabled=   0   'False
      Tab(5).Control(1)=   "fraOperativo(3)"
      Tab(5).Control(1).Enabled=   0   'False
      Tab(5).Control(2)=   "cmdBackUp(11)"
      Tab(5).Control(2).Enabled=   0   'False
      Tab(5).Control(3)=   "cmdBackUp(12)"
      Tab(5).Control(3).Enabled=   0   'False
      Tab(5).ControlCount=   4
      Begin VB.CommandButton cmdBackUp 
         Caption         =   "A&plicar"
         Height          =   405
         Index           =   12
         Left            =   -71055
         TabIndex        =   46
         Top             =   5025
         Width           =   1170
      End
      Begin VB.CommandButton cmdBackUp 
         Caption         =   "&Aceptar"
         Height          =   405
         Index           =   11
         Left            =   -73710
         TabIndex        =   45
         Top             =   5025
         Width           =   1170
      End
      Begin VB.Frame fraOperativo 
         Caption         =   "Paràmetros convenimiento de pago"
         Height          =   3285
         Index           =   3
         Left            =   -74775
         TabIndex        =   40
         Top             =   1290
         Width           =   5010
         Begin VB.TextBox txt 
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
            DataSource      =   "adoCon"
            Height          =   315
            Index           =   1
            Left            =   2475
            TabIndex        =   44
            Text            =   "0"
            Top             =   1290
            Width           =   975
         End
         Begin VB.TextBox txt 
            Alignment       =   1  'Right Justify
            DataField       =   "GastosE"
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "#,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   8202
               SubFormatType   =   0
            EndProperty
            DataSource      =   "adoCon"
            Height          =   315
            Index           =   0
            Left            =   2475
            TabIndex        =   43
            Text            =   "0"
            Top             =   690
            Width           =   975
         End
         Begin MSAdodcLib.Adodc adoCon 
            Height          =   420
            Left            =   180
            Top             =   2670
            Visible         =   0   'False
            Width           =   4665
            _ExtentX        =   8229
            _ExtentY        =   741
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
            Caption         =   "% Convenios de Pago"
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
         Begin MSDataListLib.DataCombo data 
            Bindings        =   "FrmOperativo.frx":00A8
            DataField       =   "Abogado"
            DataSource      =   "adoCon"
            Height          =   315
            Left            =   2475
            TabIndex        =   48
            Top             =   1860
            Width           =   2385
            _ExtentX        =   4207
            _ExtentY        =   556
            _Version        =   393216
            ListField       =   "Nombre"
            Text            =   ""
         End
         Begin VB.Label lblBackUp 
            Caption         =   "Abogado (predeterminado)"
            Height          =   285
            Index           =   14
            Left            =   300
            TabIndex        =   47
            Top             =   1935
            Width           =   2100
         End
         Begin VB.Label lblBackUp 
            Caption         =   "% Honorarios Extrajudiciales"
            Height          =   285
            Index           =   13
            Left            =   300
            TabIndex        =   42
            Top             =   1320
            Width           =   2100
         End
         Begin VB.Label lblBackUp 
            Caption         =   "% Gastos de Cobranza"
            Height          =   285
            Index           =   12
            Left            =   300
            TabIndex        =   41
            Top             =   750
            Width           =   1770
         End
      End
      Begin VB.CommandButton cmdBackUp 
         Caption         =   "&Aceptar"
         Height          =   405
         Index           =   10
         Left            =   -73710
         TabIndex        =   39
         Top             =   5025
         Width           =   1170
      End
      Begin VB.CommandButton cmdBackUp 
         Caption         =   "A&plicar"
         Height          =   405
         Index           =   9
         Left            =   -71055
         TabIndex        =   38
         Top             =   5025
         Width           =   1170
      End
      Begin VB.Frame fraOperativo 
         Caption         =   "Horario de Trabajo:"
         Height          =   1350
         Index           =   2
         Left            =   -74775
         TabIndex        =   33
         Top             =   1320
         Width           =   5055
         Begin MSMask.MaskEdBox mskBackUp 
            DataField       =   "HEntrada"
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "hh:mm ampm"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   8202
               SubFormatType   =   0
            EndProperty
            Height          =   315
            Index           =   1
            Left            =   1125
            TabIndex        =   36
            ToolTipText     =   "Hora formato AM PM"
            Top             =   525
            Width           =   945
            _ExtentX        =   1667
            _ExtentY        =   556
            _Version        =   393216
            PromptInclude   =   0   'False
            AllowPrompt     =   -1  'True
            MaxLength       =   8
            Format          =   "hh:mm ampm"
            Mask            =   "##:## >??"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox mskBackUp 
            DataField       =   "Hsalida"
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "hh:mm ampm"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   8202
               SubFormatType   =   0
            EndProperty
            Height          =   315
            Index           =   2
            Left            =   3585
            TabIndex        =   37
            ToolTipText     =   "Hora formato AM PM"
            Top             =   525
            Width           =   960
            _ExtentX        =   1693
            _ExtentY        =   556
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   8
            Format          =   "hh:mm ampm"
            Mask            =   "##:## >??"
            PromptChar      =   "_"
         End
         Begin VB.Label lblBackUp 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Hasta:"
            Height          =   195
            Index           =   11
            Left            =   2985
            TabIndex        =   35
            Top             =   585
            Width           =   465
         End
         Begin VB.Label lblBackUp 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Desde:"
            Height          =   195
            Index           =   10
            Left            =   420
            TabIndex        =   34
            Top             =   585
            Width           =   510
         End
      End
      Begin VB.CommandButton cmdBackUp 
         Caption         =   "&Eliminar"
         Height          =   405
         Index           =   8
         Left            =   3945
         TabIndex        =   32
         Top             =   5025
         Width           =   1170
      End
      Begin VB.CommandButton cmdBackUp 
         Caption         =   "&Agregar"
         Height          =   405
         Index           =   7
         Left            =   1290
         TabIndex        =   31
         Top             =   5025
         Width           =   1170
      End
      Begin MSDataGridLib.DataGrid DataGrid1 
         Height          =   2280
         Left            =   345
         TabIndex        =   30
         Top             =   2280
         Width           =   4800
         _ExtentX        =   8467
         _ExtentY        =   4022
         _Version        =   393216
         AllowUpdate     =   -1  'True
         HeadLines       =   1
         RowHeight       =   17
         FormatLocked    =   -1  'True
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Comic Sans MS"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Proveedores de Servicios Básicos"
         ColumnCount     =   2
         BeginProperty Column00 
            DataField       =   "Descripcion"
            Caption         =   "Servicio"
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
            DataField       =   "IDPRoveedor"
            Caption         =   "Cod.Proveedor"
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
            EndProperty
            BeginProperty Column01 
            EndProperty
         EndProperty
      End
      Begin VB.CommandButton cmdBackUp 
         Caption         =   "&Eliminar"
         Height          =   405
         Index           =   6
         Left            =   -71055
         TabIndex        =   28
         Top             =   5025
         Width           =   1170
      End
      Begin VB.CommandButton cmdBackUp 
         Caption         =   "&Aceptar"
         Height          =   405
         Index           =   5
         Left            =   -73710
         TabIndex        =   27
         Top             =   5025
         Width           =   1170
      End
      Begin VB.ListBox List1 
         Columns         =   6
         DataField       =   "CodGasto"
         Height          =   2205
         Index           =   1
         ItemData        =   "FrmOperativo.frx":00C1
         Left            =   -74730
         List            =   "FrmOperativo.frx":00C3
         Sorted          =   -1  'True
         TabIndex        =   26
         Top             =   2600
         Width           =   5055
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   1
         Left            =   -71970
         TabIndex        =   24
         Top             =   1980
         Width           =   1050
      End
      Begin VB.CommandButton cmdBackUp 
         Caption         =   "&Eliminar"
         Height          =   405
         Index           =   4
         Left            =   -71055
         TabIndex        =   21
         Top             =   5025
         Width           =   1170
      End
      Begin VB.CommandButton cmdBackUp 
         Caption         =   "&Aceptar"
         Height          =   405
         Index           =   3
         Left            =   -73710
         TabIndex        =   20
         Top             =   5025
         Width           =   1170
      End
      Begin VB.ListBox List1 
         Columns         =   6
         DataField       =   "CodGasto"
         Height          =   2205
         Index           =   0
         ItemData        =   "FrmOperativo.frx":00C5
         Left            =   -74730
         List            =   "FrmOperativo.frx":00C7
         Sorted          =   -1  'True
         TabIndex        =   19
         Top             =   2600
         Width           =   5055
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   0
         Left            =   -71970
         TabIndex        =   17
         Top             =   1980
         Width           =   1050
      End
      Begin VB.CheckBox chkBackup 
         Height          =   315
         Left            =   -74520
         TabIndex        =   15
         Top             =   1305
         Width           =   390
      End
      Begin VB.CommandButton cmdBackUp 
         Caption         =   "A&plicar"
         Height          =   405
         Index           =   2
         Left            =   -71055
         TabIndex        =   13
         Top             =   5025
         Width           =   1170
      End
      Begin VB.CommandButton cmdBackUp 
         Caption         =   "&Aceptar"
         Height          =   405
         Index           =   0
         Left            =   -73710
         TabIndex        =   12
         Top             =   5025
         Width           =   1170
      End
      Begin VB.Frame fraOperativo 
         Caption         =   "Hacer Copia:"
         Enabled         =   0   'False
         Height          =   1860
         Index           =   1
         Left            =   -74595
         TabIndex        =   4
         Top             =   2970
         Width           =   4710
         Begin VB.TextBox txtBackUp 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            Height          =   315
            Left            =   3375
            TabIndex        =   11
            Text            =   "0"
            Top             =   1335
            Width           =   975
         End
         Begin VB.ComboBox cmbBackUp 
            Enabled         =   0   'False
            Height          =   315
            ItemData        =   "FrmOperativo.frx":00C9
            Left            =   3015
            List            =   "FrmOperativo.frx":00E2
            Style           =   2  'Dropdown List
            TabIndex        =   8
            Top             =   855
            Visible         =   0   'False
            Width           =   1365
         End
         Begin MSMask.MaskEdBox mskBackUp 
            Height          =   315
            Index           =   0
            Left            =   750
            TabIndex        =   7
            ToolTipText     =   "Hora formato AM PM"
            Top             =   870
            Width           =   885
            _ExtentX        =   1561
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
            MaxLength       =   8
            Format          =   "hh:mm ampm"
            Mask            =   "##:## >??"
            PromptChar      =   "_"
         End
         Begin VB.Label lblBackUp 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Establezca una hora durante la cual el sistema no este en uso."
            Enabled         =   0   'False
            Height          =   375
            Index           =   3
            Left            =   0
            TabIndex        =   14
            Top             =   360
            Width           =   4680
         End
         Begin VB.Label lblBackUp 
            BackStyle       =   0  'Transparent
            Caption         =   "Almacenar Máximo de Copias: (Máx.30)"
            Enabled         =   0   'False
            Height          =   315
            Index           =   4
            Left            =   240
            TabIndex        =   10
            Top             =   1425
            Width           =   3015
         End
         Begin VB.Label lblBackUp 
            Caption         =   "de cada dia:"
            Enabled         =   0   'False
            Height          =   285
            Index           =   2
            Left            =   2025
            TabIndex        =   9
            Top             =   885
            Visible         =   0   'False
            Width           =   1170
         End
         Begin VB.Label lblBackUp 
            Caption         =   "a las:"
            Enabled         =   0   'False
            Height          =   285
            Index           =   1
            Left            =   285
            TabIndex        =   6
            Top             =   885
            Width           =   540
         End
      End
      Begin VB.Frame fraOperativo 
         Caption         =   "Hacer Copia:"
         Enabled         =   0   'False
         Height          =   990
         Index           =   0
         Left            =   -74595
         TabIndex        =   2
         Top             =   1785
         Width           =   4710
         Begin VB.OptionButton optBackUp 
            Caption         =   "Semanal"
            Enabled         =   0   'False
            Height          =   315
            Index           =   1
            Left            =   2610
            TabIndex        =   5
            Top             =   390
            Width           =   1215
         End
         Begin VB.OptionButton optBackUp 
            Caption         =   "Diaria"
            Enabled         =   0   'False
            Height          =   315
            Index           =   0
            Left            =   915
            TabIndex        =   3
            Top             =   390
            Value           =   -1  'True
            Width           =   1215
         End
      End
      Begin MSAdodcLib.Adodc AdoAbogado 
         Height          =   420
         Left            =   -74595
         Top             =   4575
         Visible         =   0   'False
         Width           =   4665
         _ExtentX        =   8229
         _ExtentY        =   741
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
      Begin VB.Label lblBackUp 
         BackStyle       =   0  'Transparent
         Caption         =   "Información referente a los proveedores de servicios báscios (Luz / Agua / Teléfono)"
         Height          =   510
         Index           =   9
         Left            =   105
         TabIndex        =   29
         Top             =   1395
         Width           =   5235
      End
      Begin VB.Label lblBackUp 
         BackStyle       =   0  'Transparent
         Caption         =   "Código del Gasto:"
         Height          =   300
         Index           =   8
         Left            =   -73395
         TabIndex        =   25
         Top             =   2040
         Width           =   1425
      End
      Begin VB.Label lblBackUp 
         BackStyle       =   0  'Transparent
         Caption         =   $"FrmOperativo.frx":0122
         Height          =   630
         Index           =   7
         Left            =   -74850
         TabIndex        =   23
         Top             =   1290
         Width           =   5190
      End
      Begin VB.Label lblBackUp 
         BackStyle       =   0  'Transparent
         Caption         =   "Código del Gasto:"
         Height          =   300
         Index           =   6
         Left            =   -73395
         TabIndex        =   18
         Top             =   2025
         Width           =   1425
      End
      Begin VB.Label lblBackUp 
         BackStyle       =   0  'Transparent
         Caption         =   "Ingrese los Códigos  de los gastos que no son tomados en cuenta para el cálculo del Impuesto al Débito Bancario:"
         Height          =   510
         Index           =   5
         Left            =   -74850
         TabIndex        =   16
         Top             =   1290
         Width           =   5235
      End
      Begin VB.Label lblBackUp 
         BackStyle       =   0  'Transparent
         Caption         =   "Programe la copia de seguridad de la base de datos del sistema:"
         Height          =   390
         Index           =   0
         Left            =   -74145
         TabIndex        =   1
         Top             =   1260
         Width           =   4215
      End
   End
End
Attribute VB_Name = "FrmOperativo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
    Dim rstBackUp As New ADODB.Recordset
    Dim rstFacIDB As New ADODB.Recordset
    Dim rstFacPre As New ADODB.Recordset
    Dim rstSerBa As New ADODB.Recordset
    Dim rstAmbiente As New ADODB.Recordset
    Dim cnnBackUp As New ADODB.Connection
    
    Private Sub chkBackup_Click()
    
    If chkBackup.Value = 1 Then
        Call activar(True)
    Else
        Call activar(False)
    End If
    
    End Sub

    Private Sub cmdBackUp_Click(Index%)
    'variables locales
    Dim F As Form
    '
    Select Case Index
        '
        Case 0 'Aceptar
            Call Guardar
            Unload Me
            
        Case 2: Call Guardar
        
        Case 3, 1, 5, 7: Unload Me
        
        Case 4: Call Del(rstFacIDB, 0, "Fact_IDB")
        
        Case 6: Call Del(rstFacPre, 1, "Fact_Presupuestados")
        
        Case 10, 9
            mskBackUp(1).PromptInclude = True
            mskBackUp(2).PromptInclude = True
            rstAmbiente("HEntrada") = mskBackUp(1)
            rstAmbiente("HSalida") = mskBackUp(2)
            rstAmbiente.Update
            mskBackUp(1).PromptInclude = False
            mskBackUp(2).PromptInclude = False
            If Index = 10 Then Unload Me
        
        Case 11, 12
        
            If IsNumeric(Txt(0)) And IsNumeric(Txt(1)) Then
                adoCon.Recordset("Fecha") = Date
                adoCon.Recordset("Usuario") = gcUsuario
                adoCon.Recordset.Update
                
                For Each F In Forms
                    If F.Name = "frmConvenio" Then
                        F.adoCon(0).Recordset.Requery
                        If F.SSTab1.tab = 1 Then Call F.SSTab1_Click(0)
                    End If
                    
                Next
                
            Else
                MsgBox "Introdujo valores no válidos", vbExclamation, App.ProductName
            End If
            If Index = 11 Then Unload Me
        '
    End Select
    '
    End Sub

    Private Sub data_KeyPress(KeyAscii As Integer)
    If keyasii = 13 Then SendKeys vbTab
    KeyAscii = 0
    End Sub

    Private Sub DataGrid1_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If DataGrid1.Col = 1 Then Call Validacion(KeyAscii, "0123456789")
    End Sub

    Private Sub Form_Load()
    '
    Call CenterForm(FrmOperativo)
    cnnBackUp.Open cnnOLEDB & gcPath & "\tablas.mdb"
    rstBackUp.Open "SELECT * FROM ConfigBackUp;", cnnBackUp, adOpenKeyset, adLockOptimistic, _
    adCmdText
    '
    With rstBackUp
        '
        If .RecordCount > 0 Then
            If !Diaria Then
                optBackUp(0).Value = True
            Else
                optBackUp(1).Value = True
                lblBackUp(2).Visible = True
                cmbBackUp.Visible = True
            End If
            mskBackUp(0).PromptInclude = False
            mskBackUp(0) = Format(!Hora, "hhmm AMPM")
            mskBackUp(0).PromptInclude = True
            If !Dia > 0 Then cmbBackUp = Format(Weekday(!Dia, vbUseSystemDayOfWeek), "DDDD")
            txtBackUp = !Copias
            chkBackup.Value = 1
        Else
            chkBackup.Value = 0
        End If
        '
    End With
    '
    Call ConfigList(rstFacIDB, 0, "Fact_IDB")
    Call ConfigList(rstFacPre, 1, "Fact_Presupuestados")
    rstSerBa.Open "Serviciostipo", cnnConexion, adOpenKeyset, adLockOptimistic, adCmdTable
    Set DataGrid1.DataSource = rstSerBa
    rstAmbiente.Open "Ambiente", cnnConexion, adOpenKeyset, adLockOptimistic, adCmdTable
    Set mskBackUp(1).DataSource = rstAmbiente
    Set mskBackUp(2).DataSource = rstAmbiente
    '
    With adoCon
        .CursorLocation = adUseClient
        .CursorType = adOpenKeyset
        .ConnectionString = cnnOLEDB + gcPath + "\sac.mdb"
        .RecordSource = "ParamConvenio"
        .CommandType = adCmdTable
        .Refresh
    End With
    '
    With AdoAbogado
    
        .CursorLocation = adUseClient
        .CursorType = adOpenKeyset
        .ConnectionString = cnnOLEDB + gcPath + "\sac.mdb"
        .RecordSource = "Convenio_Abogado"
        .CommandType = adCmdTable
        .Refresh
        
    End With
    '
    End Sub

    Private Sub Form_Unload(Cancel As Integer)
    'variables locales
    rstBackUp.Close
    rstFacIDB.Close
    rstFacPre.Close
    rstSerBa.Close
    rstAmbiente.Close
    cnnBackUp.Close
    '-----------------------------
    Set rstBackUp = Nothing
    Set rstFacIDB = Nothing
    Set rstFacPre = Nothing
    Set rstSerBa = Nothing
    Set rstAmbiente = Nothing
    Set cnnBackUp = Nothing
    Set FrmOperativo = Nothing
    '-----------------------------
    End Sub

    Private Sub mskBackUp_GotFocus(Index As Integer)
    If Index = 0 Then
        With mskBackUp(0)
            .SelStart = 0
            .SelLength = Len(.Text)
        End With
    End If
    End Sub

    Private Sub mskBackUp_KeyPress(Index%, KeyAscii%)
    Call Validacion(KeyAscii, "1234567890AMamPp")
    If Index = 0 Then
        If KeyAscii = 13 Then txtBackUp.SetFocus
    End If
    End Sub

    Private Sub optBackUp_Click(Index As Integer)
    cmbBackUp.Visible = Not cmbBackUp.Visible
    lblBackUp(2).Visible = Not lblBackUp(2).Visible
    End Sub


    Private Sub Guardar()
    
    If chkBackup.Value = 1 Then
        
        If Left(mskBackUp(0), 2) > 12 Or Mid(mskBackUp(0), 4, 2) > 59 Or (Right(mskBackUp(0), 2) <> "AM" _
        And Right(mskBackUp(0), 2) <> "PM") Then
            MsgBox "Error en el formato de la hora introducida", vbExclamation, "Hora errada"
            Exit Sub
        End If
        If txtBackUp = "" Then
            MsgBox "Falta Cantidad de Respaldos Máximos", vbExclamation, App.ProductName
            txtBackUp.SetFocus
            Exit Sub
        End If
        If txtBackUp > 30 Or txtBackUp = 0 Then
            MsgBox "Verifique el número de copias..", vbExclamation, "Desbordamiento"
            Exit Sub
        End If
        cnnBackUp.Execute "DELETE  FROM ConfigBackUp"
        Dim StrEstado As String
        If optBackUp(0).Value Then
            StrEstado = "True"
        Else
            StrEstado = "False"
        End If
        cnnBackUp.Execute "INSERT INTO ConfigBackUp(Diaria,Hora,Dia,Copias,Usuario,Fecha,HoraR)" _
        & " VALUES(" & StrEstado & ",'" & mskBackUp(0) & "'," & cmbBackUp.ListIndex + 1 & "," _
        & txtBackUp & ",'" & gcUsuario & "',Date(),Time());"
    Else
        rstBackUp.Delete
    End If
    
    End Sub



    Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
    Call Validacion(KeyAscii, "0123456789")
    If Index = 0 Then
        Call Nuevo(rstFacIDB, 0, "Fact_IDB", KeyAscii)
    ElseIf Index = 1 Then
        Call Nuevo(rstFacPre, 1, "Fact_Presupuestados", KeyAscii)
    End If
    End Sub


    Private Sub txt_KeyPress(Index%, KeyAscii%)
    If KeyAscii = Asc(".") Then KeyAscii = Asc(",")
    Call Validacion(KeyAscii, "0123456789,")
    End Sub

    Private Sub txtBackUp_GotFocus()
    With txtBackUp
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
    End Sub

    Private Sub txtBackUp_KeyPress(KeyAscii As Integer)
    mskBackUp(0).PromptInclude = True
    Call Validacion(KeyAscii, "1234567890")
    End Sub

    Private Sub activar(strV As String)
    '
    For I = 0 To 1
    
        fraOperativo(I).Enabled = strV
        optBackUp(I).Enabled = strV
        lblBackUp(I + 1).Enabled = strV
        lblBackUp(I + 3).Enabled = strV
        
    Next
    mskBackUp(0).Enabled = strV
    cmbBackUp.Enabled = strV
    txtBackUp.Enabled = strV
    '
    End Sub

    
    Private Sub ConfigList(rst As ADODB.Recordset, Indice%, Tabla$)
    '
    With rst
        .Open Tabla, cnnConexion, adOpenKeyset, adLockOptimistic, adCmdTable
        If Not .EOF Or Not .BOF Then
            .MoveFirst
            Do
                List1(Indice).AddItem !codGasto
                .MoveNext
            Loop Until .EOF
        End If
    End With
    '
    End Sub

    Private Sub Del(rst As ADODB.Recordset, Indice%, Tabla$)
    '
    With rst
        .MoveFirst
        .Find "CodGasto='" & List1(Indice).List(List1(Indice).ListIndex) & "'"
        If Not .EOF Then
            Call rtnBitacora("Eliminar Cod. Gasto '" & List1(Indice).List(List1(Indice).ListIndex) _
            & "' de la tabla " & Tabla)
            .Delete
            List1(Indice).RemoveItem List1(Indice).ListIndex
        End If
        '
    End With
    '
    End Sub

    Private Sub Nuevo(rst As ADODB.Recordset, Indice%, Tabla$, Tecla%)
    '
    If Tecla = 13 And Text1(Indice) <> "" Then
    
        List1(Indice).AddItem Text1(Indice)
        Call rtnBitacora("Agregar Cod. Gasto '" & Text1(Indice) & "' a la tabla " & Tabla)
        rst.AddNew "CodGasto", Text1(Indice)
        Text1(Indice) = ""
        
    End If
    '
    End Sub

