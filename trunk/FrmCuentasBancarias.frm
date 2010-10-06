VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form FrmCuentasBancarias 
   Caption         =   "Cuentas Bancarias"
   ClientHeight    =   15
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   1665
   ControlBox      =   0   'False
   Icon            =   "FrmCuentasBancarias.frx":0000
   MDIChild        =   -1  'True
   ScaleHeight     =   15
   ScaleWidth      =   1665
   Tag             =   "1"
   WindowState     =   2  'Maximized
   Begin ComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   480
      Left            =   0
      TabIndex        =   49
      Top             =   0
      Width           =   1665
      _ExtentX        =   2937
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
            Key             =   "Edit"
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
      MouseIcon       =   "FrmCuentasBancarias.frx":000C
      Begin VB.Frame Frame3 
         BackColor       =   &H00C0C0C0&
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
         Left            =   3660
         TabIndex        =   56
         Top             =   3315
         Width           =   5520
         Begin VB.TextBox Text4 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            Enabled         =   0   'False
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
            Left            =   1320
            TabIndex        =   59
            Top             =   825
            Width           =   690
         End
         Begin VB.TextBox Text3 
            BackColor       =   &H00FFFFFF&
            Height          =   285
            Left            =   1020
            TabIndex        =   58
            Top             =   285
            Width           =   4335
         End
         Begin VB.CommandButton Command1 
            Height          =   330
            Left            =   4935
            Picture         =   "FrmCuentasBancarias.frx":0326
            Style           =   1  'Graphical
            TabIndex        =   57
            ToolTipText     =   "Buscar"
            Top             =   750
            Width           =   375
         End
         Begin VB.Label Label20 
            AutoSize        =   -1  'True
            Caption         =   "Total Bancos :"
            Height          =   195
            Left            =   165
            TabIndex        =   61
            Top             =   870
            Width           =   1035
         End
         Begin VB.Label Label19 
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
            TabIndex        =   60
            Top             =   315
            Width           =   630
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00C0C0C0&
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
         Left            =   180
         TabIndex        =   53
         Top             =   3330
         Width           =   3210
         Begin VB.OptionButton OptBusca 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Por Número de Cuenta"
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
            Index           =   3
            Left            =   120
            TabIndex        =   55
            Top             =   825
            Width           =   2355
         End
         Begin VB.OptionButton OptBusca 
            BackColor       =   &H00C0C0C0&
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
            Index           =   2
            Left            =   120
            TabIndex        =   54
            Top             =   390
            Value           =   -1  'True
            Width           =   2355
         End
      End
   End
   Begin MSAdodcLib.Adodc adoCuentas 
      Height          =   330
      Index           =   1
      Left            =   6435
      Tag             =   "SELECT * FROM Monedas"
      Top             =   5190
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
      Caption         =   "Moneda"
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
   Begin MSAdodcLib.Adodc adoCuentas 
      Height          =   330
      Index           =   0
      Left            =   4005
      Tag             =   "SELECT * FROM Bancos ORDER BY NombreBanco"
      Top             =   5190
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
      Caption         =   "Bancos"
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
   Begin VB.Frame FraCuentas 
      Enabled         =   0   'False
      Height          =   2355
      Index           =   0
      Left            =   330
      TabIndex        =   1
      Top             =   1035
      Visible         =   0   'False
      Width           =   9105
      Begin VB.CheckBox chkP 
         Alignment       =   1  'Right Justify
         Caption         =   "Predeterminada"
         DataField       =   "Pred"
         Height          =   270
         Left            =   7395
         TabIndex        =   67
         Top             =   1740
         Width           =   1515
      End
      Begin VB.CheckBox chk 
         Alignment       =   1  'Right Justify
         Caption         =   "Inactiva"
         DataField       =   "Inactiva"
         Height          =   270
         Left            =   6030
         TabIndex        =   66
         Top             =   1740
         Width           =   1110
      End
      Begin VB.TextBox TxtCuentas 
         DataField       =   "Titular"
         DataSource      =   "adoCuentas(3)"
         Height          =   315
         Index           =   2
         Left            =   1200
         TabIndex        =   64
         Text            =   " "
         Top             =   1740
         Width           =   4620
      End
      Begin VB.TextBox TxtCuentas 
         DataField       =   "Sucursal"
         DataSource      =   "adoCuentas(3)"
         Height          =   315
         Index           =   1
         Left            =   5820
         TabIndex        =   9
         Text            =   " "
         Top             =   825
         Width           =   3105
      End
      Begin VB.TextBox TxtCuentas 
         DataField       =   "NumCuenta"
         DataSource      =   "adoCuentas(3)"
         Height          =   315
         Index           =   0
         Left            =   4290
         TabIndex        =   5
         Text            =   " "
         Top             =   375
         Width           =   4620
      End
      Begin MSDataListLib.DataCombo DtcCuentas 
         Bindings        =   "FrmCuentasBancarias.frx":0428
         DataField       =   "Moneda"
         DataSource      =   "adoCuentas(3)"
         Height          =   315
         Index           =   1
         Left            =   8115
         TabIndex        =   15
         Top             =   1290
         Width           =   825
         _ExtentX        =   1455
         _ExtentY        =   556
         _Version        =   393216
         ListField       =   "IdMoneda"
         Text            =   " "
      End
      Begin MSDataListLib.DataCombo DtcCuentas 
         Bindings        =   "FrmCuentasBancarias.frx":0444
         Height          =   315
         Index           =   0
         Left            =   1185
         TabIndex        =   7
         Top             =   825
         Width           =   3210
         _ExtentX        =   5662
         _ExtentY        =   556
         _Version        =   393216
         ListField       =   "NombreBanco"
         BoundColumn     =   ""
         Text            =   " "
      End
      Begin MSMask.MaskEdBox mskTelefono 
         DataField       =   "Telefono"
         DataSource      =   "adoCuentas(3)"
         Height          =   315
         Index           =   0
         Left            =   1200
         TabIndex        =   11
         Top             =   1290
         Width           =   1500
         _ExtentX        =   2646
         _ExtentY        =   556
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   16
         Format          =   "(####)-###-##-##"
         Mask            =   "(####)-###-##-##"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox mskTelefono 
         DataField       =   "Telefono1"
         DataSource      =   "adoCuentas(3)"
         Height          =   315
         Index           =   1
         Left            =   2880
         TabIndex        =   12
         Top             =   1290
         Width           =   1500
         _ExtentX        =   2646
         _ExtentY        =   556
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   16
         Format          =   "(####)-###-##-##"
         Mask            =   "(####)-###-##-##"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox mskTelefono 
         Bindings        =   "FrmCuentasBancarias.frx":0460
         DataField       =   "FechaApertura"
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "dd/MM/yyyy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   8202
            SubFormatType   =   0
         EndProperty
         DataSource      =   "adoCuentas(3)"
         Height          =   315
         Index           =   2
         Left            =   5820
         TabIndex        =   65
         TabStop         =   0   'False
         ToolTipText     =   "Documento Fecha 1"
         Top             =   1290
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
      Begin VB.Label LblCuentas 
         Caption         =   "Titular :"
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
         Index           =   14
         Left            =   195
         TabIndex        =   63
         Top             =   1755
         Width           =   1005
      End
      Begin VB.Line Line1 
         X1              =   2715
         X2              =   2820
         Y1              =   1590
         Y2              =   1290
      End
      Begin VB.Label LblCuentas 
         Alignment       =   1  'Right Justify
         Caption         =   "Fecha de Inicio :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   6
         Left            =   4170
         TabIndex        =   13
         Top             =   1335
         Width           =   1605
      End
      Begin VB.Label LblCuentas 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         DataField       =   "IDCuenta"
         DataSource      =   "adoCuentas(3)"
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
         Index           =   7
         Left            =   1200
         TabIndex        =   3
         Top             =   375
         Width           =   1005
      End
      Begin VB.Label LblCuentas 
         Caption         =   "Código :"
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
         Left            =   195
         TabIndex        =   2
         Top             =   390
         Width           =   1005
      End
      Begin VB.Label LblCuentas 
         Caption         =   "Banco :"
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
         Index           =   1
         Left            =   195
         TabIndex        =   6
         Top             =   840
         Width           =   1005
      End
      Begin VB.Label LblCuentas 
         Alignment       =   1  'Right Justify
         Caption         =   "Sucursal :"
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
         Left            =   4725
         TabIndex        =   8
         Top             =   840
         Width           =   1005
      End
      Begin VB.Label LblCuentas 
         Caption         =   "Teléfonos :"
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
         Left            =   195
         TabIndex        =   10
         Top             =   1305
         Width           =   1005
      End
      Begin VB.Label LblCuentas 
         Alignment       =   1  'Right Justify
         Caption         =   "Moneda :"
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
         Left            =   7140
         TabIndex        =   14
         Top             =   1342
         Width           =   900
      End
      Begin VB.Label LblCuentas 
         Caption         =   "Número de Cuenta :"
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
         Left            =   2565
         TabIndex        =   4
         Top             =   427
         Width           =   1605
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   5715
      Left            =   105
      TabIndex        =   0
      Top             =   615
      Width           =   9615
      _ExtentX        =   16960
      _ExtentY        =   10081
      _Version        =   393216
      Tab             =   2
      TabsPerRow      =   4
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
      TabCaption(0)   =   "     Datos Generales     "
      TabPicture(0)   =   "FrmCuentasBancarias.frx":0482
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "FraCuentas(1)"
      Tab(0).Control(1)=   "adoCuentas(2)"
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "  Datos Adicionales"
      TabPicture(1)   =   "FrmCuentasBancarias.frx":049E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "FraCuentas(2)"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "      Lista      "
      TabPicture(2)   =   "FrmCuentasBancarias.frx":04BA
      Tab(2).ControlEnabled=   -1  'True
      Tab(2).Control(0)=   "DtgCuentas"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "FraCuentas(3)"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "FraCuentas(4)"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).ControlCount=   3
      Begin MSAdodcLib.Adodc adoCuentas 
         Height          =   330
         Index           =   2
         Left            =   -73560
         Tag             =   "SELECT Cuentas.*,Bancos.NombreBanco FROM Cuentas INNER JOIN Bancos ON Bancos.IdBanco=Cuentas.IDBanco"
         Top             =   4590
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
         Caption         =   "Grid"
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
      Begin VB.Frame FraCuentas 
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
         Index           =   4
         Left            =   3900
         TabIndex        =   43
         Top             =   3773
         Width           =   5520
         Begin VB.TextBox TxtCuentas 
            BackColor       =   &H00FFFFFF&
            Height          =   285
            Index           =   15
            Left            =   1020
            TabIndex        =   45
            Top             =   285
            Width           =   4335
         End
         Begin VB.CommandButton BotBusca 
            Height          =   330
            Left            =   4935
            Picture         =   "FrmCuentasBancarias.frx":04D6
            Style           =   1  'Graphical
            TabIndex        =   46
            ToolTipText     =   "Buscar"
            Top             =   750
            Width           =   375
         End
         Begin VB.Label Label17 
            AutoSize        =   -1  'True
            Caption         =   "Total Bancos :"
            Height          =   195
            Left            =   165
            TabIndex        =   62
            Top             =   870
            Width           =   1035
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
            TabIndex        =   44
            Top             =   315
            Width           =   630
         End
      End
      Begin VB.Frame FraCuentas 
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
         Index           =   3
         Left            =   390
         TabIndex        =   40
         Top             =   3780
         Width           =   3210
         Begin VB.OptionButton OptBusca 
            Caption         =   "Por Número de Cuenta"
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
            Left            =   120
            TabIndex        =   42
            Top             =   825
            Width           =   2355
         End
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
            Left            =   120
            TabIndex        =   41
            Top             =   390
            Value           =   -1  'True
            Width           =   2355
         End
         Begin MSAdodcLib.Adodc adoCuentas 
            Height          =   330
            Index           =   3
            Left            =   1170
            Tag             =   "SELECT * FROM Cuentas"
            Top             =   90
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
            Caption         =   "Cuentas"
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
      Begin VB.Frame FraCuentas 
         Height          =   2385
         Index           =   2
         Left            =   -74790
         TabIndex        =   52
         Top             =   2865
         Width           =   9015
         Begin VB.TextBox TxtCuentas 
            DataField       =   "Campo1"
            DataSource      =   "AdoBancos"
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
            Index           =   7
            Left            =   1395
            TabIndex        =   24
            Top             =   285
            Width           =   2565
         End
         Begin VB.TextBox TxtCuentas 
            DataField       =   "Campo2"
            DataSource      =   "AdoBancos"
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
            Index           =   8
            Left            =   1425
            TabIndex        =   26
            Top             =   780
            Width           =   2565
         End
         Begin VB.TextBox TxtCuentas 
            DataField       =   "Campo3"
            DataSource      =   "AdoBancos"
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
            Index           =   9
            Left            =   1410
            TabIndex        =   28
            Top             =   1275
            Width           =   2565
         End
         Begin VB.TextBox TxtCuentas 
            DataField       =   "Campo4"
            DataSource      =   "AdoBancos"
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
            Index           =   10
            Left            =   1395
            TabIndex        =   30
            Top             =   1815
            Width           =   2565
         End
         Begin VB.TextBox TxtCuentas 
            DataField       =   "Campo5"
            DataSource      =   "AdoBancos"
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
            Index           =   11
            Left            =   6210
            TabIndex        =   32
            Top             =   285
            Width           =   2565
         End
         Begin VB.TextBox TxtCuentas 
            DataField       =   "Campo6"
            DataSource      =   "AdoBancos"
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
            Index           =   12
            Left            =   6195
            TabIndex        =   34
            Top             =   795
            Width           =   2565
         End
         Begin VB.TextBox TxtCuentas 
            DataField       =   "Campo7"
            DataSource      =   "AdoBancos"
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
            Index           =   13
            Left            =   6210
            TabIndex        =   36
            Top             =   1290
            Width           =   2565
         End
         Begin VB.TextBox TxtCuentas 
            DataField       =   "Campo8"
            DataSource      =   "AdoBancos"
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
            Index           =   14
            Left            =   6210
            TabIndex        =   38
            Top             =   1800
            Width           =   2565
         End
         Begin VB.Label Label1 
            Caption         =   "Campo 6"
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
            Index           =   8
            Left            =   4950
            TabIndex        =   33
            Top             =   855
            Width           =   810
         End
         Begin VB.Label Label1 
            Caption         =   "Campo 1"
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
            Index           =   1
            Left            =   225
            TabIndex        =   23
            Top             =   315
            Width           =   750
         End
         Begin VB.Label Label1 
            Caption         =   "Campo 2"
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
            Left            =   195
            TabIndex        =   25
            Top             =   810
            Width           =   780
         End
         Begin VB.Label Label1 
            Caption         =   "Campo 3"
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
            Left            =   195
            TabIndex        =   27
            Top             =   1290
            Width           =   765
         End
         Begin VB.Label Label1 
            Caption         =   "Campo 4"
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
            Left            =   195
            TabIndex        =   29
            Top             =   1830
            Width           =   810
         End
         Begin VB.Label Label1 
            Caption         =   "Campo 5"
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
            Left            =   4950
            TabIndex        =   31
            Top             =   360
            Width           =   735
         End
         Begin VB.Label Label1 
            Caption         =   "Campo 7"
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
            Index           =   6
            Left            =   4995
            TabIndex        =   35
            Top             =   1395
            Width           =   810
         End
         Begin VB.Label Label1 
            Caption         =   "Campo 8"
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
            Left            =   4980
            TabIndex        =   37
            Top             =   1875
            Width           =   825
         End
      End
      Begin VB.Frame FraCuentas 
         Enabled         =   0   'False
         Height          =   1545
         Index           =   1
         Left            =   -74775
         TabIndex        =   16
         Top             =   2895
         Width           =   9135
         Begin VB.TextBox TxtCuentas 
            DataField       =   "Saldo_ic"
            DataSource      =   "AdoCuentas"
            Height          =   315
            Index           =   3
            Left            =   2625
            TabIndex        =   18
            Text            =   "0,00"
            Top             =   293
            Width           =   2220
         End
         Begin VB.TextBox TxtCuentas 
            DataField       =   "Saldo_c"
            DataSource      =   "AdoCuentas"
            Height          =   315
            Index           =   4
            Left            =   2625
            TabIndex        =   20
            Text            =   "0,00"
            Top             =   885
            Width           =   2220
         End
         Begin VB.TextBox TxtCuentas 
            DataField       =   "SaldoInicial"
            DataSource      =   "AdoCuentas"
            Height          =   315
            Index           =   5
            Left            =   6525
            TabIndex        =   48
            Text            =   "0,00"
            Top             =   293
            Width           =   2220
         End
         Begin VB.TextBox TxtCuentas 
            DataField       =   "Saldo_a"
            DataSource      =   "AdoCuentas"
            Height          =   315
            Index           =   6
            Left            =   6525
            TabIndex        =   47
            Text            =   "0,00"
            Top             =   900
            Width           =   2220
         End
         Begin VB.Label LblCuentas 
            AutoSize        =   -1  'True
            Caption         =   "Saldo Conciliado :"
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
            Left            =   300
            TabIndex        =   51
            Top             =   675
            Width           =   1410
         End
         Begin VB.Label LblCuentas 
            AutoSize        =   -1  'True
            Caption         =   "Inicial :"
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
            Left            =   2025
            TabIndex        =   17
            Top             =   360
            Width           =   555
         End
         Begin VB.Label LblCuentas 
            AutoSize        =   -1  'True
            Caption         =   "Inicial :"
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
            Left            =   5910
            TabIndex        =   21
            Top             =   360
            Width           =   555
         End
         Begin VB.Label LblCuentas 
            AutoSize        =   -1  'True
            Caption         =   "Actual :"
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
            Left            =   5835
            TabIndex        =   22
            Top             =   945
            Width           =   630
         End
         Begin VB.Label LblCuentas 
            AutoSize        =   -1  'True
            Caption         =   "Actual :"
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
            Left            =   1980
            TabIndex        =   19
            Top             =   945
            Width           =   630
         End
         Begin VB.Label LblCuentas 
            AutoSize        =   -1  'True
            Caption         =   "Saldo :"
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
            Index           =   13
            Left            =   5235
            TabIndex        =   50
            Top             =   645
            Width           =   555
         End
      End
      Begin MSDataGridLib.DataGrid DtgCuentas 
         Bindings        =   "FrmCuentasBancarias.frx":0A60
         Height          =   3105
         Left            =   210
         TabIndex        =   39
         Top             =   450
         Width           =   9270
         _ExtentX        =   16351
         _ExtentY        =   5477
         _Version        =   393216
         AllowUpdate     =   0   'False
         BorderStyle     =   0
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
         ColumnCount     =   4
         BeginProperty Column00 
            DataField       =   "IdCuenta"
            Caption         =   "Código"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   2058
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   "NumCuenta"
            Caption         =   "Cuenta Número"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   2058
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column02 
            DataField       =   "NombreBanco"
            Caption         =   "Banco"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   2058
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column03 
            DataField       =   "Sucursal"
            Caption         =   "Agencia"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   2058
               SubFormatType   =   0
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            ScrollBars      =   2
            Locked          =   -1  'True
            BeginProperty Column00 
               ColumnWidth     =   945,071
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   1920,189
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   3449,764
            EndProperty
            BeginProperty Column03 
               ColumnWidth     =   2129,953
            EndProperty
         EndProperty
      End
   End
   Begin ComctlLib.ImageList ImageList1 
      Left            =   7560
      Top             =   330
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
            Picture         =   "FrmCuentasBancarias.frx":0A7C
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmCuentasBancarias.frx":0BFE
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmCuentasBancarias.frx":0D80
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmCuentasBancarias.frx":0F02
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmCuentasBancarias.frx":1084
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmCuentasBancarias.frx":1206
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmCuentasBancarias.frx":1388
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmCuentasBancarias.frx":150A
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmCuentasBancarias.frx":168C
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmCuentasBancarias.frx":180E
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmCuentasBancarias.frx":1990
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmCuentasBancarias.frx":1B12
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "FrmCuentasBancarias"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
    Private CuentasCnn As ADODB.Connection
    Private blnEdit As Boolean
    
    '---------------------------------------------------------------------------------------------
    Private Sub DtcCuentas_Click(Index As Integer, Area As Integer) '
    '---------------------------------------------------------------------------------------------
    '
    If Area = 2 Then
    '
        Select Case Index
    '
            Case 0: Call rtnBBanco
            Case 1: TxtCuentas(3).SetFocus
                    
        End Select
    '
    End If
    '
    End Sub

    '---------------------------------------------------------------------------------------------
    Private Sub DtcCuentas_KeyPress(Index As Integer, KeyAscii As Integer)  '
    '---------------------------------------------------------------------------------------------
    '
    Select Case Index
    '
        Case 0
    '   ---------------------
            If KeyAscii = 13 Then Call rtnBBanco
            KeyAscii = 0
       End Select
    '
    End Sub

    '---------------------------------------------------------------------------------------------
    Private Sub DtgCuentas_Click()  '
    '---------------------------------------------------------------------------------------------
    '
    With adoCuentas(3).Recordset
        If .EOF Then Exit Sub
        .MoveFirst
        .Find "IDcuenta =" & adoCuentas(2).Recordset!IDCuenta
        DtcCuentas(0) = fntBanco
        SSTab1.Tab = 0
    End With
    
    '
    End Sub

    '---------------------------------------------------------------------------------------------
    Private Sub DtpCuentas_KeyUp(KeyCode As Integer, Shift As Integer)  '
    '---------------------------------------------------------------------------------------------
    If KeyCode = 13 Then DtcCuentas(1).SetFocus
    '
    End Sub

    '---------------------------------------------------------------------------------------------
    Private Sub Form_Load() '
    '---------------------------------------------------------------------------------------------
    Screen.MousePointer = vbHourglass
    Set CuentasCnn = New ADODB.Connection
    'CREA UNA CONEXION Y UN ADODB.Recordset
    CuentasCnn.CursorLocation = adUseClient
    CuentasCnn.Open cnnOLEDB + mcDatos
    '
    For i = 0 To 3
        adoCuentas(i).ConnectionString = cnnOLEDB + mcDatos
        adoCuentas(i).RecordSource = adoCuentas(i).Tag
        adoCuentas(i).Refresh
    Next
    '
    If adoCuentas(3).Recordset.RecordCount > 0 Then
        DtcCuentas(0) = fntBanco
    End If
    '
    Set chk.DataSource = adoCuentas(3)
    Set chkP.DataSource = adoCuentas(3)
    '
    Screen.MousePointer = vbDefault
    End Sub

    Private Sub Form_Resize()
    '
    If FrmCuentasBancarias.WindowState <> vbMinimized Then
        With SSTab1
            .Left = (FrmCuentasBancarias.Width - .Width) / 2
            FraCuentas(0).Left = (SSTab1.Left - 105) + 360
        End With
    End If
    '
    End Sub

    Private Sub Form_Unload(Cancel As Integer)
    CuentasCnn.Close: Set CuentasCnn = Nothing
    End Sub

    '---------------------------------------------------------------------------------------------
    Private Sub MskTelefono_KeyPress(Index As Integer, KeyAscii As Integer) '
    '---------------------------------------------------------------------------------------------
    '
    If KeyAscii = 13 Then
        Select Case Index
            Case 0
    '       ---------------------
                mskTelefono(1).SetFocus
            Case 1
    '       ---------------------
                mskTelefono(2).SetFocus
        End Select
    End If
    '
    End Sub

    '---------------------------------------------------------------------------------------------
    Private Sub SSTab1_Click(PreviousTab As Integer)    '
    '---------------------------------------------------------------------------------------------
    '
    Select Case SSTab1.Tab
        
        Case 0
    '   ---------------------
            FraCuentas(0).Visible = True
        Case 1
    '   ---------------------
            FraCuentas(0).Visible = True
        Case 2
    '   ---------------------
            FraCuentas(0).Visible = False
            adoCuentas(2).Refresh
    
    End Select
    '
    End Sub

    '---------------------------------------------------------------------------------------------
    Private Sub Toolbar1_ButtonClick(ByVal Button As ComctlLib.Button)  '
    '---------------------------------------------------------------------------------------------
    '
    With adoCuentas(3).Recordset
        
        Select Case UCase(Button.Key)
    '
            Case "FIRst"    'Primer Registro
    '       ---------------------
                If .EOF Then Exit Sub
                .MoveFirst
                DtcCuentas(0) = fntBanco
            
            Case "PREVIOUS" 'Registro Anterior
    '       ---------------------
                If .EOF Then Exit Sub
                .MovePrevious
                If .BOF Then .MoveLast
                DtcCuentas(0) = fntBanco
            
            Case "NEXT" 'Registro siguiente
    '       ---------------------
                If .EOF Then Exit Sub
                .MoveNext
                If .EOF Then .MoveFirst
                DtcCuentas(0) = fntBanco
        
            Case "END"  'IR AL ULTIMO
    '       ---------------------
                If .EOF Then Exit Sub
                .MoveLast
                If .EOF Then Exit Sub
                DtcCuentas(0) = fntBanco
            
            Case "NEW"  'Agregar Registro
    '       ---------------------
                FraCuentas(0).Enabled = True
                FraCuentas(1).Enabled = True
                .AddNew
                Call RtnEstado(5, Toolbar1)
                LblCuentas(7) = FntMax
                SSTab1.Tab = 0
                TxtCuentas(0).SetFocus
                mskTelefono(2) = Date
                DtcCuentas(0) = ""
                
        
            Case "SAVE"     'Actualizar Registro
    '       ---------------------
                If ftnValidate = True Then Exit Sub
                Screen.MousePointer = vbHourglass
                For i = 0 To 2
                    mskTelefono(i).PromptInclude = True
                Next
                adoCuentas(2).Recordset.Find "NombreBanco='" & DtcCuentas(0) & "'"
                !IDBanco = adoCuentas(0).Recordset!IDBanco
                Call rtnBitacora(IIf(.EditMode = adEditAdd, "Adición ", "Edición ") & "cuenta ban" _
                & "caria Nº " & !NumCuenta & " Inm:" & gcCodInm)
                
                If .EditMode = adEditInProgress And TxtCuentas(0).Tag <> TxtCuentas(0) Then
                
                    cnnConexion.Execute "UPDATE Cheque SET Cuenta='" & TxtCuentas(0) & "' WHERE" _
                    & " Cuenta='" & TxtCuentas(0).Tag & "'"
                    
                End If
                TxtCuentas(0).Tag = ""
                .Update
                For i = 0 To 2
                    mskTelefono(i).PromptInclude = False
                Next
                FraCuentas(0).Enabled = False
                FraCuentas(1).Enabled = False
                '
                adoCuentas(2).Refresh
                With adoCuentas(2).Recordset
                    '
                    .Requery
                    If Not .EOF Or Not .BOF Then .MoveFirst: i = 1
                    Do Until .EOF Or i > 3
                        'actualiza las cuentas de los inmuebles en la tabla caja
                        strCuenta = !NombreBanco & " " & !NumCuenta
                        cnnConexion.Execute "UPDATE Caja SET Caja.Cuenta" & i & " = '" & _
                        strCuenta & "' WHERE CodigoCaja IN (SELECT Caja FROM Inmueble WHERE Cod" _
                        & "Inm='" & gcCodInm & "');"
                        .MoveNext: i = i + 1
                    Loop
                    '
                End With
                
                blnEdit = False
                'DtgCuentas.ReBind
                MsgBox "Registro Actualizado", vbInformation, App.ProductName
                Screen.MousePointer = vbDefault
                Call RtnEstado(6, Toolbar1)
                '
            Case "FIND" 'Buscar Registro
    '       ---------------------
                SSTab1.Tab = 3
                TxtBus.SetFocus
            
            Case "UNDO"   'Cancelar
    '       ---------------------
                For i = 0 To 2
                    mskTelefono(i).PromptInclude = True
                Next
                .CancelUpdate
                For i = 0 To 2
                    mskTelefono(i).PromptInclude = False
                Next
                FraCuentas(0).Enabled = False
                FraCuentas(1).Enabled = False
                Call RtnEstado(6, Toolbar1)
                blnEdit = False
                
            Case "DELETE"   'Eliminar Registro
    '       ---------------------
                MsgBox "Opción No disponible....Por ahora", vbInformation, App.ProductName
                Exit Sub
                Dim Confirma As Integer
                Confirma = MsgBox("Confirma eliminar el registro actual ?", vbOKCancel, "Eliminar Registro")
                If Confirma = vbOK Then
                    .Delete
                    .MoveNext
                    If .EOF Then
                        .MoveLast
                    End If
                End If
            
            Case "EDIT" 'Modificar Registro
    '       -----------------
                blnEdit = True
                FraCuentas(0).Enabled = True
                FraCuentas(1).Enabled = True
                Call rtnBBanco
                SSTab1.Tab = 0
                TxtCuentas(0).SetFocus
                TxtCuentas(0).Tag = TxtCuentas(0)
                Call RtnEstado(5, Toolbar1)
                
            
            Case "PRINT"    'Imprimir Reporte
            
            Case "CLOSE"    'Cerrar Formulario
                Unload Me
                Set FrmCuentasBancarias = Nothing
    '
        End Select
    '
    End With
    '
    End Sub

    '---------------------------------------------------------------------------------------------
    Private Function FntMax()   '
    '---------------------------------------------------------------------------------------------
    '
    Dim RstCtasMax As ADODB.Recordset
    Set RstCtasMax = New ADODB.Recordset
    '
    RstCtasMax.Open "SELECT MAX(IdCuenta) as M FROM Cuentas;", CuentasCnn, adOpenKeyset, adLockOptimistic
    FntMax = IIf(IsNull(RstCtasMax!m), 0, RstCtasMax!m) + 1
    RstCtasMax.Close: Set RstCtasMax = Nothing
    '
    End Function
    
    '---------------------------------------------------------------------------------------------
    Private Function fntBanco()
    '---------------------------------------------------------------------------------------------
    '
    With adoCuentas(2).Recordset
        .MoveFirst
        .Find "IDCuenta =" & adoCuentas(3).Recordset!IDCuenta
        fntBanco = !NombreBanco
    End With
    '
    End Function

    
    '---------------------------------------------------------------------------------------------
    Private Sub TxtCuentas_KeyPress(Index As Integer, KeyAscii As Integer)  '
    '---------------------------------------------------------------------------------------------
    '
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    Select Case Index
        Case 0
    '   ---------------------
            Call Validacion(KeyAscii, "1234567890")
            If KeyAscii = 13 Then DtcCuentas(0).SetFocus
        Case 1
    '   ---------------------
            If KeyAscii = 13 Then mskTelefono(0).SetFocus
        Case 3, 4, 5, 6
    '   ---------------------
            If KeyAscii = 46 Then KeyAscii = 44
            Call Validacion(KeyAscii, "1234567890.,")
            If KeyAscii = 13 Then If Index <> 6 Then TxtCuentas(Index + 1).SetFocus
    '
    End Select
    '
    End Sub
    
    
    '---------------------------------------------------------------------------------------------
    Private Function ftnValidate() As Boolean   '
    '---------------------------------------------------------------------------------------------
    '
    If LblCuentas(7) = "" Then ftnValidate = MsgBox("La Operacion generó errores en alguno de s" _
    & "us pasos, presione el boton 'Cancelar'" & vbCrLf & "e intente guardar el registro nuevam" _
    & "ente, si el problema persiste, consulte al administrador del sistema")
    '
    If TxtCuentas(0) = "" Then ftnValidate = MsgBox("Falta Número de Cuenta..", _
    vbInformation + vbOKOnly)
    '
    If DtcCuentas(0) = "" Then ftnValidate = MsgBox("Debe seleccionar el banco a/c de la cuenta.." _
    , vbInformation + vbOKOnly)
    '
    If DtcCuentas(1) = "" Then ftnValidate = MsgBox("Falta Tipo de Modena..", _
    vbInformation + vbOKOnly)
    '
    If blnEdit = False Then
        With adoCuentas(2).Recordset
            If .EOF Then Exit Function
            .MoveFirst
            .Find "NumCuenta ='" & TxtCuentas(0) & "'"
            If Not .EOF Then
                ftnValidate = MsgBox("Número de Cuenta ya registrado..", vbExclamation + vbOKOnly)
            End If
        End With
    End If
    '
    End Function
    
    '---------------------------------------------------------------------------------------------
    Private Sub rtnBBanco() '
    '---------------------------------------------------------------------------------------------
    '
    With adoCuentas(0).Recordset
        .MoveFirst
        .Find "NombreBanco ='" & DtcCuentas(0) & "'"
        If .EOF Then
            MsgBox "Vuelva a Seleccionar el Banco de la Lista, por favor", vbInformation, _
            App.ProductName
        End If
        If SSTab1.Tab = 0 Then TxtCuentas(1).SetFocus
    End With
    '
    End Sub
    

