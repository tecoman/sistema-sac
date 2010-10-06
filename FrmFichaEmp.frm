VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form FrmFichaEmp 
   Caption         =   "Ficha de Empleados"
   ClientHeight    =   15
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   1965
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   MDIChild        =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   15
   ScaleWidth      =   1965
   Tag             =   "1"
   WindowState     =   2  'Maximized
   Begin ComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   480
      Left            =   0
      TabIndex        =   44
      Top             =   0
      Width           =   1965
      _ExtentX        =   3466
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
            Object.Width           =   1e-4
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
   Begin VB.Frame fraEmp 
      Enabled         =   0   'False
      Height          =   2715
      Index           =   0
      Left            =   1005
      TabIndex        =   58
      Top             =   1005
      Width           =   9525
      Begin VB.ComboBox cmbStatus 
         DataSource      =   "adoEmp(0)"
         Height          =   315
         ItemData        =   "FrmFichaEmp.frx":0000
         Left            =   6075
         List            =   "FrmFichaEmp.frx":0010
         TabIndex        =   110
         Top             =   240
         Width           =   3240
      End
      Begin VB.ComboBox cmbSexo 
         DataField       =   "Sexo"
         DataSource      =   "adoEmp(0)"
         Height          =   315
         ItemData        =   "FrmFichaEmp.frx":0039
         Left            =   4200
         List            =   "FrmFichaEmp.frx":0043
         TabIndex        =   11
         Top             =   1185
         Width           =   1455
      End
      Begin VB.TextBox txtEmp 
         DataField       =   "CodEmp"
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "000000"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   8202
            SubFormatType   =   0
         EndProperty
         DataSource      =   "adoEmp(0)"
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
         Index           =   42
         Left            =   4020
         Locked          =   -1  'True
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   270
         Width           =   960
      End
      Begin VB.CheckBox ChkInactivo 
         Caption         =   "&Zurdo"
         DataField       =   "Zurdo"
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
         Index           =   0
         Left            =   7875
         TabIndex        =   19
         Top             =   2145
         Width           =   1350
      End
      Begin VB.TextBox txtEmp 
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "000000"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   8202
            SubFormatType   =   0
         EndProperty
         DataSource      =   "adoEmp(0)"
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
         Index           =   17
         Left            =   255
         Locked          =   -1  'True
         TabIndex        =   63
         Top             =   2235
         Visible         =   0   'False
         Width           =   540
      End
      Begin VB.TextBox txtEmp 
         DataField       =   "NSSO"
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "#,##0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   0
         EndProperty
         DataSource      =   "adoEmp(0)"
         Height          =   345
         Index           =   5
         Left            =   7875
         TabIndex        =   18
         Top             =   1650
         Width           =   1395
      End
      Begin VB.TextBox txtEmp 
         Alignment       =   1  'Right Justify
         DataField       =   "Cedula"
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "#,##0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   8202
            SubFormatType   =   0
         EndProperty
         DataSource      =   "adoEmp(0)"
         Height          =   315
         Index           =   3
         Left            =   2115
         TabIndex        =   9
         Top             =   1170
         Width           =   1260
      End
      Begin VB.TextBox txtEmp 
         DataField       =   "CodInm"
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "000000"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   0
         EndProperty
         DataSource      =   "adoEmp(0)"
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
         Left            =   1695
         TabIndex        =   1
         Top             =   270
         Width           =   960
      End
      Begin VB.TextBox txtEmp 
         DataField       =   "Apellidos"
         DataSource      =   "adoEmp(0)"
         Height          =   315
         Index           =   1
         Left            =   1695
         TabIndex        =   5
         Top             =   705
         Width           =   3300
      End
      Begin VB.TextBox txtEmp 
         DataField       =   "Nombres"
         DataSource      =   "adoEmp(0)"
         Height          =   315
         Index           =   2
         Left            =   6075
         TabIndex        =   7
         Top             =   705
         Width           =   3240
      End
      Begin VB.CommandButton cmdFIngreso 
         Height          =   255
         Index           =   0
         Left            =   8985
         Picture         =   "FrmFichaEmp.frx":005C
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   1215
         Width           =   270
      End
      Begin VB.CheckBox chkNac 
         Caption         =   "&V"
         DataField       =   "Nacionalidad"
         DataSource      =   "adoEmp(0)"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   0
         Left            =   1695
         TabIndex        =   60
         Top             =   1155
         Width           =   435
      End
      Begin VB.CheckBox chkNac 
         Caption         =   "&E"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   1
         Left            =   1695
         TabIndex        =   59
         Top             =   1350
         Width           =   375
      End
      Begin MSMask.MaskEdBox mskemp 
         Bindings        =   "FrmFichaEmp.frx":01A6
         DataField       =   "FNaci"
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "dd/MM/yyyy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   8202
            SubFormatType   =   3
         EndProperty
         DataSource      =   "adoEmp(0)"
         Height          =   315
         Index           =   0
         Left            =   7890
         TabIndex        =   13
         Top             =   1185
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   556
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   10
         Format          =   "dd/MM/yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.TextBox txtEmp 
         DataField       =   "Direccion"
         DataSource      =   "adoEmp(0)"
         Height          =   730
         Index           =   4
         Left            =   1695
         MaxLength       =   249
         MultiLine       =   -1  'True
         TabIndex        =   16
         Top             =   1740
         Width           =   5445
      End
      Begin VB.Label lblEmpleado 
         Caption         =   "Situación:"
         Height          =   315
         Index           =   23
         Left            =   5235
         TabIndex        =   111
         Top             =   285
         Width           =   780
      End
      Begin VB.Label lblEmpleado 
         Caption         =   "Código Emp.:"
         Height          =   315
         Index           =   22
         Left            =   2865
         TabIndex        =   2
         Top             =   285
         Width           =   1410
      End
      Begin VB.Label lblEmpleado 
         Alignment       =   1  'Right Justify
         Caption         =   "Nº S.S.:"
         Height          =   315
         Index           =   16
         Left            =   7140
         TabIndex        =   17
         Top             =   1710
         Width           =   675
      End
      Begin VB.Label lblEmpleado 
         Caption         =   "&Dirección:"
         Height          =   315
         Index           =   15
         Left            =   255
         TabIndex        =   15
         Top             =   1800
         Width           =   1410
      End
      Begin VB.Label lblEmpleado 
         Caption         =   "Código &Inmueble:"
         Height          =   315
         Index           =   7
         Left            =   255
         TabIndex        =   0
         Top             =   300
         Width           =   1410
      End
      Begin VB.Label lblEmpleado 
         Caption         =   "&Apellidos:"
         Height          =   315
         Index           =   8
         Left            =   255
         TabIndex        =   4
         Top             =   750
         Width           =   1410
      End
      Begin VB.Label lblEmpleado 
         Caption         =   "Nº de &Cédula:"
         Height          =   315
         Index           =   9
         Left            =   255
         TabIndex        =   8
         Top             =   1230
         Width           =   1410
      End
      Begin VB.Label lblEmpleado 
         Alignment       =   1  'Right Justify
         Caption         =   "Fecha de Nacimien&to:"
         Height          =   240
         Index           =   11
         Left            =   6210
         TabIndex        =   12
         Top             =   1230
         Width           =   1605
      End
      Begin VB.Label lblEmpleado 
         Caption         =   "&Nombres:"
         Height          =   315
         Index           =   12
         Left            =   5235
         TabIndex        =   6
         Top             =   720
         Width           =   780
      End
      Begin VB.Label lblEmpleado 
         Caption         =   "&Sexo:"
         Height          =   315
         Index           =   13
         Left            =   3750
         TabIndex        =   10
         Top             =   1215
         Width           =   525
      End
   End
   Begin TabDlg.SSTab SSTab2 
      Height          =   6705
      Left            =   675
      TabIndex        =   43
      Top             =   555
      Width           =   10185
      _ExtentX        =   17965
      _ExtentY        =   11827
      _Version        =   393216
      Tabs            =   4
      TabsPerRow      =   5
      TabHeight       =   520
      ShowFocusRect   =   0   'False
      TabCaption(0)   =   "Trabajador"
      TabPicture(0)   =   "FrmFichaEmp.frx":01C8
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "fraEmp(1)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Campos Adicionales"
      TabPicture(1)   =   "FrmFichaEmp.frx":01E4
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "fraEmp(2)"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "S.S.O / L.P.H."
      TabPicture(2)   =   "FrmFichaEmp.frx":0200
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "adoEmp(3)"
      Tab(2).Control(1)=   "adoEmp(2)"
      Tab(2).Control(2)=   "adoEmp(1)"
      Tab(2).Control(3)=   "adoEmp(0)"
      Tab(2).Control(4)=   "fraEmp(3)"
      Tab(2).ControlCount=   5
      TabCaption(3)   =   "Lista"
      TabPicture(3)   =   "FrmFichaEmp.frx":021C
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "DataGrid1"
      Tab(3).Control(1)=   "fraLis(1)"
      Tab(3).Control(2)=   "fraLis(0)"
      Tab(3).ControlCount=   3
      Begin VB.Frame fraEmp 
         Caption         =   "Declaración  de Familiares"
         Enabled         =   0   'False
         Height          =   2955
         Index           =   3
         Left            =   -74640
         TabIndex        =   73
         Top             =   3255
         Width           =   9525
         Begin VB.TextBox txtEmp 
            DataField       =   "NomApe1"
            DataSource      =   "adoEmp(0)"
            Height          =   315
            Index           =   37
            Left            =   4350
            TabIndex        =   87
            Top             =   1185
            Width           =   3300
         End
         Begin VB.TextBox txtEmp 
            DataField       =   "NomApe2"
            DataSource      =   "adoEmp(0)"
            Height          =   315
            Index           =   38
            Left            =   4350
            TabIndex        =   92
            Top             =   1500
            Width           =   3300
         End
         Begin VB.TextBox txtEmp 
            DataField       =   "NomApe3"
            DataSource      =   "adoEmp(0)"
            Height          =   315
            Index           =   39
            Left            =   4350
            TabIndex        =   97
            Top             =   1815
            Width           =   3300
         End
         Begin VB.TextBox txtEmp 
            DataField       =   "NomApe4"
            DataSource      =   "adoEmp(0)"
            Height          =   315
            Index           =   40
            Left            =   4350
            TabIndex        =   102
            Top             =   2130
            Width           =   3300
         End
         Begin VB.TextBox txtEmp 
            DataField       =   "NomApe5"
            DataSource      =   "adoEmp(0)"
            Height          =   315
            Index           =   41
            Left            =   4350
            TabIndex        =   107
            Top             =   2445
            Width           =   3300
         End
         Begin VB.TextBox txtEmp 
            DataField       =   "NomApe"
            DataSource      =   "adoEmp(0)"
            Height          =   315
            Index           =   36
            Left            =   4350
            TabIndex        =   82
            Top             =   870
            Width           =   3300
         End
         Begin VB.TextBox txtEmp 
            Alignment       =   2  'Center
            DataField       =   "Masculino1"
            BeginProperty DataFormat 
               Type            =   0
               Format          =   "#,##0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   8202
               SubFormatType   =   0
            EndProperty
            DataSource      =   "adoEmp(0)"
            Height          =   315
            Index           =   31
            Left            =   3195
            MaxLength       =   1
            TabIndex        =   86
            Top             =   1185
            Width           =   1155
         End
         Begin VB.TextBox txtEmp 
            Alignment       =   2  'Center
            DataField       =   "Masculino2"
            BeginProperty DataFormat 
               Type            =   0
               Format          =   "#,##0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   8202
               SubFormatType   =   0
            EndProperty
            DataSource      =   "adoEmp(0)"
            Height          =   315
            Index           =   32
            Left            =   3195
            MaxLength       =   1
            TabIndex        =   91
            Top             =   1500
            Width           =   1155
         End
         Begin VB.TextBox txtEmp 
            Alignment       =   2  'Center
            DataField       =   "Masculino3"
            BeginProperty DataFormat 
               Type            =   0
               Format          =   "#,##0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   8202
               SubFormatType   =   0
            EndProperty
            DataSource      =   "adoEmp(0)"
            Height          =   315
            Index           =   33
            Left            =   3195
            MaxLength       =   1
            TabIndex        =   96
            Top             =   1815
            Width           =   1155
         End
         Begin VB.TextBox txtEmp 
            Alignment       =   2  'Center
            DataField       =   "Masculino4"
            BeginProperty DataFormat 
               Type            =   0
               Format          =   "#,##0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   8202
               SubFormatType   =   0
            EndProperty
            DataSource      =   "adoEmp(0)"
            Height          =   315
            Index           =   34
            Left            =   3195
            MaxLength       =   1
            TabIndex        =   101
            Top             =   2130
            Width           =   1155
         End
         Begin VB.TextBox txtEmp 
            Alignment       =   2  'Center
            DataField       =   "Masculino5"
            BeginProperty DataFormat 
               Type            =   0
               Format          =   "#,##0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   8202
               SubFormatType   =   0
            EndProperty
            DataSource      =   "adoEmp(0)"
            Height          =   315
            Index           =   35
            Left            =   3195
            MaxLength       =   1
            TabIndex        =   106
            Top             =   2445
            Width           =   1155
         End
         Begin VB.TextBox txtEmp 
            Alignment       =   2  'Center
            DataField       =   "Masculino"
            BeginProperty DataFormat 
               Type            =   0
               Format          =   "#,##0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   8202
               SubFormatType   =   0
            EndProperty
            DataSource      =   "adoEmp(0)"
            Height          =   315
            Index           =   30
            Left            =   3195
            MaxLength       =   1
            TabIndex        =   81
            Top             =   870
            Width           =   1155
         End
         Begin VB.TextBox txtEmp 
            Alignment       =   1  'Right Justify
            DataField       =   "SSOCedula1"
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "#,##0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   8202
               SubFormatType   =   1
            EndProperty
            DataSource      =   "adoEmp(0)"
            Height          =   315
            Index           =   25
            Left            =   1785
            TabIndex        =   85
            Top             =   1185
            Width           =   1410
         End
         Begin VB.TextBox txtEmp 
            Alignment       =   1  'Right Justify
            DataField       =   "SSOCedula2"
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "#,##0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   8202
               SubFormatType   =   1
            EndProperty
            DataSource      =   "adoEmp(0)"
            Height          =   315
            Index           =   26
            Left            =   1785
            TabIndex        =   90
            Top             =   1500
            Width           =   1410
         End
         Begin VB.TextBox txtEmp 
            Alignment       =   1  'Right Justify
            DataField       =   "SSOCedula3"
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "#,##0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   8202
               SubFormatType   =   1
            EndProperty
            DataSource      =   "adoEmp(0)"
            Height          =   315
            Index           =   27
            Left            =   1785
            TabIndex        =   95
            Top             =   1815
            Width           =   1410
         End
         Begin VB.TextBox txtEmp 
            Alignment       =   1  'Right Justify
            DataField       =   "SSOCedula4"
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "#,##0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   8202
               SubFormatType   =   1
            EndProperty
            DataSource      =   "adoEmp(0)"
            Height          =   315
            Index           =   28
            Left            =   1785
            TabIndex        =   100
            Top             =   2130
            Width           =   1410
         End
         Begin VB.TextBox txtEmp 
            Alignment       =   1  'Right Justify
            DataField       =   "SSOCedula5"
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "#,##0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   8202
               SubFormatType   =   1
            EndProperty
            DataSource      =   "adoEmp(0)"
            Height          =   315
            Index           =   29
            Left            =   1785
            TabIndex        =   105
            Top             =   2445
            Width           =   1410
         End
         Begin VB.TextBox txtEmp 
            Alignment       =   1  'Right Justify
            DataField       =   "SSOCedula"
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "#,##0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   8202
               SubFormatType   =   1
            EndProperty
            DataSource      =   "adoEmp(0)"
            Height          =   315
            Index           =   24
            Left            =   1785
            TabIndex        =   80
            Top             =   870
            Width           =   1410
         End
         Begin VB.TextBox txtEmp 
            DataField       =   "Parentesco"
            DataSource      =   "adoEmp(0)"
            Height          =   315
            Index           =   18
            Left            =   375
            TabIndex        =   79
            Top             =   870
            Width           =   1410
         End
         Begin VB.TextBox txtEmp 
            DataField       =   "Parentesco5"
            DataSource      =   "adoEmp(0)"
            Height          =   315
            Index           =   23
            Left            =   375
            TabIndex        =   104
            Top             =   2445
            Width           =   1410
         End
         Begin VB.TextBox txtEmp 
            DataField       =   "Parentesco4"
            DataSource      =   "adoEmp(0)"
            Height          =   315
            Index           =   22
            Left            =   375
            TabIndex        =   99
            Top             =   2130
            Width           =   1410
         End
         Begin VB.TextBox txtEmp 
            DataField       =   "Parentesco3"
            DataSource      =   "adoEmp(0)"
            Height          =   315
            Index           =   21
            Left            =   375
            TabIndex        =   94
            Top             =   1815
            Width           =   1410
         End
         Begin VB.TextBox txtEmp 
            DataField       =   "Parentesco2"
            DataSource      =   "adoEmp(0)"
            Height          =   315
            Index           =   20
            Left            =   375
            TabIndex        =   89
            Top             =   1500
            Width           =   1410
         End
         Begin VB.TextBox txtEmp 
            DataField       =   "Parentesco1"
            DataSource      =   "adoEmp(0)"
            Height          =   315
            Index           =   19
            Left            =   375
            TabIndex        =   84
            Top             =   1185
            Width           =   1410
         End
         Begin MSMask.MaskEdBox mskemp 
            Bindings        =   "FrmFichaEmp.frx":0238
            DataField       =   "FechaNac1"
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "dd/MM/yyyy"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   8202
               SubFormatType   =   3
            EndProperty
            DataSource      =   "adoEmp(0)"
            Height          =   315
            Index           =   5
            Left            =   7650
            TabIndex        =   88
            TabStop         =   0   'False
            Top             =   1185
            Width           =   1395
            _ExtentX        =   2461
            _ExtentY        =   556
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   10
            Format          =   "dd/MM/yyyy"
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox mskemp 
            Bindings        =   "FrmFichaEmp.frx":025A
            DataField       =   "FechaNac2"
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "dd/MM/yyyy"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   8202
               SubFormatType   =   3
            EndProperty
            DataSource      =   "adoEmp(0)"
            Height          =   315
            Index           =   6
            Left            =   7650
            TabIndex        =   93
            TabStop         =   0   'False
            Top             =   1500
            Width           =   1395
            _ExtentX        =   2461
            _ExtentY        =   556
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   10
            Format          =   "dd/MM/yyyy"
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox mskemp 
            Bindings        =   "FrmFichaEmp.frx":027C
            DataField       =   "FechaNac3"
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "dd/MM/yyyy"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   8202
               SubFormatType   =   3
            EndProperty
            DataSource      =   "adoEmp(0)"
            Height          =   315
            Index           =   7
            Left            =   7650
            TabIndex        =   98
            TabStop         =   0   'False
            Top             =   1815
            Width           =   1395
            _ExtentX        =   2461
            _ExtentY        =   556
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   10
            Format          =   "dd/MM/yyyy"
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox mskemp 
            Bindings        =   "FrmFichaEmp.frx":029E
            DataField       =   "FechaNac4"
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "dd/MM/yyyy"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   8202
               SubFormatType   =   3
            EndProperty
            DataSource      =   "adoEmp(0)"
            Height          =   315
            Index           =   8
            Left            =   7650
            TabIndex        =   103
            TabStop         =   0   'False
            Top             =   2130
            Width           =   1395
            _ExtentX        =   2461
            _ExtentY        =   556
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   10
            Format          =   "dd/MM/yyyy"
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox mskemp 
            Bindings        =   "FrmFichaEmp.frx":02C0
            DataField       =   "FechaNac5"
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "dd/MM/yyyy"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   8202
               SubFormatType   =   3
            EndProperty
            DataSource      =   "adoEmp(0)"
            Height          =   315
            Index           =   9
            Left            =   7650
            TabIndex        =   108
            TabStop         =   0   'False
            Top             =   2445
            Width           =   1395
            _ExtentX        =   2461
            _ExtentY        =   556
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   10
            Format          =   "dd/MM/yyyy"
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox mskemp 
            Bindings        =   "FrmFichaEmp.frx":02E2
            DataField       =   "FechaNac"
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "dd/MM/yyyy"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   8202
               SubFormatType   =   3
            EndProperty
            DataSource      =   "adoEmp(0)"
            Height          =   315
            Index           =   4
            Left            =   7650
            TabIndex        =   83
            TabStop         =   0   'False
            Top             =   870
            Width           =   1395
            _ExtentX        =   2461
            _ExtentY        =   556
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   10
            Format          =   "dd/MM/yyyy"
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin VB.Label lblEmpleado 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Fecha de Nacimiento"
            ForeColor       =   &H80000008&
            Height          =   420
            Index           =   21
            Left            =   7650
            TabIndex        =   78
            Top             =   435
            Width           =   1410
         End
         Begin VB.Label lblEmpleado 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Apellidos y Nombres del Familiar"
            ForeColor       =   &H80000008&
            Height          =   420
            Index           =   20
            Left            =   4350
            TabIndex        =   77
            Top             =   435
            Width           =   3300
         End
         Begin VB.Label lblEmpleado 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Sexo (M / F)"
            ForeColor       =   &H80000008&
            Height          =   420
            Index           =   19
            Left            =   3195
            TabIndex        =   76
            Top             =   435
            Width           =   1155
         End
         Begin VB.Label lblEmpleado 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Cédula de Identidad"
            ForeColor       =   &H80000008&
            Height          =   420
            Index           =   18
            Left            =   1785
            TabIndex        =   75
            Top             =   435
            Width           =   1410
         End
         Begin VB.Label lblEmpleado 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Parentesto"
            ForeColor       =   &H80000008&
            Height          =   420
            Index           =   17
            Left            =   375
            TabIndex        =   74
            Top             =   435
            Width           =   1410
         End
      End
      Begin VB.Frame fraLis 
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
         Height          =   1785
         Index           =   0
         Left            =   -74655
         TabIndex        =   68
         Top             =   4740
         Width           =   4575
         Begin VB.CommandButton BotBusca 
            Height          =   330
            Left            =   3870
            Style           =   1  'Graphical
            TabIndex        =   113
            ToolTipText     =   "Buscar"
            Top             =   1305
            Width           =   480
         End
         Begin VB.TextBox txtEmp 
            BackColor       =   &H00FFFFFF&
            Height          =   315
            Index           =   15
            Left            =   195
            TabIndex        =   112
            Top             =   1335
            Width           =   3600
         End
         Begin VB.OptionButton OptBusca 
            Caption         =   "Cédula"
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
            Index           =   3
            Left            =   1470
            TabIndex        =   72
            Tag             =   "Cedula"
            Top             =   450
            Width           =   1065
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
            Height          =   225
            Index           =   2
            Left            =   1905
            TabIndex        =   71
            Tag             =   "Nombres"
            Top             =   885
            Width           =   1125
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
            Left            =   150
            TabIndex        =   70
            Tag             =   "CodEmp"
            Top             =   420
            Value           =   -1  'True
            Width           =   1575
         End
         Begin VB.OptionButton OptBusca 
            Caption         =   "Apellidos"
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
            Left            =   600
            TabIndex        =   69
            Tag             =   "Apellidos"
            Top             =   870
            Width           =   1335
         End
      End
      Begin VB.Frame fraLis 
         Caption         =   "Filtrar:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1785
         Index           =   1
         Left            =   -69765
         TabIndex        =   65
         Top             =   4740
         Width           =   4665
         Begin VB.TextBox txtEmp 
            BackColor       =   &H00FFFFFF&
            Height          =   315
            Index           =   43
            Left            =   1140
            TabIndex        =   116
            Top             =   420
            Width           =   1305
         End
         Begin VB.OptionButton OptBusca 
            Caption         =   "Activos"
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
            Index           =   6
            Left            =   3150
            TabIndex        =   115
            Tag             =   "CodEstado <> 1"
            Top             =   1290
            Width           =   1335
         End
         Begin VB.OptionButton OptBusca 
            Caption         =   "Todos"
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
            Index           =   5
            Left            =   2910
            TabIndex        =   114
            Tag             =   "0"
            Top             =   825
            Value           =   -1  'True
            Width           =   1575
         End
         Begin VB.TextBox Text2 
            BackColor       =   &H8000000F&
            BorderStyle     =   0  'None
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
            Left            =   165
            TabIndex        =   67
            Text            =   "Total Registros"
            Top             =   1245
            Width           =   1305
         End
         Begin VB.TextBox txtEmp 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            BorderStyle     =   0  'None
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
            Index           =   16
            Left            =   1470
            Locked          =   -1  'True
            TabIndex        =   66
            Top             =   1215
            Width           =   660
         End
         Begin VB.Label lblEmpleado 
            Caption         =   "Empleados:"
            Height          =   315
            Index           =   25
            Left            =   2850
            TabIndex        =   118
            Top             =   420
            Width           =   780
         End
         Begin VB.Label lblEmpleado 
            Caption         =   "Inmueble:"
            Height          =   315
            Index           =   24
            Left            =   210
            TabIndex        =   117
            Top             =   420
            Width           =   780
         End
      End
      Begin VB.Frame fraEmp 
         Enabled         =   0   'False
         Height          =   3300
         Index           =   2
         Left            =   -74670
         TabIndex        =   46
         Top             =   3255
         Width           =   9525
         Begin VB.TextBox txtEmp 
            DataField       =   "Concepto"
            DataSource      =   "adoEmp(0)"
            ForeColor       =   &H00404040&
            Height          =   630
            Index           =   45
            Left            =   1530
            MaxLength       =   150
            MultiLine       =   -1  'True
            TabIndex        =   126
            Top             =   2220
            Width           =   2940
         End
         Begin VB.TextBox txtEmp 
            Alignment       =   1  'Right Justify
            DataField       =   "BonoFijo"
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "#,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   8202
               SubFormatType   =   0
            EndProperty
            DataSource      =   "adoEmp(0)"
            Height          =   315
            Index           =   44
            Left            =   1545
            TabIndex        =   123
            Top             =   1740
            Width           =   2910
         End
         Begin VB.TextBox txtEmp 
            DataField       =   "Campo8"
            Height          =   315
            Index           =   14
            Left            =   6195
            TabIndex        =   54
            Top             =   2550
            Visible         =   0   'False
            Width           =   2565
         End
         Begin VB.TextBox txtEmp 
            DataField       =   "Campo7"
            Height          =   315
            Index           =   13
            Left            =   6195
            TabIndex        =   53
            Top             =   2115
            Visible         =   0   'False
            Width           =   2565
         End
         Begin VB.TextBox txtEmp 
            DataField       =   "Campo6"
            Height          =   315
            Index           =   12
            Left            =   6195
            TabIndex        =   52
            Top             =   1680
            Visible         =   0   'False
            Width           =   2565
         End
         Begin VB.TextBox TxtCampo5 
            DataField       =   "Campo5"
            Height          =   315
            Left            =   6195
            TabIndex        =   51
            Top             =   1245
            Visible         =   0   'False
            Width           =   2565
         End
         Begin VB.TextBox txtEmp 
            DataField       =   "Campo4"
            Height          =   315
            Index           =   11
            Left            =   6735
            TabIndex        =   50
            Top             =   2265
            Visible         =   0   'False
            Width           =   2565
         End
         Begin VB.TextBox txtEmp 
            DataField       =   "Campo3"
            Height          =   315
            Index           =   10
            Left            =   6735
            TabIndex        =   49
            Top             =   1830
            Visible         =   0   'False
            Width           =   2565
         End
         Begin VB.TextBox txtEmp 
            DataField       =   "Campo2"
            Height          =   315
            Index           =   9
            Left            =   6735
            TabIndex        =   48
            Top             =   1395
            Visible         =   0   'False
            Width           =   2565
         End
         Begin VB.TextBox TxtCampo1 
            DataField       =   "Campo1"
            Height          =   315
            Left            =   240
            TabIndex        =   47
            Top             =   705
            Width           =   8535
         End
         Begin VB.Label lblEmpleado 
            Caption         =   "Concepto:"
            Height          =   255
            Index           =   28
            Left            =   465
            TabIndex        =   125
            Top             =   2205
            Width           =   945
         End
         Begin VB.Label lblEmpleado 
            Caption         =   "Monto:"
            Height          =   255
            Index           =   27
            Left            =   465
            TabIndex        =   124
            Top             =   1770
            Width           =   945
         End
         Begin VB.Label Label1 
            Caption         =   "Si el empleado devenga una cantidad fija mensual, por otro concepto distinto al sueldo especifíquelo aquí:"
            Height          =   495
            Index           =   8
            Left            =   240
            TabIndex        =   122
            Top             =   1200
            Width           =   4455
         End
         Begin VB.Label Label1 
            Caption         =   "Si el cargo del empleado es conserje, señale brevemente las personas que vivirán en la conserjería"
            Height          =   285
            Index           =   1
            Left            =   225
            TabIndex        =   55
            Top             =   345
            Width           =   7785
         End
      End
      Begin VB.Frame fraEmp 
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
         Height          =   3285
         Index           =   1
         Left            =   315
         TabIndex        =   45
         Top             =   3255
         Width           =   9570
         Begin VB.CommandButton cmdFIngreso 
            Caption         =   "Carta Bco."
            Height          =   315
            Index           =   3
            Left            =   8385
            TabIndex        =   121
            Top             =   765
            Width           =   915
         End
         Begin MSComCtl2.MonthView Calendar 
            Height          =   2370
            Index           =   1
            Left            =   1755
            TabIndex        =   57
            Top             =   1650
            Visible         =   0   'False
            Width           =   2625
            _ExtentX        =   4630
            _ExtentY        =   4180
            _Version        =   393216
            ForeColor       =   -2147483630
            BackColor       =   -2147483633
            BorderStyle     =   1
            Appearance      =   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            StartOfWeek     =   60686337
            TitleBackColor  =   -2147483646
            TitleForeColor  =   -2147483639
            CurrentDate     =   36984
            MinDate         =   732
         End
         Begin VB.ComboBox cmbDiaL 
            DataField       =   "DiaLibre"
            DataSource      =   "adoEmp(0)"
            Height          =   315
            ItemData        =   "FrmFichaEmp.frx":0304
            Left            =   1635
            List            =   "FrmFichaEmp.frx":031D
            TabIndex        =   120
            Top             =   2055
            Width           =   3405
         End
         Begin VB.ComboBox cmb 
            DataField       =   "BonoNoc"
            DataSource      =   "adoEmp(0)"
            Height          =   315
            ItemData        =   "FrmFichaEmp.frx":035D
            Left            =   810
            List            =   "FrmFichaEmp.frx":0367
            TabIndex        =   109
            Top             =   1200
            Visible         =   0   'False
            Width           =   795
         End
         Begin VB.Frame fraDeucciones 
            Caption         =   "Ded&ucciones: "
            Height          =   780
            Left            =   5430
            TabIndex        =   39
            Top             =   2070
            Width           =   3885
            Begin VB.CheckBox chkDed 
               Caption         =   "S.&P.F"
               DataField       =   "SPF"
               DataSource      =   "adoEmp(0)"
               Height          =   285
               Index           =   2
               Left            =   2820
               TabIndex        =   42
               Top             =   315
               Width           =   900
            End
            Begin VB.CheckBox chkDed 
               Caption         =   "L.P.&H"
               DataField       =   "LPH"
               DataSource      =   "adoEmp(0)"
               Height          =   285
               Index           =   1
               Left            =   1590
               TabIndex        =   41
               Top             =   315
               Width           =   900
            End
            Begin VB.CheckBox chkDed 
               Caption         =   "S.S.&O"
               DataField       =   "SSO"
               DataSource      =   "adoEmp(0)"
               Height          =   285
               Index           =   0
               Left            =   360
               TabIndex        =   40
               Top             =   315
               Width           =   900
            End
         End
         Begin MSDataListLib.DataCombo dtcEmp 
            Bindings        =   "FrmFichaEmp.frx":0375
            DataField       =   "NombreCargo"
            Height          =   315
            Index           =   1
            Left            =   1635
            TabIndex        =   26
            Top             =   1200
            Width           =   3405
            _ExtentX        =   6006
            _ExtentY        =   556
            _Version        =   393216
            BackColor       =   -2147483643
            ListField       =   "NombreCargo"
            BoundColumn     =   ""
            Text            =   ""
         End
         Begin VB.CommandButton cmdFIngreso 
            Height          =   255
            Index           =   2
            Left            =   4725
            Picture         =   "FrmFichaEmp.frx":038D
            Style           =   1  'Graphical
            TabIndex        =   62
            Top             =   360
            Visible         =   0   'False
            Width           =   270
         End
         Begin VB.TextBox txtEmp 
            Alignment       =   1  'Right Justify
            DataField       =   "Sueldo"
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "#,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   8202
               SubFormatType   =   0
            EndProperty
            DataSource      =   "adoEmp(0)"
            Height          =   315
            Index           =   7
            Left            =   6360
            TabIndex        =   32
            Top             =   330
            Width           =   2910
         End
         Begin VB.TextBox txtEmp 
            DataField       =   "Cuenta"
            DataSource      =   "adoEmp(0)"
            Height          =   315
            Index           =   8
            Left            =   6375
            TabIndex        =   34
            Top             =   765
            Width           =   1995
         End
         Begin VB.CommandButton cmdFIngreso 
            Height          =   255
            Index           =   1
            Left            =   2730
            Picture         =   "FrmFichaEmp.frx":04D7
            Style           =   1  'Graphical
            TabIndex        =   56
            Top             =   360
            Width           =   270
         End
         Begin VB.TextBox txtEmp 
            DataField       =   "Comentario"
            DataSource      =   "adoEmp(0)"
            ForeColor       =   &H00404040&
            Height          =   315
            Index           =   6
            Left            =   1635
            MultiLine       =   -1  'True
            TabIndex        =   30
            Top             =   2505
            Width           =   3420
         End
         Begin MSDataListLib.DataCombo dtcEmp 
            Bindings        =   "FrmFichaEmp.frx":0621
            Height          =   315
            Index           =   3
            Left            =   6375
            TabIndex        =   36
            Top             =   1200
            Width           =   2910
            _ExtentX        =   5133
            _ExtentY        =   556
            _Version        =   393216
            Locked          =   -1  'True
            BackColor       =   -2147483643
            ListField       =   "NombreBanco"
            BoundColumn     =   "IDBanco"
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo dtcEmp 
            Bindings        =   "FrmFichaEmp.frx":0639
            DataField       =   "NombreContrato"
            Height          =   315
            Index           =   0
            Left            =   1635
            TabIndex        =   24
            Top             =   765
            Width           =   3405
            _ExtentX        =   6006
            _ExtentY        =   556
            _Version        =   393216
            BackColor       =   -2147483643
            ListField       =   "NombreContrato"
            BoundColumn     =   ""
            Text            =   ""
         End
         Begin MSMask.MaskEdBox mskemp 
            DataField       =   "Telefonos"
            BeginProperty DataFormat 
               Type            =   0
               Format          =   "0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   0
            EndProperty
            DataSource      =   "adoEmp(0)"
            Height          =   330
            Index           =   2
            Left            =   6375
            TabIndex        =   38
            Top             =   1627
            Width           =   2910
            _ExtentX        =   5133
            _ExtentY        =   582
            _Version        =   393216
            ClipMode        =   1
            PromptInclude   =   0   'False
            MaxLength       =   37
            Format          =   "(####) ###-####  /  (####) ###"
            Mask            =   "(####) ###-####  /  (####) ###-####"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox mskemp 
            Bindings        =   "FrmFichaEmp.frx":0651
            DataField       =   "FIngreso"
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "dd/MM/yyyy"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   8202
               SubFormatType   =   3
            EndProperty
            DataSource      =   "adoEmp(0)"
            Height          =   315
            Index           =   1
            Left            =   1635
            TabIndex        =   21
            Top             =   330
            Width           =   1395
            _ExtentX        =   2461
            _ExtentY        =   556
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   10
            Format          =   "dd/MM/yyyy"
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin MSDataListLib.DataCombo dtcEmp 
            Bindings        =   "FrmFichaEmp.frx":0673
            DataField       =   "CodGasto"
            DataSource      =   "adoEmp(0)"
            Height          =   315
            Index           =   2
            Left            =   1635
            TabIndex        =   28
            Top             =   1635
            Width           =   3405
            _ExtentX        =   6006
            _ExtentY        =   556
            _Version        =   393216
            BackColor       =   -2147483643
            ListField       =   "Cargado"
            BoundColumn     =   "CodGasto"
            Text            =   ""
         End
         Begin MSMask.MaskEdBox mskemp 
            Bindings        =   "FrmFichaEmp.frx":068A
            DataField       =   "FHasta"
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "dd/MM/yyyy"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   8202
               SubFormatType   =   3
            EndProperty
            DataSource      =   "adoEmp(0)"
            Height          =   315
            Index           =   3
            Left            =   3630
            TabIndex        =   22
            Top             =   330
            Visible         =   0   'False
            Width           =   1395
            _ExtentX        =   2461
            _ExtentY        =   556
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   10
            Format          =   "dd/MM/yyyy"
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin VB.Label lblEmpleado 
            Caption         =   "Dia Libre:"
            Height          =   250
            Index           =   26
            Left            =   255
            TabIndex        =   119
            Top             =   2102
            Width           =   1410
         End
         Begin VB.Label lblEmpleado 
            Alignment       =   1  'Right Justify
            Caption         =   "Sueldo:"
            Height          =   250
            Index           =   14
            Left            =   4905
            TabIndex        =   31
            Top             =   362
            Width           =   1320
         End
         Begin VB.Label lblEmpleado 
            Caption         =   "Código del Gasto:"
            Height          =   250
            Index           =   10
            Left            =   255
            TabIndex        =   27
            Top             =   1667
            Width           =   1410
         End
         Begin VB.Label lblEmpleado 
            Caption         =   "Fecha de Ingreso:"
            Height          =   250
            Index           =   3
            Left            =   255
            TabIndex        =   20
            Top             =   362
            Width           =   1410
         End
         Begin VB.Label lblEmpleado 
            Caption         =   "Comentario:"
            Height          =   250
            Index           =   6
            Left            =   255
            TabIndex        =   29
            Top             =   2537
            Width           =   1200
         End
         Begin VB.Label lblEmpleado 
            Caption         =   "Cargo:"
            Height          =   250
            Index           =   5
            Left            =   255
            TabIndex        =   25
            Top             =   1232
            Width           =   1410
         End
         Begin VB.Label lblEmpleado 
            Caption         =   "Tipo de Contrato:"
            Height          =   250
            Index           =   4
            Left            =   255
            TabIndex        =   23
            Top             =   797
            Width           =   1410
         End
         Begin VB.Label lblEmpleado 
            Alignment       =   1  'Right Justify
            Caption         =   "Teléfonos:"
            Height          =   250
            Index           =   2
            Left            =   5385
            TabIndex        =   37
            Top             =   1667
            Width           =   840
         End
         Begin VB.Label lblEmpleado 
            Alignment       =   1  'Right Justify
            Caption         =   "Banco:"
            Height          =   250
            Index           =   1
            Left            =   5385
            TabIndex        =   35
            Top             =   1232
            Width           =   840
         End
         Begin VB.Label lblEmpleado 
            Alignment       =   1  'Right Justify
            Caption         =   "Nº Cuenta:"
            Height          =   250
            Index           =   0
            Left            =   5385
            TabIndex        =   33
            Top             =   797
            Width           =   840
         End
      End
      Begin MSAdodcLib.Adodc adoEmp 
         Height          =   330
         Index           =   0
         Left            =   -72255
         Top             =   7020
         Visible         =   0   'False
         Width           =   2205
         _ExtentX        =   3889
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
         Caption         =   "AdoEmp"
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
      Begin MSAdodcLib.Adodc adoEmp 
         Height          =   330
         Index           =   1
         Left            =   -74580
         Top             =   7020
         Visible         =   0   'False
         Width           =   2265
         _ExtentX        =   3995
         _ExtentY        =   582
         ConnectMode     =   4
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
         Caption         =   "AdoContratos"
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
      Begin MSAdodcLib.Adodc adoEmp 
         Height          =   330
         Index           =   2
         Left            =   -69990
         Top             =   7020
         Visible         =   0   'False
         Width           =   2010
         _ExtentX        =   3545
         _ExtentY        =   582
         ConnectMode     =   4
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
         Caption         =   "AdoDpto"
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
      Begin MSAdodcLib.Adodc adoEmp 
         Height          =   330
         Index           =   3
         Left            =   -67920
         Top             =   7020
         Visible         =   0   'False
         Width           =   2115
         _ExtentX        =   3731
         _ExtentY        =   582
         ConnectMode     =   4
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
         Caption         =   "AdoBancos"
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
      Begin MSDataGridLib.DataGrid DataGrid1 
         Bindings        =   "FrmFichaEmp.frx":06AC
         Height          =   4095
         Left            =   -74850
         TabIndex        =   64
         Top             =   570
         Width           =   9900
         _ExtentX        =   17463
         _ExtentY        =   7223
         _Version        =   393216
         AllowUpdate     =   0   'False
         BackColor       =   16777215
         BorderStyle     =   0
         HeadLines       =   1
         RowHeight       =   17
         RowDividerStyle =   4
         FormatLocked    =   -1  'True
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Comic Sans MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
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
         Caption         =   "T R A B A J A D O R E S"
         ColumnCount     =   7
         BeginProperty Column00 
            DataField       =   "CodEmp"
            Caption         =   "CodEmp"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   "000000"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   8202
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   "CodInm"
            Caption         =   "Inm."
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
            DataField       =   "Cedula"
            Caption         =   "Cedula"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   "#,##0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   8202
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column03 
            DataField       =   "Apellidos"
            Caption         =   "Apellidos"
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
         BeginProperty Column04 
            DataField       =   "Nombres"
            Caption         =   "Nombres"
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
         BeginProperty Column05 
            DataField       =   "FIngreso"
            Caption         =   "Ingreso"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   "dd/MM/yyyy"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   8202
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column06 
            DataField       =   "Sueldo"
            Caption         =   "Sueldo"
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
         SplitCount      =   1
         BeginProperty Split0 
            BeginProperty Column00 
               ColumnWidth     =   720
            EndProperty
            BeginProperty Column01 
               Alignment       =   2
               ColumnWidth     =   750,047
            EndProperty
            BeginProperty Column02 
               Alignment       =   1
               ColumnWidth     =   1035,213
            EndProperty
            BeginProperty Column03 
               ColumnWidth     =   2055,118
            EndProperty
            BeginProperty Column04 
               ColumnWidth     =   1934,929
            EndProperty
            BeginProperty Column05 
               Alignment       =   2
               ColumnWidth     =   1305,071
            EndProperty
            BeginProperty Column06 
               Alignment       =   1
               ColumnWidth     =   1470,047
            EndProperty
         EndProperty
      End
   End
   Begin MSComCtl2.MonthView Calendar 
      Height          =   2370
      Index           =   0
      Left            =   7680
      TabIndex        =   61
      Top             =   2700
      Visible         =   0   'False
      Width           =   2625
      _ExtentX        =   4630
      _ExtentY        =   4180
      _Version        =   393216
      ForeColor       =   -2147483630
      BackColor       =   -2147483633
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      StartOfWeek     =   60686337
      TitleBackColor  =   -2147483646
      TitleForeColor  =   -2147483639
      CurrentDate     =   36984
      MinDate         =   732
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
            Picture         =   "FrmFichaEmp.frx":06C4
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmFichaEmp.frx":0846
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmFichaEmp.frx":09C8
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmFichaEmp.frx":0B4A
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmFichaEmp.frx":0CCC
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmFichaEmp.frx":0E4E
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmFichaEmp.frx":0FD0
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmFichaEmp.frx":1152
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmFichaEmp.frx":12D4
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmFichaEmp.frx":1456
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmFichaEmp.frx":15D8
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmFichaEmp.frx":175A
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "FrmFichaEmp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
    '---------------------------------------------------------------------------------------------
    '   Módulo de nómina - 15/11/2003
    '   Formulario ficha del trabajador
    ' variables globales a nivel de módulo
    Const Empleado$ = "000000"
    Dim rstGasto As New ADODB.Recordset
    Dim cnnGasto As New ADODB.Connection
    Dim Indice As Integer
    Public BN As Long
    Dim Valores_Ini()
    Private Enum EstatusEmpleado
        Activo
        Cesante
        Reposo
        Vacaciones
    End Enum
    
    '---------------------------------------------------------------------------------------------
    Private Sub adoEmp_MoveComplete(Index As Integer, ByVal adReason As ADODB.EventReasonEnum, _
    ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, _
    ByVal pRecordset As ADODB.Recordset)
    '
    If Index = 0 Then
    '
        If Not adoEmp(0).Recordset.EOF And Not adoEmp(0).Recordset.BOF Then
            If adoEmp(0).Recordset.EditMode <> adEditAdd Then
            '
            'If Not adoEmp(0).Recordset.BOF And Not adoEmp(0).Recordset.EOF Then
                '
                dtcEmp(0) = Contrato(adoEmp(0).Recordset("IDContrato"), True)
                dtcEmp(1) = Cargo(adoEmp(0).Recordset("CodCargo"), True)
                dtcEmp(3) = Banco(adoEmp(0).Recordset("CodInm"), True)
                Call dtcEmp_Click(1, 2)
                cmbStatus = cmbStatus.List(adoEmp(0).Recordset("CodEstado"))
                '
                If adoEmp(0).Recordset("CodEstado") = 1 Then
                    mskemp(3).Visible = True
                Else
                    mskemp(3).Visible = False
                End If
                '
            'End If
        '
            End If
        End If
    '
    End If
    '
    End Sub

    Private Sub BotBusca_Click()
    'variables locales
    Dim strCriterio$
    '
    On Error GoTo salir:
    With adoEmp(0).Recordset
        '
        If OptBusca(0) Then 'busqueda por codigo
            strCriterio = "CodEmp=" & txtEmp(15)
        ElseIf OptBusca(1) Then 'apellidos
            strCriterio = "Apellidos Like '%" & txtEmp(15) & "%'"
        ElseIf OptBusca(2) Then 'nombres
            strCriterio = "Nombres LIKE '%" & txtEmp(15) & "%'"
        ElseIf OptBusca(3) Then 'cedula
            strCriterio = "Cedula =" & txtEmp(15)
        End If
        '
        .Find strCriterio
        '
        If .EOF Then
            .MoveFirst
            .Find strCriterio
            '
            If .EOF Then
                MsgBox "No encontré la expresión seleccionada....", vbInformation, _
                App.ProductName
                .MoveFirst
                .Move .Bookmark
            End If
            '
        End If
        '
        '
    End With
salir:
    If Err.Number <> 0 Then MsgBox Err.Description, vbindormation, App.ProductName
    '
    End Sub
    
    Private Sub Calendar_DateClick(Index As Integer, ByVal DateClicked As Date)
    '
    Select Case Index
        Case 0
            mskemp(0) = DateClicked
            Calendar(0).Visible = False
        
        Case 1
            mskemp(Indice) = DateClicked
            Calendar(1).Visible = False
            
    End Select
    SendKeys vbTab
    End Sub

    Private Sub Calendar_KeyPress(Index As Integer, KeyAscii As Integer)
    '
    Select Case Index
        
        Case 0
            If KeyAscii = 13 Then
                mskemp(0) = Calendar(0).Value
                Calendar(0).Visible = False
            ElseIf KeyAscii = 27 Then
                Calendar(0).Visible = False
            End If
        
        Case 1
            If KeyAscii = 13 Then
                mskemp(Indice) = Calendar(1).Value
                Calendar(1).Visible = False
            ElseIf KeyAscii = 27 Then
                Calendar(1).Visible = False
            End If
        '
    End Select
    '
    End Sub

    Private Sub Calendar_LostFocus(Index As Integer)
    Calendar(0).Visible = False
    Calendar(1).Visible = False
    End Sub

    Private Sub ChkInactivo_KeyPress(Index As Integer, KeyAscii As Integer)
    If Index = 0 And KeyAscii = 13 Then mskemp(1).SetFocus
    End Sub

    Private Sub chkNac_Click(Index As Integer)
    Select Case Index
        Case 0: chkNac(1).Value = IIf(chkNac(0) = 0, 1, 0)
        Case 1: chkNac(0).Value = IIf(chkNac(1) = 0, 1, 0)
    End Select
    End Sub
    
    Private Sub cmb_KeyPress(KeyAscii As Integer): KeyAscii = 0
    End Sub




Private Sub cmbDiaL_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub cmbDiaL_LostFocus()
'VARIABLES LOCALES
Select Case cmbDiaL.Text
    Case "MIERCOLES", "MIECOLES", "MIERCOLE"
        cmbDiaL = "MIÉRCOLES"
    Case "SABADO", "SABADOS"
        cmbDiaL = "SÁBADO"
    Case "LUNE", "LUMES", "LNE"
        cmbDiaL = "LUNES"
End Select
'
End Sub

    Private Sub cmbSexo_KeyPress(KeyAscii As Integer)
    'variables locales
    Dim Item%
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If Chr(KeyAscii) = "F" Then
        Item = 1
    ElseIf Chr(KeyAscii) = "M" Then
        Item = 0
    End If
    KeyAscii = 0
    cmbSexo = cmbSexo.List(Item)
    End Sub


Private Sub cmbStatus_Click()
If cmbStatus = cmbStatus.List(1) Then
        mskemp(3).Visible = True
    Else
        mskemp(3).Visible = False
    End If
End Sub

    Private Sub cmbStatus_KeyPress(KeyAscii%): KeyAscii = 0
    End Sub

    Private Sub CmdFingreso_Click(Index As Integer)
    '
    Select Case Index
        Case 0
            Calendar(0).Visible = Not Calendar(0).Visible
            If Calendar(0).Visible Then Calendar(0).SetFocus
            '
        Case 1, 2
            Calendar(1).Left = IIf(Index = 1, 420, 2415)
            Calendar(1).Visible = Not Calendar(1).Visible
            If Calendar(1).Visible Then Calendar(1).SetFocus
            Indice = 3
            If Index = 1 Then Indice = 1
        Case 3
            MousePointer = vbHourglass
            cmdFIngreso(3).Enabled = Not cmdFIngreso(3).Enabled
            PrintCartaBanco
            cmdFIngreso(3).Enabled = Not cmdFIngreso(3).Enabled
            MousePointer = vbDefault
    End Select
    '
    End Sub
    
    
    Private Sub dtcEmp_Change(Index As Integer)
    '
    If Index = 0 Then
        If dtcEmp(0) = "CONTRATADO" Then
            mskemp(3).Visible = True
            cmdFIngreso(2).Visible = True
        Else
            mskemp(3).Visible = False
            cmdFIngreso(2).Visible = False
        End If
    End If
    '
    End Sub

    Private Sub dtcEmp_Click(Index As Integer, Area As Integer)
    '
    If Area = 2 Then
    
        dtcEmp(Index).Tag = dtcEmp(Index).BoundText
        If Index = 1 Then
            With adoEmp(2).Recordset
                .MoveFirst
                .Find "NombreCargo='" & dtcEmp(1) & "'"
                dtcEmp(1).Tag = .Fields("CodCargo")
                If Not .EOF Then
                    If .Fields("CodCargo") = BN Then
                        'muestra el control de bono_noc
                        cmb.Visible = True
                    Else
                        'oculta el control
                        cmb.Visible = False
                    End If
                End If
                '
            End With
        ElseIf Index = 0 Then
             With adoEmp(1).Recordset
                .MoveFirst
                .Find "NombreContrato='" & dtcEmp(0) & "'"
                If Not .EOF Then
                    dtcEmp(0).Tag = .Fields("IDContrato")
                End If
             End With
        End If
        '
    End If
    '
    End Sub

Private Sub dtcEmp_GotFocus(Index As Integer)
'variable local
Dim strInm As String
'
If adoEmp(0).Recordset.EditMode = adEditAdd Then
    strInm = txtEmp(0)
Else
    strInm = adoEmp(0).Recordset("CodInm")
End If
'
If strInm = "" Then Exit Sub
'
If Index = 2 Then

    cnnGasto.CursorLocation = adUseClient
    If cnnGasto.State = 1 Then cnnGasto.Close
    cnnGasto.Open cnnOLEDB & gcPath & "\" & strInm & "\inm.mdb"
    '
    rstGasto.Open "SELECT CodGasto & ' ' &  Titulo as Cargado,CodGasto FROM Tgastos WHERE Titul" _
    & "o LIKE 'SUELDO%' or Titulo LIKE 'SUPLENCIA%';", cnnGasto, adOpenStatic, adLockReadOnly, adCmdText
    Set dtcEmp(2).RowSource = rstGasto
    
End If
'
End Sub

    Private Sub dtcEmp_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If Index = 2 Then If KeyCode = 46 Then KeyCode = 0
    End Sub
    
    Private Sub dtcEmp_KeyPress(Index As Integer, KeyAscii As Integer)
    '
    Select Case Index
    '
        Case 1, 0
            If dtcEmp(Index) <> "" Then Call dtcEmp_Click(Index, 2)
        Case 2
            If KeyAscii = 13 Then
                Call dtcEmp_Click(Index, 2)
                
            Else
                KeyAscii = 0
            End If
            '
    End Select
    SendKeys vbTab
    '
    End Sub


    Private Sub Form_Resize()
    'variables locales
    On Error Resume Next
    'configura la presentacion de las fichas
    SSTab2.Top = Toolbar1.Height + 100
    SSTab2.Height = ScaleHeight - SSTab2.Top - 100
    SSTab2.Left = (ScaleWidth - SSTab2.Width) / 2
    fraEmp(0).Left = SSTab2.Left + 330
    fraEmp(1).Height = SSTab2.Height - fraEmp(1).Top - 200
    fraEmp(2).Height = SSTab2.Height - fraEmp(1).Top - 200
    fraEmp(3).Height = SSTab2.Height - fraEmp(1).Top - 200
    DataGrid1.Height = SSTab2.Height - DataGrid1.Top - fraLis(0).Height - 400
    fraLis(0).Top = DataGrid1.Top + DataGrid1.Height + 200
    fraLis(1).Top = fraLis(0).Top
    End Sub

    Private Sub Form_Unload(Cancel As Integer)
    '    On Error Resume Next
    If rstGasto.State = 1 Then rstGasto.Close
    Set rstGasto = Nothing
    If cnnGasto.State = 1 Then cnnGasto.Close
    Set cnnGasto = Nothing
    End Sub
    
    Private Sub LisSexo_Click(): mskemp(0).SetFocus
    End Sub
    
    Private Sub LisSexo_KeyPress(KeyAscii As Integer): mskemp(0).SetFocus
    End Sub
    
    Private Sub mskEmp_GotFocus(Index As Integer)
    On Error Resume Next
    With mskemp(Index)
        .SelStart = 0
        .SelLength = Len(.Text)
        .SetFocus
    End With
    End Sub

    Private Sub mskEmp_KeyPress(Index As Integer, KeyAscii As Integer)
    Call Validacion(KeyAscii, "0123456789")
    If KeyAscii = 13 Then SendKeys vbTab
    If Index = 0 And mskemp(0) <> "" And KeyAscii = 13 Then SendKeys vbTab
    End Sub
    
    Private Sub OptBusca_Click(Index As Integer)
'        Dim strSQL As String
'        For i = 0 To 3
'            If optBusca(i).Value = True Then strSQL = "' ORDER BY " & optBusca(i).Tag: Exit For
'        Next i
'        adoEmp(0).RecordSource = "SELECT * FROM EMP WHERE COdInm='" & gcCodInm & strSQL
'        adoEmp(0).Refresh
        '
    If Index > 4 Then
        If Index = 5 Then
        
            adoEmp(0).Recordset.Filter = 0
        Else
            adoEmp(0).Recordset.Filter = OptBusca(Index).Tag
        End If
        If txtEmp(43) <> "" Then
            If adoEmp(0).Recordset.Filter = 0 Then
                adoEmp(0).Recordset.Filter = "CodInm='" & txtEmp(43) & "'"
            Else
                adoEmp(0).Recordset.Filter = adoEmp(0).Recordset.Filter & " AND CodInm='" & txtEmp(43) & "'"
            End If
        End If
        txtEmp(16) = adoEmp(0).Recordset.RecordCount
    Else
        adoEmp(0).Recordset.Sort = OptBusca(Index).Tag
    End If
    End Sub

    Private Sub SSTab2_Click(PreviousTab As Integer)
    '
    Select Case SSTab2.tab
        Case 0, 1, 2: fraEmp(0).Visible = True
        Case 3: fraEmp(0).Visible = False
    End Select
    '
    End Sub


    Private Sub Form_Load()
    'variables locales
    Dim rstlocal As ADODB.Recordset
    Set rstlocal = New ADODB.Recordset
    '
    
    rstlocal.Open "Nom_Calc", cnnConexion, adOpenKeyset, adLockOptimistic, adCmdTable
    BN = rstlocal.Fields("bono_noc")
    rstlocal.Close
    Set rstlocal = Nothing
    '
    'configura control ADO
    
    'TABLA CONTRATOS
    adoEmp(1).ConnectionString = cnnOLEDB + gcPath + "\sac.mdb"
    adoEmp(1).RecordSource = "Select * from Emp_Contratos order by NombreContrato"
    adoEmp(1).Refresh
    '
    'TABLA CARGOS
    adoEmp(2).ConnectionString = cnnOLEDB + gcPath + "\sac.mdb"
    adoEmp(2).RecordSource = "SELECT * FROM Emp_Cargos ORDER BY NombreCargo"
    adoEmp(2).Refresh
    '
    'tabla empleados
    adoEmp(0).ConnectionString = cnnOLEDB + gcPath + "\sac.mdb"
    adoEmp(0).CursorLocation = adUseClient
    adoEmp(0).CommandType = adCmdTable
    adoEmp(0).RecordSource = "Emp"
    'adoEmp(0).RecordSource = "SELECT Emp_Cargos.NombreCargo, Emp.*, Emp_contratos.NombreContrat" _
    & "o FROM Emp_contratos RIGHT JOIN (Emp_Cargos RIGHT JOIN Emp ON Emp_Cargos.CodCargo = Emp." _
    & "CodCargo) ON Emp_contratos.IDContrato = Emp.IDContrato WHERE Emp.CodInm='" & gcCodInm & "'"
    adoEmp(0).Refresh
    adoEmp(0).Recordset.Sort = "CodInm, CodEmp"
    '--------------------------------------------
    'configura conexion
'    cnnGasto.CursorLocation = adUseClient
'    cnnGasto.Open cnnOLEDB & gcPath & "\" & adoEmp(0).Recordset("CodInm") & "\inm.mdb"
'
'    rstGasto.Open "SELECT CodGasto & ' ' &  Titulo as Cargado,CodGasto FROM Tgastos WHERE Titul" _
'    & "o LIKE 'SUELDO%';", cnnGasto, adOpenStatic, adLockReadOnly, adCmdText
'    Set dtcEmp(2).RowSource = rstGasto
    BotBusca.Picture = LoadResPicture("Buscar", vbResIcon)
    '
    txtEmp(16) = adoEmp(0).Recordset.RecordCount
    SSTab2.tab = 0
    If Not adoEmp(0).Recordset.EOF And Not adoEmp(0).Recordset.BOF Then _
    adoEmp(0).Recordset.MoveFirst
    '
    Call RtnEstado(6, Toolbar1, adoEmp(0).Recordset.EOF Or adoEmp(0).Recordset.BOF)
    ReDim Valores_Ini(adoEmp(0).Recordset.Fields.count - 1)
    '
    '
    End Sub

    
    Private Sub Toolbar1_ButtonClick(ByVal Button As Button)
    'variables actuales
    Dim Resp As Long
    Dim blnNuevo As Boolean
    Dim sMensaje As String
    '
    With adoEmp(0).Recordset
    '
        Select Case Button.Key
        
            ' Primer Registro
            Case Is = "First"
                .MoveFirst
                
            ' Registro Anterior
            Case Is = "Previous"
                .MovePrevious
                If .BOF Then .MoveLast
                
            ' Registro Siguiente
            Case Is = "Next"
                .MoveNext
                If .EOF Then .MoveFirst
                
            ' Ultimo Registro
            Case Is = "End"
                .MoveLast
                
            ' Nuevo FrmEmpleados.
            Case Is = "New"
            
                SSTab2.tab = 0
                Call Frames
                .AddNew
                chkNac(0).Value = vbChecked
                For I = 0 To 2: chkDed(I).Value = 0
                Next I
                Call RtnEstado(Button.Index, Toolbar1)
                dtcEmp(0) = ""
                dtcEmp(1) = ""
                dtcEmp(3) = ""
                cmbStatus = ""
                txtEmp(0).SetFocus
                
            ' Actualizar.
            Case Is = "Save"
                If dtcEmp(3) = "" Then
                    MsgBox "Falta el campo relacionado al código del banco", vbInformation, _
                    App.ProductName
                    Exit Sub
                End If
                Call EditBox
                
                If validar_registro Then Call EditBox: Exit Sub
                Call Frames
                '.Fields("CodInm") = gcCodInm
                .Fields("FecAct") = Date
                .Fields("Usuario") = gcUsuario
                .Fields("IDContrato") = IIf(dtcEmp(0).Tag = "", .Fields("IDContrato"), dtcEmp(0).Tag)
                .Fields("CodCargo") = dtcEmp(1).Tag
                .Fields("CodGasto") = IIf(dtcEmp(2).Tag = "", .Fields("CodGasto"), dtcEmp(2).Tag)
                .Fields("Banco") = dtcEmp(3).Tag
                If cmb <> "15" And cmb <> "30" Then
                If cmb.ListIndex = 1 Then
                    .Fields("BonoNoc") = 15
                ElseIf cmb.ListIndex = 0 Then
                    .Fields("BonoNoc") = 30
                Else
                    .Fields("BonoNoc") = 0
                End If
                End If
'                If cmb.Visible Then
'                    .Fields("BonoNoc") = IIf(cmb Like "30*", 30, IIf(cmb Like "15*", 15, 0))
'                Else
'                    .Fields("BonoNoc") = 0
'                End If
                '
                If adoEmp(0).Recordset.EditMode = adEditAdd Then
                    txtEmp(42) = Format(NEmpleado(txtEmp(0)), Empleado)
                    .Fields("CodEstado") = 0
                    blnNuevo = False
                Else
                    blnNuevo = True
                    If cmbStatus.ListIndex < 0 Then
                        For I = 0 To cmb.ListCount
                            If cmbStatus.Text = cmbStatus.List(I) Then
                                .Fields("CodEstado") = I
                                Exit For
                            End If
                        Next
                    Else
                        .Fields("CodEstado") = cmbStatus.ListIndex
                    End If
                    
                    'busca los cambios efectuados en la fila
                    For I = 0 To .Fields.count - 1
                        If .Fields(I) <> Valores_Ini(I) And .Fields(I).Name <> "Usuario" And .Fields(I).Name <> "FecAct" Then
                            Call rtnBitacora("Actulizado el campo " & .Fields(I).Name & " Valor" _
                            & " Inicial: " & Valores_Ini(I) & " Valor Actual: " & .Fields(I))
                        End If
                    Next
                    Dim Respuesta As Long
'                    If Valores_Ini(.Fields("CodEstado").Properties.Item) = Vacaciones And .Fields("CodEstad0") = Activo Then
'                        'la persona la están activando en forma manual
'                        respuesa = MsgBox("Está activando un empleado que se encuentra de " & _
'                        "vacaciones" & vbCrLf & "y que no le corresponde reintegrarse aún." & _
'                        "¿Desea eliminar la orden de pago por concepto de vacaciones?", _
'                        vbQuestion + vbYesNo, App.ProductName)
'                        '
'                        'FALTA CODIGO
'
'                    End If
                End If
                If Not IsNumeric(txtEmp(44)) Then txtEmp(44) = "0,00"
                .Update
                '.Requery
                If Not blnNuevo Then
                    sMensaje = "¿Desea emitir la Solicitud Apertura Cuenta Nómina?" + _
                    vbCrLf + "Recuerde que ud. podrá emitirla en otro momento"
                    Resp = MsgBox(sMensaje, vbQuestion + vbYesNo, App.ProductName)
                    If Resp = vbYes Then PrintCartaBanco
                End If
                Call rtnBitacora("Registro emp. " & !CodEmp & " actualizado")
                Call RtnEstado(Button.Index, Toolbar1)
                Call EditBox
                frmNomina.blnModif = MsgBox("Registro Actualizado ... ", vbInformation, App.ProductName)
                
            ' Buscar.
            Case Is = "Find"
                '
                SSTab2.tab = 2
                TxtBus.SetFocus
                
            ' Cencelar
            Case Is = "Undo"
                '
                Set DataGrid1.DataSource = Nothing
                Call EditBox
                .CancelUpdate
                Call rtnBitacora("Edición de ficha emp. " & !CodEmp & " cancelada por el usuario")
                Call EditBox
                Set DataGrid1.DataSource = adoEmp(0)
                Call Frames
                Call RtnEstado(Button.Index, Toolbar1)
            
            ' Eliminar Registro
            Case Is = "Delete"
                Resp = MsgBox("Está seguro de eliminar de nuestro registro a '" & !Apellidos & _
                ", " & !Nombres & "'", vbQuestion + vbYesNo, App.ProductName)
                If Resp = vbYes Then
                    .Delete
                    MsgBox "Registro eliminado...", vbInformation, App.ProductName
                    Call rtnBitacora("Empleado " & !CodEmp & " eliminado...")
                End If
                Call RtnEstado(Button.Index, Toolbar1)
            
            'Editar Registro
            Case Is = "Edit1"
            '
                Call Frames
                Call RtnEstado(Button.Index, Toolbar1)
                'almacena los valores iniciales
                For I = 0 To .Fields.count - 1: Valores_Ini(I) = .Fields(I)
                Next
                Call rtnBitacora("Editar ficha empleado " & .Fields("CodEmp"))
                
            ' Cerrar y Salir
            Case Is = "Close": Unload FrmFichaEmp: Set FrmFichaEmp = Nothing
            
            ' Imprimir
            Case Is = "Print"
                '
                '
                mcTitulo = "Ficha de Empleados"
                mcReport = "Emp_list.Rpt"
                If txtEmp(43) <> "" Then
                    mcCrit = "{Emp.CodInm}='" & txtEmp(43) & "'"
                Else
                    mcCrit = ""
                End If
                If OptBusca(6) Then
                    mcCrit = IIf(mcCrit <> "", mcCrit & " AND ", "") & "{Emp.CodEstado}<>" & Cesante
                End If
                'If OptBusca(6) Then mcCrit = mcCrit & IIf(mcCrit <> "", " AND ", "") & _
                "{Emp.CodEstado}=0"
                FrmReport.Show
           
        End Select
        '
    End With
    '
    End Sub
    
    
    Private Sub EditBox()
    '
    For I = mskemp.LBound To mskemp.UBound: mskemp(I).PromptInclude = Not mskemp(I).PromptInclude
    Next I
    '
    End Sub
    
    Private Sub Frames()
    '
    For I = fraEmp.LBound To fraEmp.UBound: fraEmp(I).Enabled = Not fraEmp(I).Enabled
    Next I
    '
    End Sub
    
    '-------------------------------------------------------------------------------------------------
    '   Funcion:    NEmpleado
    '
    '   Devuelde el código de empleado correlativo para un inmueble determinado
    '-------------------------------------------------------------------------------------------------
    Private Function NEmpleado(Inm$) As Long
    'variables locales
    Dim objRst As New ADODB.Recordset
    '
    objRst.Open "SELECT Max(CodEmp) + 1 as N FROM Emp WHERE CodInm='" & Inm & "';", _
    cnnConexion, adOpenStatic, adLockReadOnly, adCmdText
    '
    If Not IsNull(objRst!N) Then
        NEmpleado = objRst!N
    Else: NEmpleado = Right(Inm, 2) & "0001"
    End If
    '
    objRst.Close
    Set objRst = Nothing
    '
    End Function
    
    
    Private Sub txtEmp_KeyPress(Index As Integer, KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then
        If Index = 43 Then
            If txtEmp(43) = "" Then
                adoEmp(0).Recordset.Filter = 0
            Else
                adoEmp(0).Recordset.Filter = "Codinm ='" & txtEmp(43) & "'"
            End If
        ElseIf Index = 45 Then
            KeyAscii = 0
        End If
        SendKeys vbTab
        Exit Sub
    End If
    '
    Select Case Index
        Case 0, 3, 43
            Call Validacion(KeyAscii, "1234567890")
        Case 7, 44
            If KeyAscii = Asc(".") Then KeyAscii = Asc(",")
            Call Validacion(KeyAscii, "0123456789,")
    
    End Select
    '
    End Sub
    
    Private Sub txtEmp_LostFocus(Index As Integer)
    'variables locales
    Select Case Index
    
        Case 7: txtEmp(7) = Format(txtEmp(7), "#,##0.00")
        
        Case 0
            
            If adoEmp(0).Recordset.EditMode = adEditAdd Then
            '
                If txtEmp(0) = "" Then
                    MsgBox "Debe introducir el código del inmueble para continuar", vbInformation, _
                    App.ProductName
                    txtEmp(0).SetFocus
                Else
                
                    txtEmp(42) = Format(NEmpleado(txtEmp(0)), Empleado)
                    dtcEmp(3) = Banco(txtEmp(0), True)
                    
'                    cnnGasto.CursorLocation = adUseClient
'                    cnnGasto.Open cnnOLEDB & gcPath & "\" & txt(0) & "\inm.mdb"
'
'                    rstGasto.Open "SELECT CodGasto & ' ' &  Titulo as Cargado,CodGasto FROM Tgastos WHERE Titul" _
'                    & "o LIKE 'SUELDO%';", cnnGasto, adOpenStatic, adLockReadOnly, adCmdText
'                    dtcEmp(2).Refresh
                    
                End If
                '
            End If
            '
    End Select
    '
    End Sub
    
    '---------------------------------------------------------------------------------------------
    '   Función:    validar_registro
    '
    '   Devuelve un valor "TRUE" si falta algún parámetro req. para guardar
    '   un registro
    '---------------------------------------------------------------------------------------------
    Private Function validar_registro() As Boolean
    'variables locales
    Dim strFalta As String
    Dim vecDia(6) As String
    vecDia(0) = "LUNES"
    vecDia(1) = "MARTES"
    vecDia(2) = "MIÉRCOLES"
    vecDia(3) = "JUEVES"
    vecDia(4) = "VIERNES"
    vecDia(5) = "SÁBADO"
    vecDia(6) = "DOMINGO"
    Dim bln As Boolean
    '
    If txtEmp(42) = "" Then strFalta = "- Código Empleado"
    If txtEmp(1) = "" Then strFalta = IIf(strFalta = "", "", strFalta & ", ") & "- Apellido(s)"
    If txtEmp(2) = "" Then strFalta = IIf(strFalta = "", "", strFalta & ", ") & "- Nombre(s)"
    If txtEmp(3) = "" Then strFalta = IIf(strFalta = "", "", strFalta & ", ") & "- Cédula de Identidad"
    If Not IsDate(mskemp(0)) Then strFalta = IIf(strFalta = "", "", strFalta & ", ") & "- Fecha de " _
    & "Nacimiento "
    If Not IsDate(mskemp(1)) Then strFalta = IIf(strFalta = "", "", strFalta & ", ") & "- Fecha de " _
    & "Ingreso"
    If dtcEmp(0) = "" Then strFalta = IIf(strFalta = "", "", strFalta & ", ") & "- Tipo de Contrato"
    If dtcEmp(1) = "" Then strFalta = IIf(strFalta = "", "", strFalta & ", ") & "- Cargo"
    If dtcEmp(2) = "" Then strFalta = IIf(strFalta = "", "", strFalta & ", ") & "- Código Gasto Nómina"
    If dtcEmp(3) = "" Then strFalta = IIf(strFalta = "", "", strFalta & ", ") & "- Banco"
    If txtEmp(7) = "" Then strFalta = IIf(strFalta = "", "", strFalta & ", ") & "- Sueldo"
    If dtcEmp(0) = "CONTRATADO" And Not IsDate(mskemp(3)) Then
        strFalta = IIf(strFalta = "", "", strFalta & ", ") & "- Fecha fin del Contrato"
    End If
    'If dtcEmp(0).Tag = "" Then strFalta = IIf(strFalta = "", "", strFalta & ", ") & "ID. Contrato"
    If dtcEmp(1).Tag = "" Then strFalta = IIf(strFalta = "", "", strFalta & ", ") & "- ID Cargo"
    'If dtcEmp(2).Tag = "" Then strFalta = IIf(strFalta = "", "", strFalta & ", ") & "Cod. Gasto"
    If dtcEmp(3).Tag = "" Then strFalta = IIf(strFalta = "", "", strFalta & ", ") & "- ID Banco"
    
    If cmbStatus.ListIndex = 1 Then 'empleado cesante
        If Not IsDate(mskemp(3)) Then strFalta = IIf(strFalta = "", "", strFalta & ", ") _
        & "- Fecha de Cesantía."
    End If
    For I = 0 To 6
        If cmbDiaL = vecDia(I) Then bln = True: Exit For
    Next
    If Not bln Then strFalta = IIf(strFalta = "", "", strFalta & ", ") _
        & "- Día Libre del trabajador."
        
    If Not strFalta = "" Then validar_registro = MsgBox("Falta Información:" & vbCrLf & vbCrLf & strFalta, _
    vbInformation, App.ProductName)
    '
    End Function

    '---------------------------------------------------------------------------------------------
    '   Funcion:    Busca_banco
    '
    '   Entrada:    Nombre del Banco
    '
    '   Salida:     ID del banco
    '---------------------------------------------------------------------------------------------
    Function Busca_Banco(Nombre$) As Long
    '
    With adoEmp(3).Recordset
    
        .MoveFirst
        .Find "NombreBanco='" & Nombre & "'"
        Busca_Banco = .Fields("IDBanco")
        
    End With
    '
    End Function

    '---------------------------------------------------------------------------------------------
    '   Funcion:    Contrato
    '
    '   Entrada:    Código del contrato
    '
    '   Devuelve una variable cadena que contiene el nombre del contrato
    '---------------------------------------------------------------------------------------------
    Function Contrato(Ctto&, Optional Retorna As Boolean) As String
    '
    With adoEmp(1).Recordset
        .MoveFirst
        .Find "IDContrato=" & Ctto
        If Not .EOF Then If Retorna Then Contrato = !NombreContrato
        If Retorna Then dtcEmp(0).Tag = !IDContrato
    End With
    '
    End Function

    '---------------------------------------------------------------------------------------------
    '   Funcion:    Cargo
    '
    '   Entrada:    Código del cargo
    '
    '   Devuelve una variable cadena que contiene el nombre del contrato
    '---------------------------------------------------------------------------------------------
    Function Cargo(CodCargo&, Optional Retorna As Boolean) As String
    '
    With adoEmp(2).Recordset
        .MoveFirst
        .Find "CodCargo=" & CodCargo
        If Not .EOF Then
            If Retorna Then Cargo = !NombreCargo
            dtcEmp(1).Tag = CodCargo
        End If
    End With
    '
    End Function
    
    
    '---------------------------------------------------------------------------------------------
    '   Funcion:    Banco
    '
    '   Entrada:    Código del Inmueble
    '
    '   Devuelve una variable cadena que contiene el nombre del banco de la cta de nomina
    '---------------------------------------------------------------------------------------------
    Function Banco(Inm$, Optional Retorna As Boolean) As String
    'variables locales
    Dim RstBco As New ADODB.Recordset
    '
    RstBco.Open "SELECT * FROM Inmueble WHERE CodInm='" & Inm & "'", cnnConexion, adOpenKeyset, _
    adLockOptimistic, adCmdText
    
    If Not RstBco.EOF And Not RstBco.BOF Then
        'TABLA BANCOS
        If RstBco("Caja") = sysCodCaja Then
            adoEmp(3).ConnectionString = cnnOLEDB + gcPath & "\" & sysCodInm & "\inm.mdb"
        Else
            adoEmp(3).ConnectionString = cnnOLEDB + gcPath & "\" & Inm & "\inm.mdb"
        End If
    
        adoEmp(3).RecordSource = "SELECT * FROM BANCOS WHERE IDBANCO =(SELECT IDBANCO FROM CUEN" _
        & "TAS WHERE IDCuenta=(SELECT CtaInm FROM Inmueble in '" & gcPath & "\sac.mdb' WHERE CO" _
        & "dInm='" & Inm & "'))"
        adoEmp(3).Refresh
        '
        If adoEmp(3).Recordset.EOF Or adoEmp(3).Recordset.BOF Then
            'MsgBox "Falta información de la cuenta de nómina", vbInformation, App.ProductName
            Banco = "CUENTA DE NOMINA NO REGISTRADA"
        Else
            If Retorna = True Then Banco = adoEmp(3).Recordset("NombreBanco")
            dtcEmp(3).Tag = adoEmp(3).Recordset("IDBanco")
        End If
        '
    Else
        Banco = "FALTA INF."
    End If
    'Cierra y descarga el objeto
    RstBco.Close
    Set RstBco = Nothing
    '
    End Function

    Private Sub PrintCartaBanco()
    Dim Archivo1 As String
    Dim Archivo2 As String
    Dim rpReporte As ctlReport
    Dim Cuenta As String
    Dim rstlocal As ADODB.Recordset
    
    Archivo1 = gcPath & "\sac.mdb"
    
    With FrmAdmin.objRst
        If Not (.EOF And .BOF) Then
            .MoveFirst
            .Find "CodInm='" & adoEmp(0).Recordset!CodInm & "'"
            If Not .EOF Then
                Set rstlocal = New ADODB.Recordset
                If !Caja = sysCodCaja Then
                    Archivo2 = gcPath & "\9999\inm.mdb"
                Else
                    Archivo2 = gcPath & "\" & !CodInm & "\inm.mdb"
                End If
                rstlocal.CursorLocation = adUseClient
                rstlocal.Open "Cuentas", cnnOLEDB + Archivo2, adOpenKeyset, adLockOptimistic, adCmdTable
                rstlocal.Filter = "IDCuenta=" & !CtaInm
                If Not .EOF Then
                    Cuenta = rstlocal!NumCuenta
                Else
                    MsgBox "No se encuentra la información relacionada con la cuenta del inmueble.", vbInformation, App.ProductName
                End If
                rstlocal.Close
                Set rstlocal = Nothing
            End If
            .MoveFirst
        Else
            MsgBox "No se puede imprimir este reporte." & vbCrLf & _
            "Consulte al administrador del sistema.", vbCritical, App.ProductName
            Exit Sub
        End If
    End With
    Set rpReporte = New ctlReport
    With rpReporte
        .Reporte = gcReport & "nom_apecta.rpt"
        .TituloVentana = "Carta Apertura Cuenta"
        .OrigenDatos(0) = Archivo1
        .OrigenDatos(1) = Archivo1
        .OrigenDatos(2) = Archivo1
        .Formulas(0) = "sCta='" & Cuenta & "'"
        .FormuladeSeleccion = "{Emp.CodEmp}=" & adoEmp(0).Recordset!CodEmp
        .Salida = crPantalla
        .Imprimir
        Call rtnBitacora("Imprimiendo Sol.Apertura Cta Emp " & adoEmp(0).Recordset!CodEmp)
    End With
    Set rpReporte = Nothing
    End Sub
