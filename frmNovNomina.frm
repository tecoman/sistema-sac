VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmNovNomina 
   Caption         =   "Novedades de Nómina"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   WindowState     =   2  'Maximized
   Begin ComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   480
      Left            =   0
      TabIndex        =   9
      Top             =   0
      Width           =   4680
      _ExtentX        =   8255
      _ExtentY        =   847
      ButtonWidth     =   714
      ButtonHeight    =   688
      AllowCustomize  =   0   'False
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
            Object.Visible         =   0   'False
            Key             =   "Delete"
            Object.ToolTipText     =   "Eliminar Registro"
            Object.Tag             =   ""
            ImageIndex      =   9
         EndProperty
         BeginProperty Button10 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Enabled         =   0   'False
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
      MouseIcon       =   "frmNovNomina.frx":0000
   End
   Begin VB.Frame fraNom 
      Enabled         =   0   'False
      Height          =   3675
      Index           =   2
      Left            =   180
      TabIndex        =   10
      Top             =   4545
      Width           =   11385
      Begin VB.TextBox TXT 
         Alignment       =   1  'Right Justify
         DataField       =   "diaslib"
         Height          =   315
         Index           =   0
         Left            =   1395
         TabIndex        =   47
         Text            =   "0"
         Top             =   1620
         Width           =   765
      End
      Begin VB.TextBox TXT 
         Alignment       =   1  'Right Justify
         DataField       =   "Nombres"
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "#,##0"
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
         Index           =   13
         Left            =   4140
         Locked          =   -1  'True
         TabIndex        =   36
         Top             =   315
         Width           =   2490
      End
      Begin VB.TextBox TXT 
         Alignment       =   1  'Right Justify
         DataField       =   "Apellidos"
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "#,##0"
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
         Left            =   1815
         Locked          =   -1  'True
         TabIndex        =   35
         Top             =   315
         Width           =   2295
      End
      Begin VB.TextBox TXT 
         Alignment       =   1  'Right Justify
         DataField       =   "OtrasDed"
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
         Left            =   9375
         TabIndex        =   32
         Text            =   "0,00"
         Top             =   2460
         Width           =   1680
      End
      Begin VB.TextBox TXT 
         Alignment       =   1  'Right Justify
         DataField       =   "Descuentos"
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
         Left            =   6135
         TabIndex        =   30
         Text            =   "0,00"
         Top             =   2460
         Width           =   1680
      End
      Begin VB.TextBox TXT 
         Alignment       =   1  'Right Justify
         DataField       =   "DiasNTrab"
         Height          =   315
         Index           =   2
         Left            =   3360
         TabIndex        =   28
         Text            =   "0"
         Top             =   2475
         Width           =   765
      End
      Begin VB.TextBox TXT 
         Alignment       =   1  'Right Justify
         DataField       =   "OtrosBonos"
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
         Left            =   9405
         TabIndex        =   26
         Text            =   "0,00"
         Top             =   1635
         Width           =   1680
      End
      Begin VB.TextBox TXT 
         DataField       =   "Descripcion1"
         Height          =   315
         Index           =   7
         Left            =   1440
         MaxLength       =   240
         TabIndex        =   22
         Top             =   3045
         Width           =   9600
      End
      Begin VB.TextBox TXT 
         Alignment       =   1  'Right Justify
         DataField       =   "OtrasAsig"
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
         Left            =   6135
         TabIndex        =   21
         Text            =   "0,00"
         Top             =   1620
         Width           =   1680
      End
      Begin VB.TextBox TXT 
         Alignment       =   1  'Right Justify
         DataField       =   "diasfer"
         Height          =   315
         Index           =   1
         Left            =   3570
         TabIndex        =   24
         Text            =   "0"
         Top             =   1635
         Width           =   765
      End
      Begin VB.TextBox TXT 
         DataField       =   "NombreCargo"
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
         Index           =   14
         Left            =   7410
         Locked          =   -1  'True
         TabIndex        =   14
         Top             =   315
         Width           =   3630
      End
      Begin VB.TextBox TXT 
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
         Left            =   4155
         Locked          =   -1  'True
         TabIndex        =   13
         Top             =   750
         Width           =   1150
      End
      Begin VB.TextBox TXT 
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
         Left            =   1815
         Locked          =   -1  'True
         TabIndex        =   12
         Top             =   750
         Width           =   1150
      End
      Begin VB.TextBox TXT 
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
         Left            =   7425
         Locked          =   -1  'True
         TabIndex        =   11
         Text            =   "0,00"
         Top             =   750
         Width           =   1680
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         Caption         =   "Descripción:"
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
         Index           =   20
         Left            =   210
         TabIndex        =   49
         Top             =   3075
         Width           =   1110
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         Caption         =   "Dias &Libres:"
         Height          =   285
         Index           =   0
         Left            =   450
         TabIndex        =   48
         Top             =   1665
         Width           =   945
      End
      Begin VB.Label lbl 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "N"
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
         Index           =   15
         Left            =   9210
         TabIndex        =   38
         Top             =   765
         Width           =   1800
      End
      Begin VB.Line lin 
         BorderColor     =   &H00FFFFFF&
         Index           =   3
         X1              =   1530
         X2              =   11000
         Y1              =   2250
         Y2              =   2250
      End
      Begin VB.Line lin 
         BorderColor     =   &H00808080&
         BorderWidth     =   2
         Index           =   2
         X1              =   1530
         X2              =   11000
         Y1              =   2250
         Y2              =   2235
      End
      Begin VB.Label lbl 
         Caption         =   "Deducciones:"
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
         Index           =   14
         Left            =   210
         TabIndex        =   34
         Top             =   2115
         Width           =   1170
      End
      Begin VB.Label lbl 
         Caption         =   "Asignaciones:"
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
         Index           =   13
         Left            =   240
         TabIndex        =   33
         Top             =   1245
         Width           =   1170
      End
      Begin VB.Line lin 
         BorderColor     =   &H00FFFFFF&
         Index           =   1
         X1              =   1575
         X2              =   11000
         Y1              =   1380
         Y2              =   1380
      End
      Begin VB.Line lin 
         BorderColor     =   &H00808080&
         BorderWidth     =   2
         Index           =   0
         X1              =   1560
         X2              =   11000
         Y1              =   1380
         Y2              =   1365
      End
      Begin VB.Label lbl 
         Caption         =   "&Otras Deducciones:"
         Height          =   285
         Index           =   12
         Left            =   7890
         TabIndex        =   31
         Top             =   2520
         Width           =   1575
      End
      Begin VB.Label lbl 
         Caption         =   "&Descuentos:"
         Height          =   285
         Index           =   11
         Left            =   4695
         TabIndex        =   29
         Top             =   2520
         Width           =   1575
      End
      Begin VB.Label lbl 
         Caption         =   "Dias No &Trabajados:"
         Height          =   285
         Index           =   10
         Left            =   1860
         TabIndex        =   27
         Top             =   2520
         Width           =   1575
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         Caption         =   "Otros &Bonos:"
         Height          =   285
         Index           =   4
         Left            =   7965
         TabIndex        =   25
         Top             =   1650
         Width           =   1170
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         Caption         =   "Dias &Feriados:"
         Height          =   285
         Index           =   2
         Left            =   2445
         TabIndex        =   23
         Top             =   1680
         Width           =   1125
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         Caption         =   "Otra &Asignación:"
         Height          =   285
         Index           =   3
         Left            =   4695
         TabIndex        =   20
         Top             =   1665
         Width           =   1170
      End
      Begin VB.Label lbl 
         Caption         =   "Apellidos y Nombres:"
         Height          =   285
         Index           =   5
         Left            =   240
         TabIndex        =   19
         Top             =   330
         Width           =   1575
      End
      Begin VB.Label lbl 
         Caption         =   "Cédula de Identidad:"
         Height          =   285
         Index           =   6
         Left            =   240
         TabIndex        =   18
         Top             =   765
         Width           =   1575
      End
      Begin VB.Label lbl 
         Caption         =   "Cargo:"
         Height          =   285
         Index           =   7
         Left            =   6780
         TabIndex        =   17
         Top             =   330
         Width           =   450
      End
      Begin VB.Label lbl 
         Caption         =   "Cód. Empleado:"
         Height          =   285
         Index           =   8
         Left            =   3030
         TabIndex        =   16
         Top             =   780
         Width           =   1425
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         Caption         =   "Sueldo Mensual:"
         Height          =   285
         Index           =   9
         Left            =   6000
         TabIndex        =   15
         Top             =   765
         Width           =   1230
      End
   End
   Begin VB.Frame fraNom 
      Caption         =   "Inmueble: "
      Height          =   1305
      Index           =   0
      Left            =   210
      TabIndex        =   1
      Top             =   540
      Width           =   11385
      Begin VB.CommandButton cmd 
         Caption         =   "Ver todo"
         Height          =   345
         Left            =   8250
         TabIndex        =   37
         Top             =   795
         Width           =   2805
      End
      Begin VB.TextBox TXT 
         DataField       =   "NombreBanco"
         Height          =   315
         Index           =   15
         Left            =   1590
         Locked          =   -1  'True
         TabIndex        =   4
         Top             =   810
         Width           =   1965
      End
      Begin VB.TextBox TXT 
         DataField       =   "NumCuenta"
         Height          =   315
         Index           =   16
         Left            =   3675
         Locked          =   -1  'True
         TabIndex        =   3
         Top             =   810
         Width           =   3315
      End
      Begin VB.ComboBox cmb 
         Height          =   315
         Left            =   8250
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   390
         Width           =   2820
      End
      Begin MSDataListLib.DataCombo dtc 
         Height          =   315
         Index           =   0
         Left            =   270
         TabIndex        =   5
         Tag             =   "CodInm"
         Top             =   390
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   556
         _Version        =   393216
         ListField       =   "CodInm"
         BoundColumn     =   "Nombre"
         Text            =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDataListLib.DataCombo dtc 
         Height          =   315
         Index           =   1
         Left            =   1590
         TabIndex        =   6
         Tag             =   "Nombre"
         Top             =   390
         Width           =   5415
         _ExtentX        =   9551
         _ExtentY        =   556
         _Version        =   393216
         ListField       =   "Nombre"
         BoundColumn     =   "CodInm"
         Text            =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label lbl 
         Caption         =   "Cuenta:"
         Height          =   285
         Index           =   1
         Left            =   270
         TabIndex        =   8
         Top             =   840
         Width           =   1095
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         Caption         =   "Nómina:"
         Height          =   285
         Index           =   27
         Left            =   7170
         TabIndex        =   7
         Top             =   420
         Width           =   975
      End
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid gridNom 
      Height          =   2580
      Left            =   210
      TabIndex        =   0
      Tag             =   "250|800|2200|2200|2000|1500"
      Top             =   1965
      Width           =   9405
      _ExtentX        =   16589
      _ExtentY        =   4551
      _Version        =   393216
      Cols            =   6
      BackColorFixed  =   -2147483646
      ForeColorFixed  =   -2147483639
      BackColorBkg    =   -2147483636
      GridColor       =   -2147483636
      FormatString    =   "|^CodEmp|Apellidos|Nombres|Cargo|Sueldo Mensual"
      _NumberOfBands  =   1
      _Band(0).Cols   =   6
   End
   Begin VB.Frame fraNom 
      Caption         =   "Buscar por: "
      Height          =   2685
      Index           =   1
      Left            =   9690
      TabIndex        =   40
      Top             =   1860
      Width           =   1890
      Begin VB.TextBox TXT 
         DataField       =   "NombreBanco"
         Height          =   315
         Index           =   19
         Left            =   150
         TabIndex        =   46
         Top             =   2175
         Width           =   1575
      End
      Begin VB.TextBox TXT 
         DataField       =   "NombreBanco"
         Height          =   315
         Index           =   18
         Left            =   150
         TabIndex        =   45
         Top             =   1410
         Width           =   1575
      End
      Begin VB.TextBox TXT 
         DataField       =   "NombreBanco"
         Height          =   315
         Index           =   17
         Left            =   150
         TabIndex        =   44
         Top             =   645
         Width           =   1575
      End
      Begin VB.Label lbl 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Código Empleado"
         ForeColor       =   &H80000008&
         Height          =   690
         Index           =   19
         Left            =   105
         TabIndex        =   43
         Top             =   330
         Width           =   1680
      End
      Begin VB.Label lbl 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Nombres"
         ForeColor       =   &H80000008&
         Height          =   660
         Index           =   18
         Left            =   90
         TabIndex        =   42
         Top             =   1890
         Width           =   1695
      End
      Begin VB.Label lbl 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Apellidos"
         ForeColor       =   &H80000008&
         Height          =   690
         Index           =   16
         Left            =   90
         TabIndex        =   41
         Top             =   1110
         Width           =   1695
      End
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      Caption         =   "Nómina:"
      Height          =   285
      Index           =   17
      Left            =   0
      TabIndex        =   39
      Top             =   0
      Width           =   975
   End
   Begin VB.Image img 
      Height          =   240
      Left            =   3000
      Picture         =   "frmNovNomina.frx":031A
      Top             =   1335
      Visible         =   0   'False
      Width           =   240
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
            Picture         =   "frmNovNomina.frx":0460
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmNovNomina.frx":05E2
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmNovNomina.frx":0764
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmNovNomina.frx":08E6
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmNovNomina.frx":0A68
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmNovNomina.frx":0BEA
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmNovNomina.frx":0D6C
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmNovNomina.frx":0EEE
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmNovNomina.frx":1070
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmNovNomina.frx":11F2
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmNovNomina.frx":1374
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmNovNomina.frx":14F6
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmNovNomina"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rstNom(3) As New ADODB.Recordset
Dim vecNom(3, 1) As String
Dim IDNomina As Long
Const m$ = "#,##0.00 "


Private Sub cmb_Click()
If rstNom(3).State = 1 Then
    IDNomina = vecNom(cmb.ListIndex, 1)
    'rstNom(3).Close
    'rstNom(3).Source = "SELECT * FROM Nom_Temp WHERE IDnomina=" & IDNomina
    'rstNom(3).Open
    'Me.Refresh
    rstNom(3).Filter = "IDNomina = " & IDNomina
    Call muestra_employes
End If
End Sub

Private Sub cmd_Click()
dtc(0) = ""
dtc(1) = ""
Call muestra_employes
End Sub

Private Sub dtc_Click(Index As Integer, Area As Integer)
'variables locales
If Area = 2 Then
    If Index = 0 Then dtc(1) = dtc(0).BoundText
    If Index = 1 Then dtc(0) = dtc(1).BoundText
    If dtc(0) <> "" Then Call muestra_employes("CodInm='" & dtc(0) & "'")
End If
'
End Sub

    Private Sub Form_Load()
    'variableslocales
    Dim strNom As String
    Dim FechaU As Date
    Dim strSQL As String
    Dim j%, K%, I%
    '----------------------------
    rstNom(0).Open "SELECT IDNomina as UNP FROM Nom_Inf WHERE Efectivo = (SELECT Max(Efectivo) " _
    & "FROM Nom_Inf WHERE IDNomina <> 312" & Year(Date) & ")", cnnConexion, adOpenKeyset, adLockOptimistic, adCmdText
    
    If Not rstNom(0).EOF And Not rstNom(0).BOF Then ultNomPro = rstNom(0).Fields("UNP")
    rstNom(0).Close
    
    For I = 0 To 1
        rstNom(I).Open "SELECT * FROM Inmueble ORDER BY " & dtc(I).Tag, cnnConexion, _
        adOpenKeyset, adLockOptimistic, adCmdText
        Set dtc(I).RowSource = rstNom(I)
    Next I
    '
    
    
    FechaU = DateSerial(Right(ultNomPro, 4), Mid(ultNomPro, 2, 2), IIf(Left(ultNomPro, 1) = 1, 15, 30))
    j = IIf(Left(ultNomPro, 1) = 1, 2, 1)
    For I = 0 To 2
        'agrega las nominas siguientes
        If Day(FechaU) = 15 Then
            strNom = UCase(j & "  quincena mes: " & Format(DateAdd("d", -1, DateAdd("m", 1 * (I + 1), "01/" & Format(FechaU, "mm/yy"))), "mmm-yyyy"))
        Else
            strNom = UCase(j & "  quincena mes: " & Format(DateAdd("d", 15 * (I + 1), FechaU), "mmm-yyyy"))
        End If
        cmb.AddItem strNom
        vecNom(K, 0) = strNom
        If Day(FechaU) = 15 Then
            vecNom(K, 1) = j & Format(DateAdd("d", -1, DateAdd("m", 1 * (I + 1), "01/" & Format(FechaU, "mm/yy"))), "mmyyyy")
        Else
            vecNom(K, 1) = j & Format(DateAdd("d", 15 * (I + 1), FechaU), "mmyyyy")
        End If
        K = K + 1
        j = IIf(j = 1, 2, 1)
    Next

    cmb = vecNom(0, 0)
    IDNomina = vecNom(0, 1)
    Set gridNom.FontFixed = LetraTitulo(LoadResString(527), 7.5, , True)
    Set gridNom.Font = LetraTitulo(LoadResString(528), 8)
    Call centra_titulo(gridNom, True)
    
    strSQL = "SELECT Emp.*,Emp_Cargos.NombreCargo FROM Emp LEFT JOIN " _
    & "Emp_Cargos ON Emp.CodCargo = Emp_Cargos.CodCargo WHERE " _
    & "Emp.CodEstado=0 ORDER BY emp.CodInm, Emp.Apellidos"
    rstNom(2).Open strSQL, cnnConexion, adOpenKeyset, adLockOptimistic, adCmdText
    For I = 9 To 14: Set txt(I).DataSource = rstNom(2)
    Next
    '-------
    strSQL = "SELECT * FROM Nom_Temp" ' WHERE IDnomina=" & IDNomina
    rstNom(3).Open strSQL, cnnConexion, adOpenKeyset, adLockOptimistic, adCmdText
    rstNom(3).Filter = "IDNomina = " & IDNomina
    For I = 0 To 7: Set txt(I).DataSource = rstNom(3)
    Next
    
    Call muestra_employes
    
    End Sub

    Private Sub Form_Resize()
    On Error Resume Next
    fraNom(0).Left = (ScaleWidth - fraNom(0).Width) / 2
    fraNom(2).Left = fraNom(0).Left
    gridNom.Left = fraNom(0).Left
    gridNom.Height = ScaleHeight - gridNom.Top - fraNom(2).Height
    fraNom(1).Left = fraNom(2).Width - fraNom(1).Width + fraNom(2).Left
    fraNom(1).Height = gridNom.Height
    fraNom(2).Top = gridNom.Top + gridNom.Height
    '
    End Sub

    Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    'cierra y descarga los objetos ADODB.Recordsets
    For I = 0 To 3
        rstNom(I).Close
        Set rstNom(I) = Nothing
    Next
    '
    End Sub

    Private Sub gridNom_EnterCell()
    'variables locales
    Dim colTemp As Long
    '
    colTemp = gridNom.ColSel
    gridNom.Col = 0
    gridNom.CellBackColor = vbActiveTitleBarText
    Set gridNom.CellPicture = img
    gridNom.CellPictureAlignment = flexAlignCenterCenter
    gridNom.Col = colTemp
    '
    With rstNom(2)
        .Find "CodEmp=" & gridNom.TextMatrix(gridNom.RowSel, 1)
        If .EOF Then
            .MoveFirst
            .Find "CodEmp=" & gridNom.TextMatrix(gridNom.RowSel, 1)
            
        End If
        lbl(15) = .AbsolutePosition & "/" & .RecordCount
        '
        With rstNom(3)
            Toolbar1.Buttons("New").Enabled = False
            Toolbar1.Buttons("Edit").Enabled = False
            If .EOF And .BOF Then
                Toolbar1.Buttons("New").Enabled = True
                Exit Sub
            End If
            .Find "CodEmp=" & rstNom(2)("CodEmp")
            If .EOF Then
                .MoveFirst
                .Find "CodEmp=" & rstNom(2)("CodEmp")
                If .EOF Then
                    Toolbar1.Buttons("New").Enabled = True
                    Toolbar1.Buttons("Edit").Enabled = False
                Else
                    Toolbar1.Buttons("Edit").Enabled = True
                End If
            Else
                Toolbar1.Buttons("Edit").Enabled = True
            End If
            '
        End With
    End With
    '
    End Sub

    Private Sub gridNom_LeaveCell()
    'variables locales
    Dim colTemp As Long
    '
    colTemp = gridNom.ColSel
    gridNom.Col = 0
    gridNom.CellBackColor = vbActiveTitleBar
    Set gridNom.CellPicture = Nothing
    gridNom.Col = colTemp
    End Sub
    
    Private Sub Toolbar1_ButtonClick(ByVal Button As ComctlLib.Button)
    'variables locales
    Dim ventana As Form
    Dim strSQL As String
    '
    With rstNom(2)
        '
        Select Case Button.Key
        
            Case "First"    'ir al primer registro
                .MoveFirst
                gridNom_LeaveCell
                gridNom.Row = .AbsolutePosition
                gridNom_EnterCell
                
                
            Case "Previous" 'registro anterior
                .MovePrevious
                If .BOF Then .MoveLast
                gridNom_LeaveCell
                gridNom.Row = .AbsolutePosition
                gridNom_EnterCell
                
            Case "Next" 'próximo registro
                .MoveNext
                If .EOF Then .MoveFirst
                gridNom_LeaveCell
                gridNom.Row = .AbsolutePosition
                gridNom_EnterCell
                
            Case "End"  'ir al último registro
                .MoveLast
                gridNom_LeaveCell
                gridNom.Row = .AbsolutePosition
                gridNom_EnterCell
                
            Case "New"  'nuevo registro
                'limpia los contronles
                fraNom(2).Enabled = True
                fraNom(0).Enabled = False
                rstNom(3).AddNew
                For I = 0 To 2: txt(I) = 0
                Next
                For I = 3 To 6: txt(I) = "0,00"
                Next
                txt(7) = ""
                Call RtnEstado(Button.Index, Toolbar1)
                
            Case "Save" 'guardar registro
                If Error Then Exit Sub
                fraNom(2).Enabled = False
                fraNom(0).Enabled = True
                If rstNom(3).EditMode = adEditAdd Then
                    rstNom(3)("IdNomina") = IDNomina
                    rstNom(3)("CodEmp") = txt(9)
                End If
                rstNom(3).Update
                For Each ventana In Forms
                    '
                    If ventana.Name = "frmNomina" Then
                        frmNomina.blnModif = True
                        Exit For
                    End If
                    '
                Next
                Call RtnEstado(Button.Index, Toolbar1)
                MsgBox "Registro actualizado", vbInformation, App.ProductName
            
                '
            Case "Undo" 'deshacer cambios
                fraNom(2).Enabled = False
                fraNom(0).Enabled = True
                rstNom(3).CancelUpdate
                MsgBox "Registro Cancelado", vbInformation, App.ProductName
                Call RtnEstado(Button.Index, Toolbar1)
            
            Case "Edit" 'editar registro
                fraNom(2).Enabled = True
                fraNom(0).Enabled = False
                Call RtnEstado(Button.Index, Toolbar1)
                '
            Case "Close"    'cerrar ventana
                Unload Me
                Set frmNovNomina = Nothing
            
            Case "Print"
                'elimina los registros innecesarios
                cnnConexion.Execute "DELETE * FROM Nom_Temp WHERE IDNomina=" & IDNomina & " AND" _
                & " DiasFer=0 AND DiasLib=0 and OtrasAsig=0 and OtrosBonos=0 and DiasNTrab=0 an" _
                & "d Descuentos=0 and OtrasDed=0;"
                '
                FrmReport.apto = IDNomina
                FrmReport.Show
                mcReport = "nom_nov.rpt"
                
        End Select
        '
    End With
    '
    End Sub
    
    '-------------------------------------------------------------------------------------------------
    '   Rutina: muestra_employes
    '
    '   Entrada: strFiltro
    '
    '   lista los empleados, aplicando el filtro
    '-------------------------------------------------------------------------------------------------
    Private Sub muestra_employes(Optional strFriltro As String)
    'variables locales
    Dim I As Long
    '
    rstNom(2).Filter = strFriltro
    With rstNom(2)
        gridNom.Rows = 2
        Call rtnLimpiar_Grid(gridNom)
        If Not .EOF Or Not .BOF Then
            .MoveFirst
            gridNom.Rows = .RecordCount + 1
            Do
                '
                I = I + 1
                With gridNom
                '
                    .TextMatrix(I, 1) = Format(rstNom(2)("CodEmp"), "000000")
                    .TextMatrix(I, 2) = IIf(IsNull(rstNom(2)("Apellidos")), "", _
                    rstNom(2)("Apellidos"))
                    .TextMatrix(I, 3) = IIf(IsNull(rstNom(2)("Nombres")), "", _
                    rstNom(2)("Nombres"))
                    .TextMatrix(I, 4) = IIf(IsNull(rstNom(2)("NombreCargo")), "", _
                    rstNom(2)("NombreCargo"))
                    .TextMatrix(I, 5) = Format(rstNom(2)("Sueldo"), m)
                    '
                End With
                .MoveNext
                '
            Loop Until .EOF
        End If
        lbl(15) = .RecordCount
    End With
    'On Error Resume Next
    'gridNom.SetFocus
    gridNom.Col = 5
    Call gridNom_EnterCell
    gridNom.Col = 1
    Call gridNom_EnterCell
    On Error Resume Next
    gridNom.SetFocus
    '
    End Sub
    
    
    Private Sub txt_GotFocus(Index As Integer)
    '
    If Index = 3 Or Index = 4 Or Index = 5 Or Index = 6 Then txt(Index) = CCur(txt(Index))
    txt(Index).SelStart = 0
    txt(Index).SelLength = Len(txt(Index))
    End Sub
    
    Private Sub txt_KeyPress(Index As Integer, KeyAscii As Integer)
    '
    KeyAscii = Asc(UCase(Chr(KeyAscii)))    'convierte todo a mayúsculas
    Select Case Index
        
        Case 0, 1, 2
            Call Validacion(KeyAscii, "0123456789")
            
        Case 3, 4, 5, 6
            If KeyAscii = Asc(".") Then KeyAscii = Asc(",")
            Call Validacion(KeyAscii, "0123456789,")
            
    End Select
    If KeyAscii = 13 Then SendKeys vbTab
    End Sub
    
    Private Sub txt_LostFocus(Index As Integer)
    '
    If Index = 3 Or Index = 4 Or Index = 5 Or Index = 6 Then
        If IsNumeric(txt(Index)) Then
            txt(Index) = Format(txt(Index), "#,##0.00")
        Else
            MsgBox "Introdujo un valor no válido", vbInformation, App.ProductName
            txt(Index).SelStart = 0
            txt(Index).SelLength = Len(txt(Index))
            txt(Index).SetFocus
        End If
    End If
    '
    End Sub
    
    '-------------------------------------------------------------------------------------------------
    '   Función:    error
    '
    '   Valida los datos mínimos requeridos para guardar un registro
    '
    '-------------------------------------------------------------------------------------------------
    Private Function Error() As Boolean
    'variables locales
    '
    For I = 0 To 6: If txt(I) = "" Or IsNull(txt(I)) Then txt(I) = 0
    Next I
    '
    End Function
