VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmNomina 
   Caption         =   "Detalle de Nómina"
   ClientHeight    =   435
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   2850
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   435
   ScaleWidth      =   2850
   WindowState     =   2  'Maximized
   Begin ComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   870
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Width           =   2850
      _ExtentX        =   5027
      _ExtentY        =   1535
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
            Key             =   "Save"
            Object.ToolTipText     =   "Guardar Registro"
            Object.Tag             =   ""
            ImageIndex      =   6
         EndProperty
         BeginProperty Button7 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Enabled         =   0   'False
            Key             =   "Find"
            Object.ToolTipText     =   "Ver todos los registros"
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
      MouseIcon       =   "frmNomina.frx":0000
   End
   Begin VB.CommandButton cmd 
      Caption         =   "Ver todo"
      Height          =   315
      Left            =   9120
      TabIndex        =   57
      Top             =   1890
      Width           =   2490
   End
   Begin VB.Frame fraNom 
      Caption         =   "Inmueble: "
      Height          =   1305
      Index           =   0
      Left            =   210
      TabIndex        =   1
      Top             =   570
      Width           =   11415
      Begin MSComCtl2.DTPicker dtp 
         Height          =   315
         Left            =   9690
         TabIndex        =   53
         Top             =   825
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   556
         _Version        =   393216
         CalendarTitleBackColor=   -2147483646
         CalendarTitleForeColor=   -2147483639
         Format          =   60686337
         CurrentDate     =   37957
      End
      Begin VB.ComboBox cmb 
         Height          =   315
         Left            =   8250
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   390
         Width           =   2820
      End
      Begin VB.TextBox TXT 
         DataField       =   "NumCuenta"
         Height          =   315
         Index           =   2
         Left            =   3675
         Locked          =   -1  'True
         TabIndex        =   6
         Top             =   810
         Width           =   3315
      End
      Begin VB.TextBox TXT 
         DataField       =   "NombreBanco"
         Height          =   315
         Index           =   1
         Left            =   1590
         Locked          =   -1  'True
         TabIndex        =   5
         Top             =   810
         Width           =   1965
      End
      Begin MSDataListLib.DataCombo dtc 
         Height          =   315
         Index           =   0
         Left            =   270
         TabIndex        =   2
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
         TabIndex        =   3
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
         Alignment       =   1  'Right Justify
         Caption         =   "Nómina:"
         Height          =   285
         Index           =   27
         Left            =   7170
         TabIndex        =   54
         Top             =   420
         Width           =   975
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         Caption         =   "Fecha Efectivo Nómina:"
         Height          =   285
         Index           =   25
         Left            =   7095
         TabIndex        =   8
         Top             =   885
         Width           =   2295
      End
      Begin VB.Label lbl 
         Caption         =   "Cuenta:"
         Height          =   285
         Index           =   1
         Left            =   270
         TabIndex        =   4
         Top             =   840
         Width           =   870
      End
   End
   Begin TabDlg.SSTab Ficha 
      Height          =   5955
      Left            =   195
      TabIndex        =   10
      Top             =   1920
      Width           =   11415
      _ExtentX        =   20135
      _ExtentY        =   10504
      _Version        =   393216
      Tabs            =   2
      TabHeight       =   520
      ShowFocusRect   =   0   'False
      TabCaption(0)   =   "Detalle Nómina"
      TabPicture(0)   =   "frmNomina.frx":031A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "fraNom(1)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "fraNom(2)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Lista Empleados"
      TabPicture(1)   =   "frmNomina.frx":0336
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "gridNom"
      Tab(1).ControlCount=   1
      Begin VB.Frame fraNom 
         Caption         =   "Empleado: "
         Height          =   1245
         Index           =   2
         Left            =   165
         TabIndex        =   58
         Top             =   330
         Visible         =   0   'False
         Width           =   11055
         Begin VB.TextBox TXT 
            Alignment       =   1  'Right Justify
            DataField       =   "Emp.Sueldo"
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
            Index           =   7
            Left            =   9255
            Locked          =   -1  'True
            TabIndex        =   62
            Text            =   "0,00"
            Top             =   750
            Width           =   1680
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
            Index           =   4
            Left            =   4710
            Locked          =   -1  'True
            TabIndex        =   61
            Top             =   300
            Width           =   1485
         End
         Begin VB.TextBox TXT 
            DataField       =   "Emp.CodEmp"
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
            Index           =   6
            Left            =   5565
            Locked          =   -1  'True
            TabIndex        =   60
            Top             =   750
            Width           =   1680
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
            Index           =   5
            Left            =   7035
            Locked          =   -1  'True
            TabIndex        =   59
            Top             =   315
            Width           =   3900
         End
         Begin MSDataListLib.DataCombo dtc 
            DataField       =   "Apellidos"
            Height          =   315
            Index           =   2
            Left            =   930
            TabIndex        =   63
            Tag             =   "Apellidos"
            Top             =   315
            Width           =   2925
            _ExtentX        =   5159
            _ExtentY        =   556
            _Version        =   393216
            ListField       =   "Apellidos"
            BoundColumn     =   "Nombres"
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
            DataField       =   "Nombres"
            Height          =   315
            Index           =   3
            Left            =   930
            TabIndex        =   64
            Tag             =   "Nombres"
            Top             =   795
            Width           =   2925
            _ExtentX        =   5159
            _ExtentY        =   556
            _Version        =   393216
            ListField       =   "Nombres"
            BoundColumn     =   "Apellidos"
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
            Caption         =   "Nombres:"
            Height          =   285
            Index           =   29
            Left            =   135
            TabIndex        =   70
            Top             =   810
            Width           =   825
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            Caption         =   "Sueldo Mensual:"
            Height          =   285
            Index           =   5
            Left            =   7905
            TabIndex        =   69
            Top             =   765
            Width           =   1230
         End
         Begin VB.Label lbl 
            Caption         =   "Código Empleado:"
            Height          =   285
            Index           =   4
            Left            =   4125
            TabIndex        =   68
            Top             =   780
            Width           =   1425
         End
         Begin VB.Label lbl 
            Caption         =   "Cargo:"
            Height          =   285
            Index           =   3
            Left            =   6495
            TabIndex        =   67
            Top             =   330
            Width           =   450
         End
         Begin VB.Label lbl 
            Caption         =   "C.I.:"
            Height          =   285
            Index           =   2
            Left            =   4200
            TabIndex        =   66
            Top             =   315
            Width           =   510
         End
         Begin VB.Label lbl 
            Caption         =   "Apellidos:"
            Height          =   285
            Index           =   0
            Left            =   105
            TabIndex        =   65
            Top             =   330
            Width           =   825
         End
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid gridNom 
         Height          =   4380
         Left            =   -74820
         TabIndex        =   0
         Tag             =   "250|800|2200|2200|2000|1500"
         Top             =   750
         Width           =   11055
         _ExtentX        =   19500
         _ExtentY        =   7726
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
         Height          =   3960
         Index           =   1
         Left            =   165
         TabIndex        =   11
         Top             =   1530
         Width           =   11070
         Begin VB.TextBox TXT 
            Alignment       =   1  'Right Justify
            DataField       =   "Bono_Otros"
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
            Index           =   0
            Left            =   2745
            Locked          =   -1  'True
            TabIndex        =   56
            Text            =   "0,00"
            Top             =   2940
            Width           =   1695
         End
         Begin VB.TextBox TXT 
            Alignment       =   2  'Center
            DataField       =   "Dias_Trab"
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
            Index           =   8
            Left            =   2070
            Locked          =   -1  'True
            TabIndex        =   32
            Tag             =   "9"
            Text            =   "0"
            Top             =   1200
            Width           =   570
         End
         Begin VB.TextBox TXT 
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
            Height          =   315
            Index           =   9
            Left            =   2775
            Locked          =   -1  'True
            TabIndex        =   31
            Text            =   "0,00"
            Top             =   1200
            Width           =   1695
         End
         Begin VB.TextBox TXT 
            Alignment       =   2  'Center
            DataField       =   "Dias_Feriados"
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
            Left            =   2070
            Locked          =   -1  'True
            TabIndex        =   30
            Tag             =   "11"
            Text            =   "0"
            Top             =   1650
            Width           =   570
         End
         Begin VB.TextBox TXT 
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
            Height          =   315
            Index           =   11
            Left            =   2730
            Locked          =   -1  'True
            TabIndex        =   29
            Text            =   "0,00"
            Top             =   1635
            Width           =   1695
         End
         Begin VB.TextBox TXT 
            Alignment       =   2  'Center
            DataField       =   "Dias_Libres"
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
            Left            =   2070
            Locked          =   -1  'True
            TabIndex        =   28
            Tag             =   "13"
            Text            =   "0"
            Top             =   2085
            Width           =   570
         End
         Begin VB.TextBox TXT 
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
            Height          =   315
            Index           =   13
            Left            =   2745
            Locked          =   -1  'True
            TabIndex        =   27
            Text            =   "0,00"
            Top             =   2070
            Width           =   1695
         End
         Begin VB.TextBox TXT 
            Alignment       =   1  'Right Justify
            DataField       =   "Bono_noc"
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
            Index           =   14
            Left            =   2745
            Locked          =   -1  'True
            TabIndex        =   26
            Text            =   "0,00"
            Top             =   2520
            Width           =   1695
         End
         Begin VB.TextBox TXT 
            Alignment       =   1  'Right Justify
            DataField       =   "Otras_Asignaciones"
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
            Index           =   15
            Left            =   2745
            Locked          =   -1  'True
            TabIndex        =   25
            Text            =   "0,00"
            Top             =   3375
            Width           =   1695
         End
         Begin VB.TextBox TXT 
            Alignment       =   1  'Right Justify
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
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   16
            Left            =   9090
            Locked          =   -1  'True
            TabIndex        =   24
            Text            =   "0,00"
            Top             =   1635
            Width           =   1695
         End
         Begin VB.TextBox TXT 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   17
            Left            =   9090
            Locked          =   -1  'True
            TabIndex        =   23
            Text            =   "0,00"
            Top             =   2505
            Width           =   1695
         End
         Begin VB.TextBox TXT 
            Alignment       =   1  'Right Justify
            DataField       =   "Otras_Deducciones"
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
            Index           =   18
            Left            =   7200
            Locked          =   -1  'True
            TabIndex        =   22
            Text            =   "0,00"
            Top             =   3375
            Width           =   1695
         End
         Begin VB.TextBox TXT 
            Alignment       =   1  'Right Justify
            DataField       =   "Descuento"
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
            Index           =   19
            Left            =   7185
            Locked          =   -1  'True
            TabIndex        =   21
            Text            =   "0,00"
            Top             =   2940
            Width           =   1695
         End
         Begin VB.TextBox TXT 
            Alignment       =   1  'Right Justify
            DataField       =   "nom_detalle.lph"
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
            Index           =   20
            Left            =   7185
            Locked          =   -1  'True
            TabIndex        =   20
            Text            =   "0,00"
            Top             =   2070
            Width           =   1695
         End
         Begin VB.TextBox TXT 
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
            Height          =   315
            Index           =   21
            Left            =   6510
            Locked          =   -1  'True
            TabIndex        =   19
            Tag             =   "20"
            Top             =   2070
            Width           =   570
         End
         Begin VB.TextBox TXT 
            Alignment       =   1  'Right Justify
            DataField       =   "nom_detalle.spf"
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
            Index           =   22
            Left            =   7170
            Locked          =   -1  'True
            TabIndex        =   18
            Text            =   "0,00"
            Top             =   1635
            Width           =   1695
         End
         Begin VB.TextBox TXT 
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
            Height          =   315
            Index           =   23
            Left            =   6510
            Locked          =   -1  'True
            TabIndex        =   17
            Tag             =   "22"
            Top             =   1635
            Width           =   570
         End
         Begin VB.TextBox TXT 
            Alignment       =   1  'Right Justify
            DataField       =   "nom_Detalle.SSO"
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
            Index           =   24
            Left            =   7185
            Locked          =   -1  'True
            TabIndex        =   16
            Text            =   "0,00"
            Top             =   1200
            Width           =   1695
         End
         Begin VB.TextBox TXT 
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
            Height          =   315
            Index           =   25
            Left            =   6510
            Locked          =   -1  'True
            TabIndex        =   15
            Tag             =   "24"
            Top             =   1200
            Width           =   570
         End
         Begin VB.TextBox TXT 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   315
            Index           =   26
            Left            =   9135
            Locked          =   -1  'True
            TabIndex        =   14
            Text            =   "0,00"
            Top             =   3135
            Width           =   1695
         End
         Begin VB.TextBox TXT 
            Alignment       =   1  'Right Justify
            DataField       =   "Dias_NoTrab"
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
            Index           =   27
            Left            =   6480
            Locked          =   -1  'True
            TabIndex        =   13
            Top             =   2490
            Width           =   570
         End
         Begin VB.TextBox TXT 
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
            Height          =   315
            Index           =   28
            Left            =   7215
            Locked          =   -1  'True
            TabIndex        =   12
            Text            =   "0,00"
            Top             =   2505
            Width           =   1695
         End
         Begin VB.Line lin 
            BorderColor     =   &H00FFFFFF&
            Index           =   14
            Visible         =   0   'False
            X1              =   1980
            X2              =   195
            Y1              =   1575
            Y2              =   1575
         End
         Begin VB.Label lbl 
            BackStyle       =   0  'Transparent
            Caption         =   "Otras Asignaciones:"
            DataField       =   "Otras_Asignaciones"
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Index           =   28
            Left            =   300
            TabIndex        =   55
            Top             =   3390
            Width           =   1695
         End
         Begin VB.Shape shp 
            Height          =   2955
            Index           =   1
            Left            =   4590
            Top             =   795
            Width           =   4440
         End
         Begin VB.Shape shp 
            DrawMode        =   1  'Blackness
            Height          =   2655
            Index           =   8
            Left            =   9015
            Top             =   1095
            Width           =   1860
         End
         Begin VB.Shape shp 
            Height          =   2955
            Index           =   0
            Left            =   165
            Top             =   795
            Width           =   4440
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H80000002&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "ASIGNACIONES"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000009&
            Height          =   345
            Index           =   6
            Left            =   165
            TabIndex        =   52
            Top             =   450
            Width           =   4440
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H80000002&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "DEDUCCIONES"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000009&
            Height          =   345
            Index           =   7
            Left            =   4620
            TabIndex        =   51
            Top             =   450
            Width           =   4425
         End
         Begin VB.Label lbl 
            BackStyle       =   0  'Transparent
            Caption         =   "Dias Trabajados:"
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Index           =   8
            Left            =   300
            TabIndex        =   50
            Top             =   1215
            Width           =   1695
         End
         Begin VB.Label lbl 
            BackStyle       =   0  'Transparent
            Caption         =   "Días Feriados:"
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Index           =   9
            Left            =   285
            TabIndex        =   49
            Top             =   1710
            Width           =   1380
         End
         Begin VB.Label lbl 
            BackStyle       =   0  'Transparent
            Caption         =   "Días Líbres:"
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Index           =   10
            Left            =   300
            TabIndex        =   48
            Top             =   2085
            Width           =   1695
         End
         Begin VB.Label lbl 
            BackStyle       =   0  'Transparent
            Caption         =   "Bono Nocturno:"
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Index           =   11
            Left            =   315
            TabIndex        =   47
            Top             =   2520
            Width           =   1695
         End
         Begin VB.Label lbl 
            BackStyle       =   0  'Transparent
            Caption         =   "Otros Bonificaciones:"
            DataField       =   "Otras_Asignaciones"
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Index           =   12
            Left            =   300
            TabIndex        =   46
            Top             =   2955
            Width           =   1695
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "TOTAL ASIGNACIONES"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   450
            Index           =   13
            Left            =   9030
            TabIndex        =   45
            Top             =   1125
            Width           =   1800
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            Caption         =   "Cant."
            Height          =   210
            Index           =   14
            Left            =   2160
            TabIndex        =   44
            Top             =   870
            Width           =   570
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            Caption         =   "Monto en Bs."
            Height          =   225
            Index           =   15
            Left            =   2805
            TabIndex        =   43
            Top             =   870
            Width           =   1695
         End
         Begin VB.Shape shp 
            Height          =   315
            Index           =   2
            Left            =   165
            Top             =   795
            Width           =   4440
         End
         Begin VB.Shape shp 
            BorderStyle     =   6  'Inside Solid
            FillColor       =   &H80000002&
            Height          =   315
            Index           =   3
            Left            =   4590
            Top             =   795
            Width           =   4440
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "TOTAL DEDUCCIONES"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   360
            Index           =   16
            Left            =   9075
            TabIndex        =   42
            Top             =   2055
            Width           =   1710
         End
         Begin VB.Label lbl 
            BackStyle       =   0  'Transparent
            Caption         =   "Otras Deducciones:"
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Index           =   17
            Left            =   4740
            TabIndex        =   41
            Top             =   3390
            Width           =   1845
         End
         Begin VB.Label lbl 
            BackStyle       =   0  'Transparent
            Caption         =   "Descuentos:"
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Index           =   18
            Left            =   4740
            TabIndex        =   40
            Top             =   2955
            Width           =   1845
         End
         Begin VB.Label lbl 
            BackStyle       =   0  'Transparent
            Caption         =   "L.P.H."
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Index           =   19
            Left            =   4740
            TabIndex        =   39
            Top             =   2085
            Width           =   1845
         End
         Begin VB.Label lbl 
            BackStyle       =   0  'Transparent
            Caption         =   "P.F."
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Index           =   20
            Left            =   4740
            TabIndex        =   38
            Top             =   1650
            Width           =   1845
         End
         Begin VB.Label lbl 
            BackStyle       =   0  'Transparent
            Caption         =   "S.S.O."
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Index           =   21
            Left            =   4740
            TabIndex        =   37
            Top             =   1245
            Width           =   1845
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Monto en Bs."
            Height          =   225
            Index           =   22
            Left            =   7185
            TabIndex        =   36
            Top             =   855
            Width           =   1695
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "%"
            Height          =   210
            Index           =   23
            Left            =   6540
            TabIndex        =   35
            Top             =   855
            Width           =   570
         End
         Begin VB.Line lin 
            Index           =   2
            Visible         =   0   'False
            X1              =   2685
            X2              =   2685
            Y1              =   1095
            Y2              =   2445
         End
         Begin VB.Line lin 
            Index           =   3
            Visible         =   0   'False
            X1              =   7125
            X2              =   7125
            Y1              =   1110
            Y2              =   2895
         End
         Begin VB.Line lin 
            Index           =   7
            Visible         =   0   'False
            X1              =   9015
            X2              =   6435
            Y1              =   1575
            Y2              =   1575
         End
         Begin VB.Line lin 
            BorderColor     =   &H00FFFFFF&
            Index           =   9
            Visible         =   0   'False
            X1              =   6405
            X2              =   4620
            Y1              =   1575
            Y2              =   1575
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "NETO A PAGAR"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Index           =   24
            Left            =   9345
            TabIndex        =   34
            Top             =   2940
            Width           =   1050
         End
         Begin VB.Line lin 
            Index           =   8
            Visible         =   0   'False
            X1              =   10860
            X2              =   6405
            Y1              =   2010
            Y2              =   2010
         End
         Begin VB.Line lin 
            Index           =   6
            Visible         =   0   'False
            X1              =   9000
            X2              =   6420
            Y1              =   2445
            Y2              =   2445
         End
         Begin VB.Line lin 
            Index           =   5
            Visible         =   0   'False
            X1              =   9000
            X2              =   6420
            Y1              =   3315
            Y2              =   3315
         End
         Begin VB.Line lin 
            BorderColor     =   &H00FFFFFF&
            Index           =   10
            Visible         =   0   'False
            X1              =   6405
            X2              =   4620
            Y1              =   2010
            Y2              =   2010
         End
         Begin VB.Line lin 
            BorderColor     =   &H00FFFFFF&
            Index           =   11
            Visible         =   0   'False
            X1              =   6405
            X2              =   4620
            Y1              =   2445
            Y2              =   2445
         End
         Begin VB.Line lin 
            BorderColor     =   &H00FFFFFF&
            Index           =   12
            Visible         =   0   'False
            X1              =   6405
            X2              =   4620
            Y1              =   3315
            Y2              =   3315
         End
         Begin VB.Line lin 
            BorderColor     =   &H00FFFFFF&
            Index           =   15
            Visible         =   0   'False
            X1              =   1995
            X2              =   210
            Y1              =   2010
            Y2              =   2010
         End
         Begin VB.Line lin 
            BorderColor     =   &H00FFFFFF&
            Index           =   16
            Visible         =   0   'False
            X1              =   1965
            X2              =   180
            Y1              =   2445
            Y2              =   2445
         End
         Begin VB.Line lin 
            BorderColor     =   &H00FFFFFF&
            Index           =   17
            Visible         =   0   'False
            X1              =   1980
            X2              =   195
            Y1              =   2880
            Y2              =   2880
         End
         Begin VB.Line lin 
            BorderColor     =   &H00FFFFFF&
            Index           =   18
            Visible         =   0   'False
            X1              =   1980
            X2              =   195
            Y1              =   3315
            Y2              =   3315
         End
         Begin VB.Line lin 
            Index           =   19
            Visible         =   0   'False
            X1              =   4575
            X2              =   1995
            Y1              =   1575
            Y2              =   1575
         End
         Begin VB.Line lin 
            Index           =   20
            Visible         =   0   'False
            X1              =   4575
            X2              =   1995
            Y1              =   2010
            Y2              =   2010
         End
         Begin VB.Line lin 
            Index           =   21
            Visible         =   0   'False
            X1              =   4575
            X2              =   1995
            Y1              =   2445
            Y2              =   2445
         End
         Begin VB.Line lin 
            Index           =   22
            Visible         =   0   'False
            X1              =   4575
            X2              =   1995
            Y1              =   2880
            Y2              =   2880
         End
         Begin VB.Line lin 
            Index           =   23
            Visible         =   0   'False
            X1              =   4575
            X2              =   1995
            Y1              =   3315
            Y2              =   3315
         End
         Begin VB.Label lbl 
            BackStyle       =   0  'Transparent
            Caption         =   "Días No Trabajados:"
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Index           =   26
            Left            =   4725
            TabIndex        =   33
            Top             =   2520
            Width           =   1845
         End
         Begin VB.Line lin 
            BorderColor     =   &H00FFFFFF&
            Index           =   24
            Visible         =   0   'False
            X1              =   6420
            X2              =   4635
            Y1              =   2880
            Y2              =   2880
         End
         Begin VB.Line lin 
            Index           =   25
            Visible         =   0   'False
            X1              =   10845
            X2              =   6420
            Y1              =   2880
            Y2              =   2880
         End
         Begin VB.Shape shp 
            BackColor       =   &H80000002&
            BackStyle       =   1  'Opaque
            BorderColor     =   &H00FFFFFF&
            FillColor       =   &H80000002&
            Height          =   2655
            Index           =   4
            Left            =   165
            Top             =   1095
            Width           =   1845
         End
         Begin VB.Shape shp 
            BackColor       =   &H80000002&
            BackStyle       =   1  'Opaque
            BorderColor     =   &H00FFFFFF&
            FillColor       =   &H80000002&
            Height          =   2655
            Index           =   5
            Left            =   4590
            Top             =   1095
            Width           =   1845
         End
      End
   End
   Begin VB.Image img 
      Height          =   240
      Left            =   3000
      Picture         =   "frmNomina.frx":0352
      Top             =   1950
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
            Picture         =   "frmNomina.frx":0498
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmNomina.frx":061A
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmNomina.frx":079C
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmNomina.frx":091E
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmNomina.frx":0AA0
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmNomina.frx":0C22
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmNomina.frx":0DA4
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmNomina.frx":0F26
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmNomina.frx":10A8
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmNomina.frx":122A
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmNomina.frx":13AC
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmNomina.frx":152E
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmNomina"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
    '---------------------------------------------------------------------------------------------
    '   Módulo de Nómina
    '
    '   Detalle de la nómina
    '   27/11/2003.-SinaiTech, C.A.
    '---------------------------------------------------------------------------------------------
    'varialbes publicas a nivel de módulo
    Dim WithEvents rst As ADODB.Recordset   'recordset principal que se activan los eventos
Attribute rst.VB_VarHelpID = -1
    Dim rstNom(5) As New ADODB.Recordset    'matriz de ADODB.Recordsets
    Dim vecNom(3, 1) As String
    Dim idnomina&        'id de la nómina    Nº Q + Nº Mes + AÑO ejm, 1022004
    Const m$ = "#,##0.00;(#,##0.00)"
    Dim pSSO@, pSPF@, pLPH@, pDFT@  'contiene el porcentaje de calculo
    Public blnModif As Boolean  'variable publica a nivel de este modulo, determina si han cambiad
    

    '-------------------------------------------------------------------------------------------------
    Private Sub cmb_Click()
    '
    idnomina = vecNom(cmb.ListIndex, 1)
    Call Pre_nomina(dtc(0))
    If dtc(0) <> "" Then
        dtc_Click 0, 2
    Else
        Call muestra_employes
    End If
    '
    End Sub

    Private Sub cmd_Click()
    dtc(0) = ""
    dtc(1) = ""
    Txt(1) = ""
    Txt(2) = ""
    Call muestra_employes
    End Sub

    Private Sub dtc_Click(Index As Integer, Area As Integer)
    '
    If Area = 2 Then    'selecciona un elemento de la lista
        '
        If Index = 0 Or Index = 1 Then
            dtc(IIf(Index = 0, 1, 0)) = dtc(IIf(Index = 0, 0, 1)).BoundText
            Call Encabezado(dtc(0))
            '
        ElseIf Index = 2 Or Index = 3 Then
            On Error Resume Next
            dtc(IIf(Index = 2, 3, 2)) = dtc(IIf(Index = 2, 2, 3)).BoundText
            If dtc(2) = "" Then
                Call busca_employes("Nombres ='" & dtc(3) & "'")
            Else
                Call busca_employes("Apellidos ='" & dtc(2) & "'")
            End If
            '
        End If
        '
    End If
    '
    End Sub

    Private Sub dtc_KeyPress(Index As Integer, KeyAscii As Integer)
    'variables locales
    Select Case Index
        '
        Case 0  '
            If Len(dtc(0)) = 4 And KeyAscii <> 13 Then KeyAscii = 0: Exit Sub
            Call Validacion(KeyAscii, "0123456789")
            If KeyAscii = 13 Then
                dtc(1) = dtc(0).BoundText
                If dtc(0) <> "" Then Call Encabezado(dtc(0))
            End If
        '
        Case 1  '
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
            If KeyAscii = 13 Then
                dtc(0) = dtc(1).BoundText
                If dtc(0) <> "" Then Call Encabezado(dtc(0))
            End If
        '
        Case 2
        '
        Case 3
        '
    End Select
    '
    End Sub

    Private Sub Ficha_Click(PreviousTab As Integer)
    'variables locales
    Select Case Ficha.tab
        '
        Case 0  'datos generales
            fraNom(2).Visible = True
            Toolbar1.Buttons(1).Enabled = True
            Toolbar1.Buttons(2).Enabled = True
            Toolbar1.Buttons(3).Enabled = True
            Toolbar1.Buttons(4).Enabled = True
        '
        Case 1  'listado
            fraNom(2).Visible = False
            Toolbar1.Buttons(1).Enabled = False
            Toolbar1.Buttons(2).Enabled = False
            Toolbar1.Buttons(3).Enabled = False
            Toolbar1.Buttons(4).Enabled = False
            gridNom.SetFocus
        '
    End Select
    '
    End Sub

    Private Sub Form_Activate()
    'variables locales
    Dim Emp As Long
    Dim resp As Integer
    '
    If blnModif Then
    
        resp = MsgBox("Algunos valores han cambiado desde la última vez que vió esta informació" _
        & "n." & vbCrLf & "¿Desea actualizar ahora?", vbQuestion + vbYesNo, App.ProductName)
        If resp = vbYes Then
            MousePointer = vbHourglass
            Emp = rst("Emp.CodEmp")
            Call Pre_nomina            'Call muestra_employes
            If dtc(0) <> "" Then dtc_Click 0, 2
            rst.MoveFirst
            rst.Find "Emp.CodEmp=" & Emp
            MousePointer = vbDefault
        End If
        blnModif = False
        
    End If
    '
    End Sub

    Private Sub Form_Load()
    'variables locales
    Dim strNom As String
    Dim FechaU As Date
    Dim ultNomPro As Long
    Dim j%, K%, I%
    '
    'activa los empleados que le corresponde reintegrarse en esta nómina
    cnnConexion.Execute "UPDATE Emp SET CodEstado = 0 WHERE CodEmp IN (SELECT CodEmp FROM " & _
     "Nom_Vaca WHERE IDNomina = " & idnomina & ")", j
    ' cnnConexion.Execute "UPDATE Emp INNER JOIN Nom_Vaca ON Emp.CodEmp = Nom_Vaca.CodEmp " & _
     "SET Emp.CodEstado = 0 WHERE (((Nom_Vaca.IDNomina)=" & IDNomina & "));"
    Call rtnBitacora(j & " empleado(s) activado(s)")
    
    Set rst = New ADODB.Recordset
    'datediff("ww","01/11/2003","30/11/2003",vbWednesday)
    'obtiene el valor de las variables de calculos
    rstNom(0).Open "Nom_calc", cnnConexion, adOpenKeyset, adLockOptimistic, adCmdTable
    dtp.Value = Date
    '
    If Not rstNom(0).EOF Or rstNom(0).BOF Then
    
        'asigna valor a las variables
        pSSO = rstNom(0).Fields("SSO")
        pSPF = rstNom(0).Fields("SPF")
        pLPH = rstNom(0).Fields("LPH")
        pDFT = rstNom(0).Fields("Dia_Feriado")
        Txt(25) = pSSO
        Txt(23) = pSPF
        Txt(21) = pLPH
        pSSO = pSSO / 100
        pSPF = pSPF / 100
        pLPH = pLPH / 100
        'pDFT = rstNom(0).Fields("Dia_Feriado")
        
    Else
        MsgBox "Faltan Tasas para el calculo de SSO,SPF,LPH"
    End If
    rstNom(0).Close
    'caculo preliminar de la nomina
    'Call Pre_nomina
    rstNom(0).Open "SELECT IDNomina as UNP FROM Nom_Inf WHERE Efectivo = (SELECT Max(Efectivo) " _
    & "FROM Nom_Inf WHERE IDNomina <> 312" & Year(Date) & ")", cnnConexion, adOpenKeyset, _
    adLockOptimistic, adCmdText
    
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
    '
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
    '
    cmb = vecNom(0, 0)
    idnomina = vecNom(0, 1)
    
    Set gridNom.FontFixed = LetraTitulo(LoadResString(527), 7.5, , True)
    Set gridNom.Font = LetraTitulo(LoadResString(528), 8)
    Call centra_titulo(gridNom, True)
    'busca la informacion de todos los empleados
    Call muestra_employes
    '
    End Sub

    Private Sub Form_Resize()
    On Error Resume Next
    
    fraNom(0).Left = (ScaleWidth - fraNom(0).Width) / 2
    Ficha.Left = (ScaleWidth - Ficha.Width) / 2
    Ficha.Height = ScaleHeight - Ficha.Top
    Cmd.Left = Ficha.Left + Ficha.Width - Cmd.Width
    gridNom.Height = Ficha.Height - Ficha.Top
    '
    End Sub

    Private Sub Form_Unload(Cancel As Integer)
    '
    On Error Resume Next
    For I = 0 To 5
        rstNom(I).Close
        Set rstNom(I) = Nothing
    Next
    rst.Close
    Set rst = Nothing
    '
    End Sub

    Private Sub gridNom_EnterCell()
    'variables locales
    Dim colTemp As Long
    '
    colTemp = gridNom.Col
    gridNom.Col = 0
    gridNom.CellBackColor = vbActiveTitleBarText
    Set gridNom.CellPicture = img
    gridNom.Col = colTemp
    '
    With rst
        .Find "Emp.CodEmp=" & gridNom.TextMatrix(gridNom.RowSel, 1)
        If .EOF Then
            .MoveFirst
            .Find "Emp.CodEmp=" & gridNom.TextMatrix(gridNom.RowSel, 1)
        End If
    End With
    '
    End Sub

    Private Sub gridNom_LeaveCell()
    'variables locales
    Dim colTemp As Long
    '
    colTemp = gridNom.Col
    gridNom.Col = 0
    gridNom.CellBackColor = vbActiveTitleBar
    Set gridNom.CellPicture = Nothing
    gridNom.Col = colTemp
    '
    End Sub

    Private Sub rst_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, _
    ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, _
    ByVal pRecordset As ADODB.Recordset)
    'variables locales
    Dim curS As Currency
    '
    If rst.EOF Or rst.BOF Then Exit Sub
    dtc(2) = IIf(IsNull(rst("Apellidos")), "", rst("Apellidos"))
    dtc(3) = rst.Fields("Nombres")
    '
    With rst
        'asignaciones
        curS = .Fields("Emp.Sueldo") '+ (.Fields("Emp.Sueldo") * .Fields("BonoNoc") / 100)
        Txt(9) = Format(curS / 30 * !Dias_Trab, m)
        Txt(11) = Format((curS + (.Fields("Emp.Sueldo") * .Fields("BonoNoc") / 100)) / 30 * _
        (!dias_Feriados * pDFT), m)
        Txt(13) = Format(curS / 30 * !dias_libres, m)
        'total asignaciones
        Txt(16) = Format(CCur(Txt(9)) + CCur(Txt(11)) + CCur(Txt(13)) + !Bono_Noc + !Bono_Otros _
        + !Otras_Asignaciones, m)
        'deducciones
        Txt(28) = Format(curS / 30 * !Dias_NoTrab, m)
        'total deducciones
        Txt(17) = Format(.Fields("Nom_Detalle.SSO") + .Fields("Nom_Detalle.SPF") + _
        .Fields("Nom_Detalle.LPH") + CCur(Txt(28)) + !Descuento + !Otras_Deducciones, m)
        'total a cobrar
        Txt(26) = Format(CCur(Txt(16)) - CCur(Txt(17)), m)
        '
    End With
    '
    End Sub


    Private Sub Toolbar1_ButtonClick(ByVal Button As ComctlLib.Button)
    'variables locales
    Dim resp As Long
    Dim strSQL As String
    Dim Fecha1 As Date
    '
    With rst
    '
        Select Case Button.Key
            '
            Case "First"    'mover al primero
                .MoveFirst
                
            Case "Next"     'mover al siguiente
                .MoveNext
                If .EOF Then .MoveFirst
                
            Case "Previous" 'mover al anterior
                .MovePrevious
                If .BOF Then .MoveLast
                
            Case "End"      'mover al final
                .MoveLast
                
            Case "Find"     'buscar
                MsgBox "Opciòn no disponible", vbInformation, App.ProductName
                
            Case "Print"    'imprimir
                '
                strSQL = "SELECT Emp.CodInm, Nom_Detalle.CodEmp, Emp.Apellidos, Emp.Nombres,  N" _
                & "om_Detalle.Sueldo, Nom_Detalle.Dias_Trab,( Nom_Detalle.Sueldo + ( Nom_Detalle.Sueldo *  Nom_Detalle.Porc_BonoNoc / 10" _
                & "0)) /30 *  Nom_Detalle.Dias_libres as DL, ( Nom_Detalle.Sueldo + ( Nom_Detalle.Sueldo *  Nom_Detalle.Porc_BonoNoc" _
                & " / 100)) /30 *  Nom_Detalle.Dias_Feriados * '" & pDFT & "'  as DF, Nom_" _
                & "Detalle.Bono_Noc, Nom_Detalle.Bono_Otros, Nom_Detalle.Otras_Asignaciones, No" _
                & "m_Detalle.Dias_NoTrab, Nom_Detalle.SSO, Nom_Detalle.SPF, Nom_Detalle.LPH, No" _
                & "m_Detalle.Otras_Deducciones, Nom_Detalle.Descuento,Emp.CodGasto FROM Emp INN" _
                & "ER JOIN Nom_Detalle ON Emp.CodEmp = Nom_Detalle.CodEmp WHERE Nom_Detalle.IDN" _
                & "om=" & idnomina
                '
                Call rtnGenerator(gcPath & "\sac.mdb", strSQL, "qdfNomina")
                mcTitulo = "Pre-Nomina"
                FrmReport.apto = cmb.Text
                mcReport = "nom_report.rpt"
                mcDatos = gcPath + "\sac.mdb"
                mcOrdCod = ""
                mcOrdAlfa = ""
                mcCrit = ""
                FrmReport.Show 1
                mcDatos = gcPath + gcUbica + "Inm.mdb"
                '
            Case "Close"    'cerrar
                Unload Me
                
            Case "Edit"
                MsgBox "Opcion no disponible"
                'Call RtnEstado(Button.Index, Toolbar1)
                
            Case "Save"
                resp = MsgBox("¿Está seguro de procesar la nómina de :" & cmb & vbCrLf & "Fecha Efectivo: " & dtp, vbQuestion + _
                vbYesNo + vbDefaultButton2, App.ProductName)
                
                If resp = vbYes Then
                    '
                    'genera las consultas necesarias
                    On Error Resume Next
                    '
                    strSQL = "SELECT Inmueble.CodInm, Inmueble.Caja, Inmueble.CtaInm, Emp.CodEm" _
                    & "p, Emp.Apellidos & ', ' & Emp.Nombres AS Name,Emp.Cedula,Emp.Cuenta,(Nom" _
                    & "_Detalle.Sueldo/30) AS D,(d*Nom_Detalle.Dias_Trab)+(Nom_Detalle.Sueldo + (Nom_Detalle.Sue" _
                    & "ldo * Emp.BonoNoc / 100)) /30 *  Nom_Detalle.Dias_libres +(Nom_Detalle.Sueldo + " _
                    & "(Nom_Detalle.Sueldo * Nom_Detalle.Porc_BonoNoc / 100)) /30 *  Nom_Detalle.Dias_Feriados * '" _
                    & pDFT & "' +Nom_Detalle.Bono_Noc+Nom_Detalle.Bono_Otros+Nom_Detalle.Otras_" _
                    & "Asignaciones-((d*Nom_Detalle.Dias_NoTrab)+Nom_Detalle.SSO+Nom_Detalle.SP" _
                    & "F+Nom_Detalle.LPH+Nom_Detalle.Otras_Deducciones+Nom_Detalle.Descuento) A" _
                    & "S Neto, Inmueble.Nombre, Emp_Cargos.NombreCargo,Emp.CodGasto,Emp.Nacionalidad FROM Emp_Ca" _
                    & "rgos INNER JOIN (Inmueble INNER JOIN (Emp INNER JOIN Nom_Detalle ON Emp." _
                    & "CodEmp=Nom_Detalle.CodEmp) ON Inmueble.CodInm = Emp.CodInm) ON Emp_Cargo" _
                    & "s.CodCargo=Emp.CodCargo Where Inmueble.Inactivo = False And Nom_Detalle." _
                    & "IDNom = " & idnomina & " ORDER BY Inmueble.Caja, Inmueble.CodInm;"
                    '
                    Call rtnGenerator(gcPath & "\SAC.MDB", strSQL, "qdfNomina_Bco")
                    '
                    strSQL = "SELECT Emp.CodInm, Nom_Detalle.CodEmp, Emp.Apellidos, Emp.Nombres" _
                    & ",  Nom_Detalle.Sueldo, Nom_Detalle.Dias_Trab,( Nom_Detalle.Sueldo + ( Nom_Detalle.Sueldo *  Nom_Detalle.Porc_BonoNoc" _
                    & " / 100)) /30 *  Nom_Detalle.Dias_libres as DL, ( Nom_Detalle.Sueldo + ( " _
                    & "Nom_Detalle.Sueldo *  Nom_Detalle.Porc_BonoNoc / 100)) /30 *  Nom_Detalle.Dias_Feriados * '" & pDFT & _
                    "'  as DF, Nom_Detalle.Bono_Noc, Nom_Detalle.Bono_Otros, Nom_Detalle.Otras_" _
                    & "Asignaciones, Nom_Detalle.Dias_NoTrab, Nom_Detalle.SSO, Nom_Detalle.SPF," _
                    & "Nom_Detalle.LPH, Nom_Detalle.Otras_Deducciones, Nom_Detalle.Descuento,Em" _
                    & "p.CodGasto FROM Emp INNER JOIN Nom_Detalle ON Emp.CodEmp = Nom_Detalle.C" _
                    & "odEmp WHERE Nom_Detalle.IDNom=" & idnomina
                    'impresión
                    Call rtnGenerator(gcPath & "\sac.mdb", strSQL, "qdfNomina")
                    '
                    If cmb = "" Then
                        MsgBox "Por favor virifique el contenido del combo", vbInformation, App.ProductName
                        Exit Sub
                    End If
                    
                    If Not cierre_nomina(idnomina, cmb, dtp, pDFT) Then
                        Call rtnBitacora("Nomina OK.")
                        Fecha1 = IIf(Left(idnomina, 1) = 1, "01/", "16/") & Mid(idnomina, 2, 2) & "/" & Right(idnomina, 4)
                        'actualiza los empleados q salen de vacaciones
                        strSQL = "UPDATE Emp SET CodEstado=3 WHERE CodEmp IN (SELECT CodEmp FROM" _
                        & " Nom_Vaca WHERE Inicia>#" & Format(Fecha1, "mm/dd/yy") & "# and Incorpora <=#" & _
                        Format(DateAdd("d", -1, "01/" & Format(DateAdd("m", 1, Fecha1), "mm/yy")), _
                        "mm/dd/yy") & "#)"
                        cnnConexion.Execute strSQL, N
                        Call rtnBitacora(N & "Empleados de vacaciones")
                        MsgBox "Nómina procesada con éxito", vbInformation, App.ProductName
                    Else
                        Call rtnBitacora("Nómina Fallida.")
                    End If
                    
                End If
            '
        End Select
    '
    End With
    '
    End Sub

    '-------------------------------------------------------------------------------------------------
    '   Rutina:     Encabezado
    '
    '   Entrada: Inm
    '
    '-------------------------------------------------------------------------------------------------
    Private Sub Encabezado(Inm As String)
    'variables locales
    Dim Cadena_Conexion As String
    Dim Titulo As String
    '----------------------------
    '
    With rstNom(0)
        '
        .MoveFirst
        .Find "CodInm='" & Inm & "'"
        If Not .EOF Then
            If .Fields("Caja") = sysCodCaja Then
                Cadena_Conexion = cnnOLEDB & gcPath & "\" & sysCodInm & "\inm.mdb"
            Else
                Cadena_Conexion = cnnOLEDB & gcPath & .Fields("Ubica") & "inm.mdb"
            End If
            Call muestra_employes(" AND CodInm='" & Inm & "'")
            '
            If Not IsNull(.Fields("CtaInm")) Then
                If rstNom(2).State = 1 Then rstNom(2).Close
                rstNom(2).Open "SELECT Bancos.NombreBanco, Cuentas.NumCuenta,Cuentas.IDCuenta F" _
                & "ROM Bancos INNER JOIN Cuentas ON Bancos.IDBanco = Cuentas.IDBanco WHERE Cuen" _
                & "tas.IDCuenta=" & .Fields("CtaInm"), Cadena_Conexion, adOpenKeyset, _
                adLockOptimistic, adCmdText
                '
                If Not .EOF Or Not .BOF Then
                    Txt(1) = rstNom(2).Fields("NombreBanco")
                    Txt(2) = rstNom(2).Fields("NumCuenta")
                Else
                    Txt(1) = ""
                    Txt(2) = ""
                End If
            Else
                Txt(1) = "FALTA INF."
                Txt(2) = "FALTA INFORMACION"
            End If
            '
        End If
        '
    End With
    '
    End Sub

    '-------------------------------------------------------------------------------------------------
    '   Rutina:     busca_employes
    '
    '   Entrada Criterio
    '
    '   Rutina que busca la información de los empleados para un determinada nómina
    '-------------------------------------------------------------------------------------------------
    Private Sub busca_employes(Criterio As String)
    'variables locales
    
    With rstNom(3)
        .MoveFirst
        .Find Criterio
    End With
    '
    End Sub


    Private Sub txt_KeyPress(Index As Integer, KeyAscii As Integer)
    '
    Select Case Index

        Case 8, 10, 12, 25, 23, 21
        
            Call Validacion(KeyAscii, "0123456789")
            If KeyAscii = 13 Then
                Txt(Txt(Index).Tag) = Format(CCur(Txt(7)) / 30 * CCur(Txt(Index)), m)
            End If
        '
        Case 14, 15, 18, 19
        
            If KeyAscii = Asc(Chr(".")) Then KeyAscii = Asc(Chr(","))
            Call Validacion(KeyAscii, "0123456789,")
            If KeyAscii = 13 Then Txt(Index) = Format(Txt(Index), m)
            '
    End Select
    '
    End Sub

    '---------------------------------------------------------------------------------------------
    '   Rutina:     pre_nomina
    '
    '
    '   Entrada:    Inm 'codigo del inmueble'
    '
    '   Rutina que efectua el calcula de  la nomina, pero no efectua el cierre
    '
    '---------------------------------------------------------------------------------------------
    Public Sub Pre_nomina(Optional Inm As String)
    'variables locales
    Dim rstNom As New ADODB.Recordset
    Dim strSQL As String, Calc As String
    Dim K As Date, q As Date, j As Date
    Dim Sem As Long
    '
    If idnomina = 0 Then   'asgina valor al id. de la nomina
        idnomina = IIf(Day(Date) <= 15, 1, 2) & Format(Month(DateAdd("m", j, Date)), "00") & _
        Year(Date)
    End If
    '
    'varifica que la nónima no este ya cerrada
    Set rstNom = cnnConexion.Execute("SELECT * FROM Nom_Inf WHERE IDNomina=" & idnomina)
    '
    If Not rstNom.EOF And Not rstNom.BOF Then
        rstNom.Close
        MsgBox "Nómina ya cerrada", vbInformation, App.ProductName
        Exit Sub
    End If
    rstNom.Close
    Set rstNom = Nothing
    
    K = "01/" & Mid(idnomina, 2, 2) & "/" & Right(idnomina, 4)
    q = DateAdd("d", -1, (DateAdd("m", 1, K)))
    
    For j = K To q: If Weekday(j, vbSunday) = 2 Then Sem = Sem + 1
    Next
    Calc = 15  'DIAS TRABAJADOS
    
    'elimina el calculo anterior de esta nómina
    Call rtnBitacora("Eliminando la información de la Nómina " & idnomina)
    cnnConexion.Execute "DELETE * FROM Nom_Detalle WHERE IDNom=" & idnomina
    '
'    'activa los empleados que regresan de vacaciones en esta nómina
'    strSQL = "UPDATE Emp SET CodEstado=0 WHERE CodEmp in " & _
'            "(SELECT CodEmp FROm Nom_Vaca WHERE IDNomina=" & IDNomina & ")"
'    cnnConexion.Execute strSQL, N
    Call rtnBitacora("Ingresando información base Nómina " & idnomina)
    
    strSQL = "INSERT INTO Nom_Detalle (IDNom,CodEmp,Sueldo,Dias_Trab,Dias_NoTrab,Bono_Noc,Bono_" _
    & "Otros,Otras_Asignaciones,SSO,SPF,LPH,Otras_Deducciones,Porc_BonoNoc) SELECT " & idnomina _
    & ",CodEmp,Sueldo," & Calc & ",0, ((sueldo / 30) * " & Calc & " * BonoNoc / 100),0,0,0,0,0," _
    & "0,BonoNoc FROM Emp WHERE CODEstado = 0 " & IIf(Inm = "", "", " AND CodInm='" & Inm & "'")
    '
    cnnConexion.Execute strSQL
    
    '
    'actualiza los dias trabajados de los trabajadores nuevo ingreso
    '
    '
    Call rtnBitacora("Actualiza dias trabajados Nuevo ingreso")
    
    K = IIf(Left(idnomina, 1) = 2, "15/", "01/") & Mid(idnomina, 2, 2) & "/" & Right(idnomina, 4)
    If q <> CDate("31/12/" & Year(Date)) Then
        q = IIf(Left(idnomina, 1) = 1, "15/" & Format(q, "mm/yy"), _
        DateAdd("d", -1, "01/" & Mid(idnomina, 2, 2) + 1 & "/" & Right(idnomina, 4)))
    End If
    
    strSQL = "UPDATE Emp INNER Join Nom_Detalle ON Emp.CodEmp = Nom_Detalle.CodEmp SET Nom_Deta" _
    & "lle.Dias_Trab = DateDiff('d',Emp.FIngreso,'" & q & "') + 1 WHERE (Emp.Fingreso Between #" _
    & Format(K, "mm/dd/yy") & "# AND #" & Format(q, "mm/dd/yy") & "#) AND Nom_Detalle.IDNom=" & idnomina
    '
    cnnConexion.Execute strSQL, N
    '
    
    'Efectua deducciones del SSO LPH SPF
    'cuando se procese la seguna nómina del mes
    'cancela las bonoficaciones fijas
    If Left(idnomina, 1) = 2 Then
        '
        'seguro social obligatorio
        Call rtnBitacora("Aplica deducción S.S.O")
        strSQL = "UPDATE Nom_Detalle INNER JOIN Emp ON Nom_Detalle.CodEmp = Emp.CodEmp SET Nom_" _
        & "Detalle.SSO = Ccur((((Emp.Sueldo + (Emp.Sueldo * Emp.BonoNoc / 100)) * 12/52) * '" & _
        pSSO & "')*" & Sem & ") WHERE Nom_Detalle.IDnom=" & idnomina & " AND Emp.SSO =True;"
        cnnConexion.Execute strSQL
        '
        'seguro paro forzoso
        Call rtnBitacora("Aplica deducción S.P.F.")
        strSQL = "UPDATE Nom_Detalle INNER JOIN Emp ON Nom_Detalle.CodEmp = Emp.CodEmp SET Nom_" _
        & "Detalle.SPF = Ccur((((Emp.Sueldo + (Emp.Sueldo * Emp.BonoNoc / 100)) * 12/52) * '" & _
        pSPF & "') * " & Sem & ") WHERE Nom_Detalle.IDnom=" & idnomina & " AND Emp.SPF = True;"
        cnnConexion.Execute strSQL
        '
        'ley política habitacional
        Call rtnBitacora("Aplica deducción L.P.H.")
        strSQL = "UPDATE Nom_Detalle INNER JOIN Emp ON Nom_Detalle.CodEmp = Emp.CodEmp SET Nom_" _
        & "Detalle.LPH = Ccur((Emp.Sueldo + (Emp.Sueldo * Emp.BonoNoc / 100)) * '" & _
        pLPH & "') WHERE Nom_Detalle.IDnom=" & idnomina & " AND Emp.LPH =True;"
        cnnConexion.Execute strSQL
        '
        'bonificaciones fijas
        Call rtnBitacora("Cancela bonificaciones fijas")
        strSQL = "UPDATE Nom_Detalle INNER JOIN Emp ON Nom_Detalle.CodEmp = Emp.CodEmp SET Nom_" _
        & "Detalle.Otras_Asignaciones = Emp.BonoFijo WHERE Nom_Detalle.IDnom=" & idnomina
        cnnConexion.Execute strSQL, N
        
    End If
    '
    If Day(K) = 15 Then K = DateAdd("D", 1, K)
    'coloca de vacaciones los empleados que salen el primer dia de la nómina
    strSQL = "UPDATE Emp SET CodEstado=3 WHERE CodEmp IN (SELECT CodEmp FROM Nom_Vaca WHERE Inicia=#" & _
    Format(K, "mm/dd/yy") & "#)"
    cnnConexion.Execute strSQL, N
    Call rtnBitacora(N & " empleados salen de vacaciones el " & K)
    
    'descuenta los dias no trabajados si el personal estaba de vacaciones
    Call rtnBitacora("Descuenta días no trabajados por regreso de vacaciones")
    'strSql = "UPDATE Nom_Detalle INNER JOIN Nom_Vaca ON Nom_Detalle.CodEmp = Nom_Vaca.CodEmp SE" _
    & "T Nom_Detalle.Dias_NoTrab = DateDiff('d','" & K & "',nom_vaca.incorpora) WHERE Nom_Va" _
    & "ca.IDNomina=" & IDNomina & " AND Nom_Detalle.IDNom=" & IDNomina & ";"
    strSQL = "UPDATE Nom_Detalle INNER JOIN Nom_Vaca ON (Nom_Detalle.IDNom = Nom_Vaca.IDNomina) AND (Nom_Detalle.CodEmp = Nom_Vaca.CodEmp) SET Nom_Detalle.Dias_NoTrab = DateDiff('d','" & K & "',nom_vaca.incorpora)WHERE (((Nom_Vaca.IDNomina)=" & idnomina & "));"

    cnnConexion.Execute strSQL, N
    'descuesta los dias no trabajados empleados q salen de vacaciones
    Call rtnBitacora("Descuenta días no trabajados por salida de vacaciones")
    
    strSQL = "UPDATE Nom_Detalle INNER JOIN Nom_Vaca ON Nom_Detalle.CodEmp = Nom_Vaca.CodEmp SE" _
    & "T Nom_Detalle.Dias_NoTrab = DateDiff('d',nom_vaca.inicia,'" & q & _
    "')+ 1,Nom_Detalle.SSO=0,Nom_Detalle.LPH=0,Nom_Detalle.SPF=0 WHERE Nom_Vaca.inicia>#" & Format(K, "mm/dd/yy") & "# and nom_vaca.inicia<=#" & _
    Format(q, "mm/dd/yy") & "# and Nom_Detalle.Idnom=" & idnomina
    
    cnnConexion.Execute strSQL, N
    'Agregar la información procesada por
    'novedades de nómina
    Call rtnBitacora("Calculos finales...")
    strSQL = "UPDATE Nom_Detalle INNER JOIN Nom_Temp ON (Nom_Detalle.IDNom = Nom_Temp.IDnomina)" _
    & " AND (Nom_Detalle.CodEmp = Nom_Temp.CodEmp) SET Nom_Detalle.Dias_libres = Nom_Detalle.Dias_libres + nom_temp.diasl" _
    & "ib, Nom_Detalle.Dias_Feriados = Nom_Detalle.Dias_Feriados + nom_temp.diasfer, Nom_Detalle.Dias_NoTrab = Nom_Detalle." _
    & "Dias_NoTrab + nom_temp.diasntrab, Nom_Detalle.Bono_Otros = Nom_Detalle.Bono_Otros + nom_temp.otrosbonos, Nom_Deta" _
    & "lle.Otras_Asignaciones = Nom_Detalle.Otras_Asignaciones + nom_temp.otrasasig, Nom_Detalle.Otras_Deducciones = Nom_Detalle.Otras_Deducciones + nom_temp.ot" _
    & "rasded, Nom_Detalle.Descuento = Nom_Detalle.Descuento + nom_temp.descuentos WHERE Nom_Detalle.IDNom=" & _
    idnomina & ";"
    '
    cnnConexion.Execute strSQL
    'UPDATE Emp INNER JOIN Nom_Detalle ON Emp.CodEmp = Nom_Detalle.CodEmp
    'SET Nom_Detalle.Bono_Noc = ((Emp.Sueldo/30)*Emp.BonoNoc/100)*(15-Nom_Detalle.Dias_NoTrab)
    'WHERE (((Nom_Detalle.IDNom)=1052005) AND ((Emp.BonoNoc)>0) AND ((Nom_Detalle.Dias_NoTrab)>0));

    'actualiza bono nocturno dias no trabajados
    strSQL = "UPDATE Emp INNER JOIN Nom_Detalle ON Emp.CodEmp = Nom_Detalle.CodEmp " _
    & "SET Nom_Detalle.Bono_Noc = ((Emp.Sueldo/30) * " & "Emp.BonoNoc/100)* (" & Calc & _
    "-Nom_Detalle.Dias_NoTrab) WHERE Emp.BonoNoc>0 AND Nom_Detalle.IDNom = " & _
    idnomina & " and Nom_Detalle.Dias_NoTrab > 0"
    '
    cnnConexion.Execute strSQL, N
    '
    End Sub


    Private Sub muestra_employes(Optional Filter As String)
    'variables locales
    
    gridNom.Rows = 2
    Call rtnLimpiar_Grid(gridNom)
    '
         
    With gridNom
        '
        For I = 3 To 4
            '
            If rstNom(I).State = 1 Then rstNom(I).Close
            rstNom(I).Open "SELECT Emp.*,Emp_Cargos.NombreCargo,Nom_Detalle.* FROM (Emp LEFT JO" _
            & "IN Emp_cargos ON Emp.CodCargo = Emp_Cargos.CodCargo) LEFT JOIN Nom_Detalle ON Em" _
            & "p.CodEmp= Nom_Detalle.CodEmp WHERE Nom_detalle.IDNom =" & idnomina & Filter & " " _
            & "ORDER BY CodInm", cnnOLEDB + gcPath + "\sac.mdb", adOpenKeyset, adLockOptimistic, adCmdText
            '
        Next
        '
        If rst.State = 1 Then rst.Close
        rst.Open rstNom(3).Source, cnnConexion, adOpenKeyset, adLockOptimistic, adCmdText
        'asigna la propiedad rowsource y datasource
        Set dtc(2).RowSource = rstNom(3)
        Set dtc(3).RowSource = rstNom(4)
        '
        For I = 0 To 27
            If Not InStr("00,04,05,06,07,08,10,12,14,15,18,19,20,22,24,27", Format(I, "00")) = 0 Then
                Set Txt(I).DataSource = rst
            End If
        Next I
        '
        If Not rst.EOF Or Not rst.BOF Then
            dtc(2) = IIf(IsNull(rst("Apellidos")), "", rst("Apellidos"))
            dtc(3) = rst.Fields("Nombres")
        End If
        '
        'CONFIGURA LA PRESENTACION DE LA INF. EN EL GRID
        If Not rstNom(3).EOF Or Not rstNom(3).BOF Then
            rstNom(3).Requery
            rstNom(3).MoveFirst: I = 0
            gridNom.Rows = rstNom(3).RecordCount + 1
            '
            Do
                I = I + 1
                .TextMatrix(I, 1) = Format(rstNom(3)("Emp.CodEmp"), "000000")
                .TextMatrix(I, 2) = IIf(IsNull(rstNom(3)("Apellidos")), "", _
                rstNom(3)("Apellidos"))
                .TextMatrix(I, 3) = rstNom(3)("Nombres")
                .TextMatrix(I, 4) = IIf(IsNull(rstNom(3)("NombreCargo")), "", _
                rstNom(3)("NombreCargo"))
                .TextMatrix(I, 5) = Format(rstNom(3)("Emp.Sueldo"), m)
                rstNom(3).MoveNext
                '
            Loop Until rstNom(3).EOF
            rstNom(3).MoveFirst
            '
        End If
        '
    End With
    '
    End Sub

