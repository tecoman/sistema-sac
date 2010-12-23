VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmConciliacion 
   Caption         =   "Conciliaciones Bancarias"
   ClientHeight    =   3135
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   4680
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   14850
   ScaleWidth      =   19080
   Tag             =   "1"
   WindowState     =   2  'Maximized
   Begin VB.Frame fraCon 
      Height          =   915
      Index           =   0
      Left            =   735
      TabIndex        =   0
      Top             =   105
      Width           =   10290
      Begin MSDataListLib.DataCombo dtcCuentas 
         DataField       =   "NumCuenta"
         Height          =   315
         Index           =   0
         Left            =   3120
         TabIndex        =   1
         Top             =   360
         Width           =   2160
         _ExtentX        =   3810
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         ListField       =   "NumCuenta"
         BoundColumn     =   "NombreBanco"
         Text            =   "DataCombo1"
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
      Begin MSDataListLib.DataCombo dtcCuentas 
         DataField       =   "NombreBanco"
         Height          =   315
         Index           =   1
         Left            =   945
         TabIndex        =   2
         Top             =   360
         Width           =   2160
         _ExtentX        =   3810
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         ListField       =   "NombreBanco"
         BoundColumn     =   "NumCuenta"
         Text            =   "DataCombo1"
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
      Begin MSMask.MaskEdBox msk 
         Bindings        =   "frmConciliacion.frx":0000
         Height          =   315
         Index           =   1
         Left            =   6615
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   330
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   556
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   12
         Format          =   "dd/MM/yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Conciliado al:"
         Height          =   330
         Index           =   1
         Left            =   5385
         TabIndex        =   4
         Top             =   390
         Width           =   1140
      End
      Begin VB.Label lbl 
         BackStyle       =   0  'Transparent
         Caption         =   "Cuenta:"
         Height          =   330
         Index           =   0
         Left            =   105
         TabIndex        =   3
         Top             =   405
         Width           =   930
      End
   End
   Begin TabDlg.SSTab ficha 
      Height          =   6450
      Index           =   0
      Left            =   720
      TabIndex        =   5
      Top             =   1260
      Width           =   10305
      _ExtentX        =   18177
      _ExtentY        =   11377
      _Version        =   393216
      TabHeight       =   520
      WordWrap        =   0   'False
      ShowFocusRect   =   0   'False
      TabCaption(0)   =   "Libro Banco"
      TabPicture(0)   =   "frmConciliacion.frx":0022
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "ficha(1)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Edo. Cta. Banco"
      TabPicture(1)   =   "frmConciliacion.frx":003E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "ficha(2)"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Conciliación"
      TabPicture(2)   =   "frmConciliacion.frx":005A
      Tab(2).ControlEnabled=   0   'False
      Tab(2).ControlCount=   0
      Begin TabDlg.SSTab ficha 
         Height          =   6090
         Index           =   1
         Left            =   60
         TabIndex        =   6
         Top             =   345
         Width           =   10215
         _ExtentX        =   18018
         _ExtentY        =   10742
         _Version        =   393216
         TabOrientation  =   2
         Tabs            =   2
         TabsPerRow      =   2
         TabHeight       =   529
         WordWrap        =   0   'False
         ShowFocusRect   =   0   'False
         BackColor       =   -2147483646
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TabCaption(0)   =   "Partidas No Registradas"
         TabPicture(0)   =   "frmConciliacion.frx":0076
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "grid(0)"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "Frame1"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).ControlCount=   2
         TabCaption(1)   =   "Errores en Libro"
         TabPicture(1)   =   "frmConciliacion.frx":0092
         Tab(1).ControlEnabled=   0   'False
         Tab(1).ControlCount=   0
         Begin VB.Frame Frame1 
            Height          =   960
            Left            =   585
            TabIndex        =   9
            Top             =   210
            Width           =   9405
            Begin VB.ComboBox cmb 
               Height          =   315
               ItemData        =   "frmConciliacion.frx":00AE
               Left            =   1320
               List            =   "frmConciliacion.frx":00BE
               Style           =   2  'Dropdown List
               TabIndex        =   12
               Top             =   480
               Width           =   900
            End
            Begin VB.TextBox txt 
               Height          =   315
               Index           =   0
               Left            =   2220
               TabIndex        =   11
               Top             =   480
               Width           =   5700
            End
            Begin VB.TextBox txt 
               Alignment       =   1  'Right Justify
               Height          =   315
               Index           =   1
               Left            =   7920
               TabIndex        =   10
               Text            =   "0,00"
               Top             =   480
               Width           =   1335
            End
            Begin MSMask.MaskEdBox msk 
               Bindings        =   "frmConciliacion.frx":00D2
               Height          =   315
               Index           =   0
               Left            =   150
               TabIndex        =   13
               TabStop         =   0   'False
               Top             =   480
               Width           =   1155
               _ExtentX        =   2037
               _ExtentY        =   556
               _Version        =   393216
               PromptInclude   =   0   'False
               MaxLength       =   12
               Format          =   "dd/MM/yyyy"
               Mask            =   "##/##/####"
               PromptChar      =   "_"
            End
            Begin VB.Label lbl 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               BorderStyle     =   1  'Fixed Single
               Caption         =   "Fecha:"
               ForeColor       =   &H80000008&
               Height          =   240
               Index           =   2
               Left            =   150
               TabIndex        =   17
               Top             =   225
               Width           =   1170
            End
            Begin VB.Label lbl 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               BorderStyle     =   1  'Fixed Single
               Caption         =   "Tipo"
               ForeColor       =   &H80000008&
               Height          =   240
               Index           =   3
               Left            =   1320
               TabIndex        =   16
               Top             =   225
               Width           =   900
            End
            Begin VB.Label lbl 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               BorderStyle     =   1  'Fixed Single
               Caption         =   "Concepto"
               ForeColor       =   &H80000008&
               Height          =   240
               Index           =   4
               Left            =   2220
               TabIndex        =   15
               Top             =   225
               Width           =   5715
            End
            Begin VB.Label lbl 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               BorderStyle     =   1  'Fixed Single
               Caption         =   "Monto"
               ForeColor       =   &H80000008&
               Height          =   240
               Index           =   5
               Left            =   7935
               TabIndex        =   14
               Top             =   225
               Width           =   1320
            End
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid grid 
            Height          =   3705
            Index           =   0
            Left            =   540
            TabIndex        =   7
            Tag             =   "1200|600|6000|1500"
            Top             =   2100
            Width           =   9465
            _ExtentX        =   16695
            _ExtentY        =   6535
            _Version        =   393216
            Cols            =   4
            FixedCols       =   0
            BackColorFixed  =   -2147483646
            ForeColorFixed  =   -2147483639
            BackColorBkg    =   -2147483636
            GridColor       =   -2147483636
            GridColorFixed  =   -2147483646
            WordWrap        =   -1  'True
            GridLinesFixed  =   1
            FormatString    =   "Fecha|Tipo|Descripción|Monto"
            _NumberOfBands  =   1
            _Band(0).Cols   =   4
            _Band(0).GridLinesBand=   1
            _Band(0).TextStyleBand=   0
            _Band(0).TextStyleHeader=   0
         End
      End
      Begin TabDlg.SSTab ficha 
         Height          =   6105
         Index           =   2
         Left            =   -74970
         TabIndex        =   8
         Top             =   330
         Width           =   10245
         _ExtentX        =   18071
         _ExtentY        =   10769
         _Version        =   393216
         TabOrientation  =   2
         Tabs            =   2
         TabsPerRow      =   2
         TabHeight       =   529
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TabCaption(0)   =   "Partidas No Registradas"
         TabPicture(0)   =   "frmConciliacion.frx":00F4
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).ControlCount=   0
         TabCaption(1)   =   "Errores"
         TabPicture(1)   =   "frmConciliacion.frx":0110
         Tab(1).ControlEnabled=   0   'False
         Tab(1).ControlCount=   0
      End
   End
End
Attribute VB_Name = "frmConciliacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
'carga el formulario
Call centra_titulo(Grid(0), True)
End Sub

