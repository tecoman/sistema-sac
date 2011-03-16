VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form FrmConsultaCxC 
   Caption         =   "Consulta de Deuda al Día de Hoy"
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
      TabIndex        =   0
      Top             =   0
      Width           =   4680
      _ExtentX        =   8255
      _ExtentY        =   847
      ButtonWidth     =   873
      ButtonHeight    =   688
      AllowCustomize  =   0   'False
      Appearance      =   1
      ImageList       =   "ImageList1"
      _Version        =   327682
      BeginProperty Buttons {0713E452-850A-101B-AFC0-4210102A8DA7} 
         NumButtons      =   6
         BeginProperty Button1 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Print"
            Object.ToolTipText     =   "Imprimir"
            Object.Tag             =   ""
            ImageIndex      =   4
         EndProperty
         BeginProperty Button2 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Mail"
            Object.ToolTipText     =   "Enviar por e-mail"
            Object.Tag             =   ""
            ImageIndex      =   3
         EndProperty
         BeginProperty Button3 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Close"
            Object.ToolTipText     =   "Salir"
            Object.Tag             =   ""
            ImageIndex      =   5
         EndProperty
         BeginProperty Button4 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button5 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "aviso"
            Object.ToolTipText     =   "Mostrar Aviso de Cobro"
            Object.Tag             =   ""
            ImageIndex      =   2
         EndProperty
         BeginProperty Button6 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "vista"
            Object.Tag             =   ""
            ImageIndex      =   6
         EndProperty
      EndProperty
      BorderStyle     =   1
      MousePointer    =   99
      MouseIcon       =   "FrmConsultaCxC.frx":0000
      Begin VB.CommandButton cmdSol 
         Caption         =   "Solvencia"
         Height          =   390
         Left            =   2595
         TabIndex        =   47
         Top             =   30
         Width           =   1215
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
            NumListImages   =   6
            BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "FrmConsultaCxC.frx":031A
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "FrmConsultaCxC.frx":0634
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "FrmConsultaCxC.frx":094E
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "FrmConsultaCxC.frx":0C68
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "FrmConsultaCxC.frx":0DEA
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "FrmConsultaCxC.frx":0F6C
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
   Begin VB.Frame FmeCuentas 
      Height          =   6660
      Left            =   400
      TabIndex        =   2
      Top             =   960
      Width           =   11220
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid FlexFacturas 
         Height          =   825
         Index           =   2
         Left            =   6975
         TabIndex        =   44
         Tag             =   "650|2300|1000"
         Top             =   1230
         Width           =   4140
         _ExtentX        =   7303
         _ExtentY        =   1455
         _Version        =   393216
         Cols            =   3
         FixedCols       =   0
         BackColorFixed  =   -2147483646
         ForeColorFixed  =   -2147483639
         BackColorSel    =   65280
         ForeColorSel    =   0
         BackColorBkg    =   -2147483636
         GridColor       =   -2147483633
         SelectionMode   =   1
         FormatString    =   "Período|Descripción|Monto"
         _NumberOfBands  =   1
         _Band(0).Cols   =   3
      End
      Begin VB.CommandButton cmd 
         Height          =   255
         Left            =   120
         Picture         =   "FrmConsultaCxC.frx":1286
         Style           =   1  'Graphical
         TabIndex        =   45
         Tag             =   "0"
         Top             =   1185
         Visible         =   0   'False
         Width           =   255
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid FlexFacturas 
         Height          =   4800
         Index           =   1
         Left            =   90
         TabIndex        =   43
         Tag             =   "1000|1400|3000|1100"
         Top             =   1155
         Width           =   6615
         _ExtentX        =   11668
         _ExtentY        =   8467
         _Version        =   393216
         Cols            =   4
         FixedCols       =   0
         BackColorFixed  =   -2147483646
         ForeColorFixed  =   -2147483639
         BackColorSel    =   65280
         ForeColorSel    =   0
         BackColorBkg    =   -2147483636
         GridColor       =   -2147483633
         AllowBigSelection=   0   'False
         SelectionMode   =   1
         AllowUserResizing=   1
         FormatString    =   "Fecha |Recibo |Descripción |Monto "
         _NumberOfBands  =   1
         _Band(0).Cols   =   4
      End
      Begin VB.TextBox Txt 
         Alignment       =   1  'Right Justify
         BeginProperty DataFormat 
            Type            =   1
            Format          =   """Bs"" #.##0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   8202
            SubFormatType   =   0
         EndProperty
         Height          =   525
         Index           =   0
         Left            =   90
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   40
         Top             =   6015
         Visible         =   0   'False
         Width           =   6630
      End
      Begin VB.Frame Frame3 
         Enabled         =   0   'False
         Height          =   1905
         Index           =   0
         Left            =   6975
         TabIndex        =   25
         Top             =   4590
         Width           =   4125
         Begin VB.TextBox Txt 
            Alignment       =   1  'Right Justify
            BeginProperty DataFormat 
               Type            =   1
               Format          =   """Bs"" #.##0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   8202
               SubFormatType   =   0
            EndProperty
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
            Index           =   10
            Left            =   1530
            TabIndex        =   28
            Text            =   " "
            Top             =   240
            Width           =   2235
         End
         Begin VB.TextBox Txt 
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
            Index           =   11
            Left            =   1530
            TabIndex        =   27
            Text            =   " "
            Top             =   1350
            Width           =   2235
         End
         Begin VB.TextBox Txt 
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
            Index           =   9
            Left            =   1530
            TabIndex        =   26
            Text            =   " "
            Top             =   750
            Width           =   2235
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   17
            Left            =   100
            TabIndex        =   31
            Top             =   285
            Width           =   1370
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
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
            Index           =   18
            Left            =   100
            TabIndex        =   30
            Top             =   1380
            Width           =   1370
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
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
            Index           =   14
            Left            =   100
            TabIndex        =   29
            Top             =   840
            Width           =   1370
         End
         Begin VB.Line Line1 
            Index           =   0
            X1              =   1530
            X2              =   3780
            Y1              =   1245
            Y2              =   1245
         End
      End
      Begin VB.Frame Frame3 
         Enabled         =   0   'False
         Height          =   1005
         Index           =   1
         Left            =   6990
         TabIndex        =   20
         Top             =   120
         Width           =   4110
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            BeginProperty DataFormat 
               Type            =   1
               Format          =   """Bs"" #,##0.00"
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
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Index           =   0
            Left            =   1920
            TabIndex        =   22
            Top             =   165
            Width           =   1995
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            BeginProperty DataFormat 
               Type            =   1
               Format          =   """Bs"" #,##0.00"
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
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Index           =   1
            Left            =   1920
            TabIndex        =   21
            Top             =   570
            Width           =   1995
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
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
            Left            =   135
            TabIndex        =   24
            Top             =   210
            Width           =   1650
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
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
            Index           =   5
            Left            =   135
            TabIndex        =   23
            Top             =   645
            Width           =   1650
         End
      End
      Begin VB.TextBox Text1 
         DataField       =   "Cobrador"
         DataSource      =   "ADOcontrol(2)"
         Height          =   360
         Index           =   12
         Left            =   7980
         TabIndex        =   3
         Top             =   1395
         Width           =   3120
      End
      Begin MSDataListLib.DataCombo Dat 
         Bindings        =   "FrmConsultaCxC.frx":13CC
         DataField       =   "AptoMovimientoCaja"
         Height          =   315
         Index           =   2
         Left            =   1590
         TabIndex        =   39
         Top             =   765
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
         Left            =   1590
         TabIndex        =   33
         Top             =   315
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
         Left            =   2655
         TabIndex        =   34
         Top             =   300
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
         Bindings        =   "FrmConsultaCxC.frx":13E1
         DataField       =   "AptoMovimientoCaja"
         Height          =   315
         Index           =   3
         Left            =   2655
         TabIndex        =   42
         Top             =   750
         Width           =   4050
         _ExtentX        =   7144
         _ExtentY        =   556
         _Version        =   393216
         ListField       =   "Nombre"
         BoundColumn     =   "Codigo"
         Text            =   " "
      End
      Begin VB.Frame Frame4 
         Enabled         =   0   'False
         Height          =   1980
         Left            =   90
         TabIndex        =   4
         Top             =   1170
         Width           =   6615
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            BeginProperty DataFormat 
               Type            =   1
               Format          =   """Bs. "" #,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   8202
               SubFormatType   =   0
            EndProperty
            DataSource      =   "ADOcontrol(2)"
            Height          =   360
            Index           =   9
            Left            =   5025
            TabIndex        =   9
            Top             =   1065
            Width           =   1500
         End
         Begin VB.TextBox Text1 
            DataSource      =   "ADOcontrol(2)"
            Height          =   360
            Index           =   7
            Left            =   5025
            TabIndex        =   8
            Top             =   645
            Width           =   1500
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            DataSource      =   "ADOcontrol(2)"
            Height          =   360
            Index           =   11
            Left            =   5025
            TabIndex        =   7
            Top             =   1500
            Width           =   1500
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "#,##0.0000"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   8202
               SubFormatType   =   1
            EndProperty
            DataSource      =   "ADOcontrol(2)"
            Height          =   360
            Index           =   10
            Left            =   1605
            TabIndex        =   6
            Top             =   1500
            Width           =   1500
         End
         Begin VB.TextBox Text1 
            DataSource      =   "ADOcontrol(2)"
            Height          =   360
            Index           =   8
            Left            =   1605
            TabIndex        =   5
            Top             =   1065
            Width           =   1500
         End
         Begin MSMask.MaskEdBox MskTelefono 
            Height          =   360
            Index           =   1
            Left            =   1605
            TabIndex        =   10
            Top             =   210
            Width           =   1500
            _ExtentX        =   2646
            _ExtentY        =   635
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   16
            Format          =   "(####)-###-##-##"
            Mask            =   "(####)-###-##-##"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MskTelefono 
            Height          =   360
            Index           =   2
            Left            =   1605
            TabIndex        =   11
            Top             =   630
            Width           =   1500
            _ExtentX        =   2646
            _ExtentY        =   635
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   16
            Format          =   "(####)-###-##-##"
            Mask            =   "(####)-###-##-##"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MskTelefono 
            Height          =   360
            Index           =   0
            Left            =   5025
            TabIndex        =   41
            Top             =   225
            Width           =   1500
            _ExtentX        =   2646
            _ExtentY        =   635
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   16
            Format          =   "(####)-###-##-##"
            Mask            =   "(####)-###-##-##"
            PromptChar      =   "_"
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
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
            Left            =   3195
            TabIndex        =   19
            Top             =   1125
            Width           =   1725
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
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
            Left            =   105
            TabIndex        =   18
            Top             =   1125
            Width           =   1455
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
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
            Left            =   3195
            TabIndex        =   17
            Top             =   705
            Width           =   1725
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
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
            Left            =   105
            TabIndex        =   16
            Top             =   705
            Width           =   1455
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
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
            Left            =   3195
            TabIndex        =   15
            Top             =   270
            Width           =   1725
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
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
            Left            =   105
            TabIndex        =   14
            Top             =   270
            Width           =   1455
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
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
            Left            =   3195
            TabIndex        =   13
            Top             =   1560
            Width           =   1725
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
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
            Left            =   105
            TabIndex        =   12
            Top             =   1560
            Width           =   1455
         End
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid FlexFacturas 
         Height          =   3270
         Index           =   0
         Left            =   90
         TabIndex        =   46
         Tag             =   "1000|700|1200|1200|1200|1200"
         Top             =   3255
         Width           =   6615
         _ExtentX        =   11668
         _ExtentY        =   5768
         _Version        =   393216
         Cols            =   6
         FixedCols       =   0
         BackColorFixed  =   -2147483646
         ForeColorFixed  =   -2147483639
         BackColorSel    =   65280
         ForeColorSel    =   0
         BackColorBkg    =   -2147483636
         GridColor       =   -2147483633
         SelectionMode   =   1
         FormatString    =   "Factura |Período |Facturado |Abonado |Saldo |Acumulado"
         _NumberOfBands  =   1
         _Band(0).Cols   =   6
      End
      Begin VB.Label lbl 
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
         Left            =   225
         TabIndex        =   38
         Top             =   795
         Width           =   1000
      End
      Begin VB.Label lbl 
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
         Left            =   465
         TabIndex        =   32
         Top             =   375
         Width           =   1005
      End
      Begin VB.Label lbl 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000002&
         Height          =   2205
         Index           =   20
         Left            =   7005
         TabIndex        =   37
         Top             =   2160
         Width           =   4095
      End
      Begin VB.Label lbl 
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
         Left            =   6975
         TabIndex        =   36
         Top             =   1440
         Width           =   855
      End
      Begin VB.Label lbl 
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
         Index           =   6
         Left            =   7005
         TabIndex        =   35
         Top             =   1875
         Width           =   1050
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   7365
      Left            =   120
      TabIndex        =   1
      Top             =   525
      Width           =   11670
      _ExtentX        =   20585
      _ExtentY        =   12991
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   4
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "&Deuda"
      TabPicture(0)   =   "FrmConsultaCxC.frx":13F6
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Image1(0)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Image1(1)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "P&agos"
      TabPicture(1)   =   "FrmConsultaCxC.frx":1412
      Tab(1).ControlEnabled=   0   'False
      Tab(1).ControlCount=   0
      Begin VB.Image Image1 
         Height          =   240
         Index           =   1
         Left            =   -180
         Picture         =   "FrmConsultaCxC.frx":142E
         Top             =   30
         Width           =   240
      End
      Begin VB.Image Image1 
         Height          =   240
         Index           =   0
         Left            =   -195
         Picture         =   "FrmConsultaCxC.frx":1503
         Top             =   180
         Width           =   240
      End
   End
End
Attribute VB_Name = "FrmConsultaCxC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
    '---------------------------------------------------------------------------------------------
    '   Módulo Consulta Cuentas por Cobrar
    '
    '   Muestra un listado de la deuda de c/cliente, detalle de facturas pendientes
    '   facturas canceladas y forma de pago.
    '---------------------------------------------------------------------------------------------
    Dim objRst As ADODB.Recordset
    Dim ObjRstP(2) As ADODB.Recordset
    Dim cnnPropietario As ADODB.Connection
    Dim StrRutaInmueble$, strPago$
    Dim IntHonoMorosidad As Integer
    Dim datIndex As Integer
    Dim vecFPS(20, 3) As Currency
    Dim mlEdit As Boolean, mlNew As Boolean
    Private IntMMora As Integer
    Private tipoCta As Byte
    '---------------------------------------------------------------------------------------------

    Private Sub cmd_Click()
    'variables locales
    If cmd.Tag = 0 Then
        cmd.Tag = 1
        FlexFacturas(1).Width = 11025
        cmd.Picture = Image1(1)
        FlexFacturas(2).Visible = False
        Frame3(0).Visible = False
        If FlexFacturas(1).Rows > 18 Then
            FlexFacturas(1).ColWidth(2) = 7200
        Else
            FlexFacturas(1).ColWidth(2) = 7300
        End If
    Else
        cmd.Tag = 0
        FlexFacturas(1).Width = 6615
        cmd.Picture = Image1(0)
        FlexFacturas(2).Visible = True
        Frame3(0).Visible = True
        If FlexFacturas(1).Rows > 18 Then
            FlexFacturas(1).ColWidth(2) = 2800
        Else
            FlexFacturas(1).ColWidth(2) = 3000
        End If
    End If
    FlexFacturas(1).SetFocus
    '
    End Sub

Private Sub cmdSol_Click()
'emision de solvencia de condominio
'arreglo con la información del propietario
' Elemento - Contenido
' 0 = Apto
' 1 = Cancelacion
' 2 = facturacion
' 3 = Inmueble
' 4 = Propietario
' 5 = Usuario
'-------------------------------------------
Dim rstlocal As ADODB.Recordset
Dim sFacturado As String, sCancelado As String
Dim n%
Set rstlocal = New ADODB.Recordset
rstlocal.CursorLocation = adUseClient
rstlocal.Open "factura", cnnOLEDB + gcPath & "\" & Dat(0) & "\inm.mdb", adOpenStatic, _
adLockReadOnly, adCmdTable
rstlocal.Filter = "codprop='" & Dat(2) & "'"
If Not rstlocal.EOF And Not rstlocal.BOF Then
    rstlocal.Sort = "Periodo DESC"
    sFacturado = UCase(Format(rstlocal!Periodo, "MMMM-YYYY") & " (" & rstlocal!FechaFactura & ")")
    
    Do
        If rstlocal!Saldo < 0.009 Then
            sCancelado = UCase(Format(rstlocal!Periodo, "MMMM-YYYY") & " (" & Text1(8) & ")")
            Exit Do
        End If
        
        rstlocal.MoveNext
        n = n + 1
        If n > 1 Then
            MsgBox "No se puede emitir solvencia: " & vbCrLf & vbCrLf & "Propietario tiene más de 1 recibo pendiente.", vbInformation, App.ProductName
            Exit Do
        End If
    Loop
    If n < 2 Then Call emitir_solvencia(Dat(2), sCancelado, sFacturado, Dat(1), Dat(3), gcUsuario, Dat(0))
Else
    MsgBox "Imposible emitir la solvencia. No se encuentra la" & vbCrLf & _
    "información del propietario '" & Dat(2) & "'", vbInformation, App.ProductName
End If
rstlocal.Close
Set rstlocal = Nothing
End Sub

    Private Sub Dat_Click(Index As Integer, Area As Integer)
    '
    If Area = 2 Then
        datIndex = Index
        txt(0) = "": strPago = ""
        Select Case Index
            '
            'LISTA DE CODIGOS DE INMUEBLE
            Case 0: Call RtnBuscaInmueble("Inmueble.CodInm", Dat(0), Dat(2))
            'LISTA DE NOMBRES DE INMUEBLES
            Case 1: Call RtnBuscaInmueble("Inmueble.Nombre", Dat(1), Dat(2))
            
            Case 2, 3 'LISTA DE CODIGOS DE PROPIETARIOS
                Dat(IIf(Index = 3, 2, 3)) = Dat(Index).BoundText
                
                Call RtnPropietario("Codigo", Dat(2).Text, True)
                If FlexFacturas(0).Rows > 8 Then _
                FlexFacturas(0).TopRow = FlexFacturas(0).Rows - 8
            
'            Case 3  'LISTA DE NOMBRES PROPIETARIOS
'                'Call RtnPropietario("Nombre", Dat(3).Text)
'
'                If FlexFacturas(0).Rows > 8 Then FlexFacturas(0).TopRow = FlexFacturas(0).Rows - 8
            '
        End Select
        '
    End If
    '
    End Sub
    '
    
    
    Private Sub Dat_KeyPress(Index As Integer, KeyAscii As Integer)
    '
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    '
    If KeyAscii = 13 Then
        txt(0) = "": strPago = ""
        
        Select Case Index
            
            Case 0  'CONTROL CODIGOS DE INMUEBLES FICHA DEUDA
                If Dat(0) = "" Then Dat(1).SetFocus: Exit Sub
                Call RtnBuscaInmueble("Inmueble.CodInm", Dat(0), Dat(2))
                            
            Case 1  'DAT(1) NOMBRE DE INMUEBLE FICHA DEUDA
                If Dat(1) = "" Then Dat(0).SetFocus: Exit Sub
                Call RtnBuscaInmueble("Inmueble.Nombre", Dat(1), Dat(2))
                
            Case 2  'DAT(2) CODIGOS DE PROPIETARIOS FICHA DEUDA
                If Dat(2) = "" Then Dat(3).SetFocus: Exit Sub
                Call RtnPropietario("Codigo", Dat(2))
                If FlexFacturas(0).Rows > 8 Then FlexFacturas(0).TopRow = FlexFacturas(0).Rows - 8
                
            Case 3  'LISTA DE NOMBRES PROPIETARIOS
                If Dat(3) = "" Then Dat(2).SetFocus: Exit Sub
                Call RtnPropietario("Nombre", Dat(3))
                If FlexFacturas(0).Rows > 8 Then FlexFacturas(0).TopRow = FlexFacturas(0).Rows - 8
                        
        End Select
    '
    End If
    '
    End Sub


    'Rev.23/08/2002-------------------------------------------------------------------------------
    Private Sub FlexFacturas_Click(Index As Integer) '-
    '---------------------------------------------------------------------------------------------
    '
    Select Case Index
        Case 1: Call RtnPagos   'Flex Pagos
    '
    End Select
    '
    End Sub



    Private Sub FlexFacturas_KeyPress(Index As Integer, KeyAscii As Integer)
    '
    Select Case Index
        Case 1: Call RtnPagos   'Flex Pagos
    '
    End Select
    '
    End Sub

    

    '-------------------------'DESCARGA DEL ARCHIVO DE RECURSOS LAS CADENAS*
    Private Sub Form_Load() '-'DE TEXTO PROPIEDAD CAPTION DE CADA ETIQUETA**
    '-----------------------------------------------------------------------
    establecerFuente FrmConsultaCxC
    Dim I As Integer
    lbl(0) = LoadResString(101):        lbl(1) = LoadResString(112)
    lbl(2) = LoadResString(107):        lbl(3) = LoadResString(102)
    lbl(4) = LoadResString(116):        lbl(5) = LoadResString(108)
    lbl(6) = LoadResString(118):        lbl(7) = LoadResString(109)
    lbl(8) = LoadResString(113):        lbl(9) = LoadResString(110)
    lbl(10) = LoadResString(114):       lbl(11) = LoadResString(111)
    lbl(12) = LoadResString(115):       lbl(13) = LoadResString(117)
    lbl(17) = LoadResString(119):       lbl(14) = LoadResString(120)
    lbl(18) = LoadResString(121):       Dat(3) = ""
    'cmd.Picture = LoadResPicture("DER", vbResIcon)
    For I = 0 To 2
        Dat(I) = ""
'        Set FlexFacturas(I).FontFixed = LetraTitulo(LoadResString(527), 7.5)
'        Set FlexFacturas(I).Font = LetraTitulo(LoadResString(528), 7.5)
        FlexFacturas(I).FontWidth = 3.5
        Call centra_titulo(FlexFacturas(I), True)
    Next
    FlexFacturas(1).ColAlignment(0) = flexAlignCenterCenter
    FlexFacturas(2).ColAlignment(0) = flexAlignLeftCenter
    '
    FlexFacturas(1).Visible = False
    FlexFacturas(2).Visible = False
    '
    Set Dat(0).RowSource = FrmAdmin.objRst
    Set Dat(1).RowSource = FrmAdmin.ObjRstNom
    '
    For I = 21 To 24
        Load lbl(I)
        lbl(I).Visible = True
        lbl(I).Width = Frame3(0).Width
        lbl(I) = ""
        lbl(I).Font.Size = lbl(I).Font.Size - 1
    Next
    Dim margen_top As Long
    margen_top = Screen.Height / Screen.TwipsPerPixelY
    margen_top = (margen_top * 300) / 768
    lbl(21).FontBold = True
    lbl(21).Left = lbl(20).Left
    lbl(21).Top = Frame3(0).Top + Frame3(0).Height + 300
    lbl(21).AutoSize = True
    lbl(22).Left = lbl(21).Left
    lbl(22).Top = lbl(21).Top + 300
    lbl(23).Left = lbl(21).Left
    lbl(23).Top = lbl(22).Top + 300
    lbl(24).Left = lbl(21).Left
    lbl(24).Top = lbl(23).Top + 300
    End Sub


    Private Sub Form_Resize()
    'adapta el formulario a la ventana
    'On Error Resume Next
    With SSTab1
    '
        If FrmAdmin.WindowState <> vbMinimized Then
            .Left = (Me.ScaleWidth - .Width) / 2
            .Height = Me.ScaleHeight - .Top - 200
            FmeCuentas.Left = .Left + 280
            FmeCuentas.Height = .Height - FmeCuentas.Top + 300
            FlexFacturas(0).Height = FmeCuentas.Height - FlexFacturas(0).Top - 100
            FlexFacturas(1).Height = FmeCuentas.Height - FlexFacturas(1).Top - 725
            txt(0).Top = FlexFacturas(1).Top + FlexFacturas(1).Height
            FlexFacturas(2).Height = FmeCuentas.Height - FlexFacturas(2).Top - 200
            lbl(20).Height = FmeCuentas.Height * 0.3
            Frame3(0).Top = lbl(20).Top + lbl(20).Height
            Frame3(0).Left = lbl(20).Left
        End If
    End With
    '
    End Sub

    Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    cnnPropietario.Close
    Set cnnPropietario = Nothing
    For I = 0 To 2
        ObjRstP(I).Close
        Set ObjRstP(I) = Nothing
    Next I
    End Sub



Private Sub lbl_Click(Index As Integer)
Dim sContacto As String, sGestion As String, sPor As String, sCadena As String
If Index >= 22 And Index <= 24 Then
    lbl(Index).Font.Underline = False
    lbl(Index).ForeColor = vbBlack
    sCadena = lbl(Index).Tag
    sContacto = Trim(Left(sCadena, InStr(sCadena, vbCrLf)))
    sCadena = Right(sCadena, Len(sCadena) - InStr(sCadena, vbLf))
    sPor = Trim(Left(sCadena, InStr(sCadena, vbCrLf)))
    sGestion = Mid(sCadena, InStr(sCadena, vbLf) + 1, _
    Len(sCadena))
    
    addToolTip sContacto, sGestion, sPor
End If
End Sub

Private Sub lbl_MouseMove(Index As Integer, Button As Integer, _
Shift As Integer, X As Single, Y As Single)
'Y = 30 - 150
'X = 15 - 3750
For I = 22 To 24
    lbl(I).Font.Underline = False
    lbl(I).ForeColor = vbBlack
Next
'toolTip.Visible = False
If Index >= 22 And Index <= 24 Then
    If (Y >= 30 And Y <= 150) And (X >= 25 And X <= 3750) Then
        lbl(Index).Font.Underline = True
        lbl(Index).ForeColor = vbBlue
        'toolTip = addToolTip(lbl(Index).Tag)
        'toolTip.Move lbl(Index).Left + 100, _
        lbl(Index).Top - toolTip.Height '_
        'lbl(Index).Width -150, CLng(TextWidth(toolTip) / 4000) * 450
        'toolTip.Visible = True
        
    End If
End If
'Debug.Print X, Y
End Sub

    'rEV.23/08/2002------Eventos que suceden al seleccionar un ficha------------------------------
    Private Sub SSTab1_Click(PreviousTab As Integer) '-
    '---------------------------------------------------------------------------------------------
    '
    Select Case SSTab1.tab
    '
        'FICHA 'PAGOS'
        Case 1: Call RtnclickFicha(True, 1)
        'FICHA 'DEUDA'
        Case 0: Call RtnclickFicha(False, 0)
    '
    End Select
    End Sub

    '---------------------------------------------------------------------------------------------
    Private Sub Toolbar1_ButtonClick(ByVal Button As ComctlLib.Button)
    '---------------------------------------------------------------------------------------------
    'variables locales
    Dim errLocal As Long
    Dim intRecibos As Integer
    Dim P As Date, TempData(6) As String
    Dim rpReporte As ctlReport
    '
    Select Case UCase(Button.Key)
    '
        Case "PRINT", "MAIL" 'BOTON IMPRIMIR
    '   -----------------------------------------
        'variables locales
        Dim mdb As DBEngine
        Dim dbSac As Database
        Dim TDF As TableDef
        Dim cnnTemp As New ADODB.Connection
        ''Dim blnExiste As Boolean
        
        '-----------------------------
        If SSTab1.tab = 0 Then
            If Trim(txt(10)) = "" Then txt(10) = 0
            If CCur(txt(10)) <> 0 Then
            MousePointer = vbHourglass
                'Set ctlReport = FrmAdmin.rptReporte
                Set edoRpt = New ctlReport
                'With ctlReport
                With edoRpt
                    '.ReportFileName = gcReport + "EdoCtaPro.rpt"
                    .Reporte = gcReport + "edoctapro.rpt"
                    '.Reset
                    '.ReportFileName = gcReport + "EdoCtaPro.rpt"
                    '.DataFiles(0) = gcPath & StrRutaInmueble & "inm.mdb"
                    .OrigenDatos(0) = gcPath & StrRutaInmueble & "inm.mdb"
                    '.DataFiles(1) = gcPath & StrRutaInmueble & "inm.mdb"
                    .OrigenDatos(1) = gcPath & StrRutaInmueble & "inm.mdb"
                    '.SortFields(0) = "+{Factura.periodo}"
                    .Formulas(0) = "Inmueble='" & Dat(1) & "'"
                    '.Formulas(1) = "Inmueble ='" & Dat(1) & "'"
                    .Formulas(1) = "MesesMora=" & intMora
                    '.Formulas(2) = "MesMora =" & IntMMora
                    .Formulas(2) = "Morosidad=" & IntHonoMorosidad
                    '.Formulas(3) = "Morosidad =" & IntHonoMorosidad
                    '.SelectionFormula = "{Factura.Saldo} > 0 AND {Propietarios.Codigo}='" _
                    & Dat(2) & "'"
                    .FormuladeSeleccion = "{Factura.Saldo} > 0 AND {Propietarios.Codigo}='" _
                    & Dat(2) & "'"
                    .Salida = crPantalla
                    .TituloVentana = "Estado de Cuenta " & Dat(0) & "/" & Dat(2)
                    .Imprimir
'                    If UCase(Button.Key) = "MAIL" Then
'
'                        If Text1(7) = "" Then
'
'                            MsgBox "Propietario No tien e-mail registrado", vbInformation, _
'                            App.ProductName
'                            MousePointer = vbDefault
'                            Exit Sub
'
'                        End If
'                        'llama la rutina
'                        '.Destination = crptMapi
'                        '.PrintFileType = crptRTF
'                        '.EMailToList = Trim(Text1(7))
'                        '.EMailMessage = "Adjunto Estado de Cuenta al " & Date
'                        '.EmailSubject = App.CompanyName & " - Estado de Cuenta"
'                        'Si el propietario está en legal no muestra el detalle de la deuda
''                        If intRecibos > 3 Then
''
''                            '.SectionFormat(0) = "DETAIL;F;X;X;X;X;X;X"
''                            '.SectionFormat(1) = "GF1;F;X;X;X;X;X;X"
''
''                        End If
'                        '
'                    Else
'                        '.Destination = crptMapi
'                        '.EMailToList = "ynfantes@cantv.net"
'
''                        .Destination = crptToWindow
''                        .WindowParentHandle = FrmAdmin.hWnd
''                        .WindowShowCloseBtn = True
''                        .WindowState = crptMaximized
''                        .WindowTitle = "Estado de Cuenta " & Dat(0) & " / " & Dat(2)
'                        '.SectionFormat(0) = "DETAIL;T;X;X;X;X;X;X"
'                        '.SectionFormat(1) = "GF1;T;X;X;X;X;X;X"
'                    End If
                    'errlocal = .PrintReport
                    Call rtnBitacora("Impresión Estado de Cuenta " & Dat(0) & "/" & Dat(2))
                    'If errlocal <> 0 Then MsgBox .LastErrorString, vbCritical, .LastErrorNumber
                    '
                End With
                MousePointer = vbNormal
            Else
                MsgBox "Deuda del Cliente en Cero(0)", vbInformation, App.ProductName
            End If
        Else    'Imprime pagos del propietario
            Call Print_Pagos
        End If
        'BOTON SALIR
        '-----------
        Case "CLOSE"
            Me.Hide
            EliminaFrmGestion
            
        Case "AVISO"    'visualiza el aviso de cobro
            
            If FlexFacturas(0).TextMatrix(FlexFacturas(0).Row, 0) <> "" Then
                Dim dPeriodo As Date
                Dim frmLocal As frmAC
                
                dPeriodo = CDate("01/" & FlexFacturas(0).TextMatrix(FlexFacturas(0).Row, 1))
                'esta caracteristica comienza a partir del 01/08/2005
                'para próximas instalaciones eliminar esta interrogante
                If dPeriodo > CDate("01/07/2005") Then
                    'emisión del aviso de cobro en formato .pdf
                    'comienza en la facturaciónd de agosto
                    Dim m_report As CRAXDRT.Report
                    Dim m_app As CRAXDRT.Application
                    Dim sArchivo As String
                    
                    '
                    Set m_report = New CRAXDRT.Report
                    Set m_app = New CRAXDRT.Application
                    
                    Set frmLocal = New frmAC
                    frmLocal.strEmail = Trim(Text1(7))
                    
                    sArchivo = gcPath & "\" & Dat(0) & "\reportes\AC" & _
                    Format(dPeriodo, "MMMYY") & ".rpt"
                    If Dir$(sArchivo, vbArchive) = "" Then
                        If Not IsNumeric(FlexFacturas(0).TextMatrix(FlexFacturas(0).Row, 0)) Then
                            MsgBox "No se puede mostrar la información solicitada." & vbCrLf & _
                            "El período no corresponde a una factura", vbInformation, App.ProductName
                        Else
                            MsgBox "No se puede mostrar este aviso de cobro." & vbCrLf & _
                            "El archivo no existe.", vbInformation, App.ProductName
                        End If
                        Exit Sub
                    End If
                    Set m_report = m_app.OpenReport(sArchivo, 1)
                    m_report.RecordSelectionFormula = "{AC.Codigo}='" & Dat(2) & "'"
                    'm_report.ExportOptions.FormatType = crEFTPortableDocFormat
                    
                    frmLocal.Caption = "Mes:" & Format(dPeriodo, "MM-YYYY")
                    frmLocal.crView.ReportSource = m_report
                    frmLocal.crView.DisplayGroupTree = False
                    frmLocal.crView.EnableExportButton = True
                    frmLocal.crView.EnableZoomControl = True
                    frmLocal.crView.ViewReport
                    
                    If Screen.Width / Screen.TwipsPerPixelX = 1024 Then
                        frmLocal.crView.Zoom (120)
                    Else
                        frmLocal.crView.Zoom (92)
                    End If
                    'frmLocal.Show
                    Exit Sub
                End If
                sArchivo = gcPath & "\" & Dat(0) & "\Reportes\" & FlexFacturas(0).TextMatrix(FlexFacturas(0).Row, 0) & ".html"
                
                If Dir(sArchivo) = "" Then
                    
                    P = "01/" & FlexFacturas(0).TextMatrix(FlexFacturas(0).Row, 1)
                    P = Format(P, "mm/dd/yy")
                    TempData(0) = mcDatos
                    TempData(1) = gcUbica
                    TempData(2) = gcCodInm
                    TempData(3) = gcNomInm
                    TempData(4) = gnPorIntMora
                    TempData(5) = gnMesesMora
                    TempData(6) = gnCta
                    gcUbica = "\" & Dat(0) & "\"
                    mcDatos = gcPath & gcUbica & "inm.mdb"
                    gcCodInm = Dat(0)
                    gcNomInm = Dat(1)
                    gnPorIntMora = IntHonoMorosidad
                    gnMesesMora = IntMMora
                    gnCta = tipoCta
                    
                    Call Enviar_ACemail(P, Dat(2), False, True)
                    Call rtnBitacora("Aviso de Cobro: " & Dat(0) & "/" & Dat(2) & " Mes:" & _
                    FlexFacturas(0).TextMatrix(FlexFacturas(0).Row, 1))
                    mcDatos = TempData(0)
                    gcUbica = TempData(1)
                    gcCodInm = TempData(2)
                    gcNomInm = TempData(3)
                    gnPorIntMora = TempData(4)
                    gnMesesMora = TempData(5)
                    gnCta = TempData(6)
                End If
                'mostrar el archivo
                Set frmLocal = New frmAC
                frmLocal.strArchivo = sArchivo 'gcPath & "\" & Dat(0) & "\Reportes\" & _
                FlexFacturas(0).TextMatrix(FlexFacturas(0).Row, 0) & ".html"
                frmLocal.crView.Visible = False
                frmLocal.Caption = "Mes:" & Format(dPeriodo, "MM-YYYY")
                frmLocal.strEmail = Trim(Text1(7))
                frmLocal.Show
            Else
                MsgBox "Seleccione una factura de la lista", vbInformation, App.ProductName
            End If
            
        Case "VISTA"
            FlexFacturas(0).Visible = Not FlexFacturas(0).Visible
        '
    End Select
    '
    End Sub

    'Rev23/08/2002--------------------------------------------------------------------------------
    Sub RtnListapro(StrSource$) 'Llena la Lista con los codigos de Propietarios del Inm
    '---------------------------------------------------------------------------------------------
    '
    If StrRutaInmueble = "" Then
        Set Dat(2).RowSource = Nothing
        Set Dat(3).RowSource = Nothing
        Exit Sub
    End If
    If Dir(gcPath + StrRutaInmueble + "inm.mdb") = "" Then
        MsgBox "Consulte el administrador del sistema, no se consigue el archivo inm.mdb", _
        vbInformation, App.ProductName
        Exit Sub
    End If
    Set cnnPropietario = New ADODB.Connection
    '---------------------------------------------------------------------------------------------
    cnnPropietario.Open cnnOLEDB + gcPath + StrRutaInmueble + "inm.mdb"
    '---------------------------------------------------------------------------------------------
    For I = 0 To 1  'Asigna los valores a las lista de los controles ordenados s/ control
        Set ObjRstP(I) = New ADODB.Recordset
        ObjRstP(I).Open "SELECT * FROM Propietarios ORDER BY " _
        & IIf(I = 2, "Nombre", "Codigo"), cnnPropietario, _
        adOpenStatic, adLockReadOnly, adCmdText
        Set Dat(I + 2).RowSource = ObjRstP(I)
        Dat(I + 2).Refresh
    Next
    '
    End Sub

    '---------------------------------------------------------------------------------------------
    Sub RtnAsignaInf(DataC As DataCombo) '-
    '---------------------------------------------------------------------------------------------
    'On Error Resume Next
    With objRst
    '
        If objRst.EOF Then
            StrRutaInmueble = ""
            IntMMora = 0
            IntHonoMorosidad = 0
    
            Exit Sub
        End If
        Dat(0).Text = .Fields("CodInm"): Dat(1) = .Fields("Nombre")
        StrRutaInmueble = .Fields("Ubica")
        IntMMora = .Fields("MesesMora")
        IntHonoMorosidad = .Fields("HonoMorosidad")
        If !Caja = sysCodCaja Then
            tipoCta = CUENTA_POTE
        Else
            tipoCta = CUENTA_INMUEBLE
        End If
        Text1(0).Text = Format(CCur(.Fields("Deuda")), "#,##0.00")
        Text1(1).Text = Format(CCur(IIf(IsNull(.Fields("FondoAct")), 0, .Fields("FondoAct"))), _
            "#,##0.00")
        DataC.SetFocus
        '
    End With
    'CIERRA EL RECORSET Y LA CONEXION DE DATOS
    objRst.Close
    Set objRst = Nothing
    End Sub

    '---------------------------------------------------------------------------------------------
    Sub RtnClear() 'AL SELECCIONAR UN INMUEBLE LIMPIA LOS CONTROLES-
    '---------------------------------------------------------------------------------------------
    '
    For I = 0 To 2
        FlexFacturas(I).Rows = 2
        Call rtnLimpiar_Grid(FlexFacturas(I))
        MskTelefono(I) = ""
    Next
    '
    Dat(2) = ""
    Dat(3) = ""
    lbl(20) = ""
    '
    For I = 4 To 11
        If I = 4 Or I = 6 Or I = 5 Then GoTo 10
        Text1(I) = ""
10      Next
    For I = 9 To 11: txt(I) = Format(0, "#,##0.00")
    Next
    '
    End Sub
 
    'Rev.23/08/2002-------------------------------------------------------------------------------
    
    Sub RtnFlexPagos() 'distribuye en el Grid Los Pagos efectuados por un cliente específico
    '---------------------------------------------------------------------------------------------
    'VARIABLES LOCALES
    Dim rPagos As New ADODB.Recordset
    '
    'llena el ADODB.Recordset con la información de los pago efectuados por el cliente
    rPagos.Open "SELECT FechaMovimientoCaja, IDRecibo,CuentaMovimientoCaja, DescripcionMovimien" _
    & "toCaja, MontoMovimientoCaja FROM MovimientoCaja WHERE InmuebleMovimientoCaja= '" & Dat(0) _
    & "' AND AptoMovimientoCaja='" & Dat(2) & "' ORDER BY FechaMovimientoCaja DESC,IDRecibo DES" _
    & "C", cnnConexion, adOpenKeyset, adLockOptimistic, adCmdText
    '
    If Not rPagos.EOF Or Not rPagos.BOF Then
    '
        FlexFacturas(1).Rows = rPagos.RecordCount + 1
        If FlexFacturas(1).Rows > 18 Then
            If cmd.Tag = 0 Then FlexFacturas(1).ColWidth(2) = 2800
            If cmd.Tag = 1 Then FlexFacturas(1).ColWidth(2) = 7200
        Else
            If cmd.Tag = 0 Then FlexFacturas(1).ColWidth(2) = 3000
            If cmd.Tag = 1 Then FlexFacturas(1).ColWidth(2) = 7300
        End If
        I = 1: rPagos.MoveFirst
        Do
        '
            With FlexFacturas(1)
        '
                .TextMatrix(I, 0) = IIf(IsNull(rPagos("FechaMovimientoCaja")), "", _
                rPagos("FechaMovimientoCaja"))
                .TextMatrix(I, 1) = IIf(IsNull(rPagos("IDRecibo")), "", rPagos("IDRecibo"))
                .TextMatrix(I, 2) = rPagos("CuentaMovimientoCaja") + " " + _
                rPagos("DescripcionMovimientoCaja")
                .TextMatrix(I, 3) = Format(rPagos("MontoMovimientoCaja"), "##,##0.00")
                rPagos.MoveNext
                I = I + 1
        '
            End With
        '
        Loop Until rPagos.EOF
    End If
    '
    rPagos.Close
    Set rPagos = Nothing  'cierra y destruye el ADODB.Recordset
    '
    End Sub

    '---------------------------------------------------------------------------------------------
    Sub RtnBuscaInmueble(Qdf$, DC1 As DataCombo, DC2 As DataCombo)
    '---------------------------------------------------------------------------------------------
    '
    Call RtnClear
    Set objRst = New ADODB.Recordset
    objRst.Open "SELECT * FROM Inmueble WHERE " & Qdf & " Like '%" & DC1.Text & "%'", _
        cnnConexion, adOpenStatic, adLockReadOnly, adCmdText
    Call RtnAsignaInf(DC2)
    Call RtnListapro("SELECT Nombre, Codigo, Deuda, Notas,TelfHab, Fax, Celular, Email, FecUltP" _
    & "ag, UltPago,Alicuota, Recibos FROM Propietarios ORDER BY Propietarios.Codigo")
    '
    End Sub

    '---------------------------------------------------------------------------------------------
    Private Sub RtnVisible(StrEstado As Boolean)
    '---------------------------------------------------------------------------------------------
    '
    FlexFacturas(0).Visible = Not StrEstado
    FlexFacturas(0).Visible = Not StrEstado
    txt(0).Visible = StrEstado
    FlexFacturas(1).Visible = StrEstado
    cmd.Visible = StrEstado
    FlexFacturas(2).Visible = StrEstado
    Frame3(0).Visible = Not StrEstado
    '
    End Sub

    '---------------------------------------------------------------------------------------------
    Private Sub RtnPropietario(StrCampo$, StrControl$, Optional Busca As Boolean) '-
    '---------------------------------------------------------------------------------------------
    'RUTINA QUE BUSCA LA INFORMACION DEL PROPIETARIO SELECCIONADO
    'BUSCA LA DEUDA/PAGOS DE ACUERDO A LOS PARAMETROS DEL USUARIO
    '---------------------------------------------------------------------------------------------
    Dim rstGestion As ADODB.Recordset
    Dim strTool As String
    'dim tl as tool
    If StrControl = "" Or StrRutaInmueble = "" Then Exit Sub
    Set ObjRstP(0) = New ADODB.Recordset    'Variables locales
    ObjRstP(0).Open "Propietarios", cnnPropietario, adOpenKeyset, adLockOptimistic, adCmdTable
    '
    With ObjRstP(0)
        '
        If Not (.EOF And .BOF) Then
            If Not Busca Then
                .MoveFirst
                StrControl = Replace(StrControl, "'", "''")
                .Find StrCampo & " LIKE '*" & StrControl & "*'" 'BUSCA COINCIDENCIA
            Else
                .Filter = "Codigo='" & Dat(2) & "' AND Nombre='" & Replace(Dat(3), "'", "''") & "'"
            End If
            If Not .EOF Then                    'SI LA CONSIGUE
                Dat(2).Text = !Codigo           'ASGINA VALORES A LOS CONTROLES
                'If Not datIndex = 3 Then Dat(3).Text = !Nombre
                If Not Busca Then Dat(3).Text = !Nombre
                'muestra las trés ultimas gestiones de cobranza
                '(si las tiene)
                Set rstGestion = New ADODB.Recordset
                With rstGestion
                    .CursorLocation = adUseClient
                    .Open "Gestion", cnnPropietario, adOpenKeyset, adLockOptimistic, adCmdTable
                    .Filter = "Apto='" & Dat(2) & "'"
                    .Sort = "Fecha DESC"
                    If Not (.EOF And .BOF) Then
                        lbl(21) = "Gestiones de Cobranza (Últ.3)"
                        I = 22
                        Do
                            lbl(I) = !fecha & " " & !Resultado
                            
                            strTool = !Contacto & vbCrLf _
                            & !Usuario & vbCrLf & !Resultado
                            lbl(I).Tag = strTool
                            'AddCustomToolTip lbl(i), strTool, Me
                            I = I + 1
                            .MoveNext
                        Loop Until .EOF Or I = 25
                    Else
                        lbl(21) = ""
                        lbl(22) = ""
                        lbl(23) = ""
                        lbl(24) = ""
                    End If
                End With
                If !Demanda Then    'cliente demandado
                
                    MsgBox "El Propietario '" & !Codigo & " - " & !Nombre & "'," & vbCrLf & "se" _
                    & " encuentra demandado, imposible suministrar esta información." & vbCrLf _
                    & "Pongase en contacto con un supervisor.", vbInformation, "ATENCION !!!!"
                    Call rtnBitacora("Consultando Edo. Cuenta Inm: " & Dat(0) & " Prop.: " & Dat(2) _
                    & " Estatus: Demandado")
                    If gcNivel > nuSUPERVISOR Then Exit Sub
                    
                End If
                
                FlexFacturas(0).Visible = True
                If !Recibos >= 4 Then FlexFacturas(0).Visible = False
                lbl(20) = IIf(IsNull(!Notas), "", !Notas)
                txt(10) = Format(!Deuda, "#,##0.00")
                txt(9) = Format(0, "#,##0.00")
                txt(11) = txt(10)
                MskTelefono(0) = IIf(IsNull(!Fax), "", !Fax)
                MskTelefono(1) = IIf(IsNull(!TelfHab), "", !TelfHab)
                MskTelefono(2) = IIf(IsNull(!Celular), "", !Celular)
                Text1(7) = IIf(IsNull(!email), "", !email)
                Text1(8) = IIf(IsNull(!FecUltPag), "", !FecUltPag)
                Text1(9) = IIf(IsNull(!UltPago), "", !UltPago)
                Text1(10) = IIf(IsNull(!Alicuota), "", !Alicuota)
                Text1(11) = IIf(IsNull(!Recibos), "", !Recibos)
                Text1(9) = "Bs. " + Format(Text1(9), "#,##0.00")
                
                '
                If SSTab1.tab = 0 Then  'BUSCA LA DEUDA
                    If !Convenio Then frmConsultaCon.Show vbModal, FrmAdmin
                    
                    Call rtnBitacora("Consul.Edo.Cuenta Apto.: " & Dat(2) & "; Inmueble: " & Dat(0))
                    FlexFacturas(0).Rows = 2
                    Call rtnLimpiar_Grid(FlexFacturas(0))
                    
                    Call RtnFlex(Dat(2).Text, FlexFacturas(0), IntMMora, IntHonoMorosidad, _
                        6, txt(9), cnnPropietario, Dat(0))
                    If IsNull(txt(10)) Or txt(10) = "" Then txt(10) = 0
                    txt(11) = Format(CCur(txt(9)) + CCur((txt(10))), "#,##0.00")
                    
                Else                    'BUSCA LOS PAGOS
                    Call rtnBitacora("Consultando Pagos Apto: " & Dat(2) & "; Inmueble: " & Dat(0))
                    For I = 1 To 2
                            FlexFacturas(I).Rows = 2
                            Call rtnLimpiar_Grid(FlexFacturas(I))
                    Next
                    Call RtnFlexPagos
'                    If FlexFacturas(1).Rows > 15 Then
'                        FlexFacturas(1).ColWidth(2) = 2800
'                    Else
'                        FlexFacturas(1).ColWidth(2) = 3000
'                    End If
                End If
        
            Else    'SI NO CONSIGUE SIMILITUD
                Call rtnLimpiar_Grid(FlexFacturas(0))
                MsgBox "No Tengo Registrado ese Propietario " & "'" & StrControl & "'", _
                vbExclamation, App.ProductName
                Dat(2).SetFocus
            End If
            '
        End If
    '
    End With
    '----------------------------------
    On Error Resume Next
    ObjRstP(0).Close
    Set ObjRstP(0) = Nothing
    '----------------------------------
    '
    End Sub
    
    'Rev.23/08/2002--Rutina que busca el detalle de un pago específico----------------------------
    Private Sub RtnPagos()
    '---------------------------------------------------------------------------------------------
    'variables locales
    Dim rPeriodos As New ADODB.Recordset
    Dim strFPago As String
    Dim strVar As String
    Dim strSQL As String
    Dim strRecibo As String
    Dim I As Long
    '
    With FlexFacturas(1)
    '
        If strPago = FlexFacturas(1).TextMatrix(FlexFacturas(1).RowSel, 1) Then Exit Sub
        strRecibo = FlexFacturas(1).TextMatrix(FlexFacturas(1).RowSel, 1)
    '
    '       Periodos Cancelados con el recibo buscado
            strSQL = "SELECT MC.*,P.CodGasto as CGP,P.Descripcion as DP,P.Periodo as PP,P.Monto" _
            & " as MP, D.CodGasto as CGD, D.Titulo as DD, D.Monto as MD FROM (MovimientoCaja AS" _
            & " MC LEFT JOIN Periodos AS P ON MC.IDRecibo = P.IDRecibo) LEFT JOIN Deducciones " _
            & "AS D ON P.IDPeriodos = D.IDPeriodos WHERE MC.IDRecibo='" & strRecibo & _
            "' ORDER BY IIf(IsNull(P.Periodo) or P.Periodo='',Date(),CDate('01/' & P.Periodo))"
            '
            rPeriodos.Open strSQL, cnnConexion, adOpenKeyset, adLockOptimistic, adCmdText
            '
            strPago = FlexFacturas(1).TextMatrix(FlexFacturas(1).RowSel, 1)
            '
            If Not rPeriodos.EOF Or Not rPeriodos.BOF Then    '
                'forma de pago
                txt(0) = "": strFPago = ""
                'Si tiene efectivo
                strFPago = IIf(rPeriodos!EfectivoMovimientoCaja > 0, "Efectivo:......." _
                & rPeriodos("EfectivoMovimientoCaja"), "")
                'Documento 1/2/3
                For I = 0 To 2
                    strVar = IIf(I = 0, "", CStr(I))
                    strFPago = strFPago + IIf(IsNull(rPeriodos("Fpago" & strVar)), "", _
                    rPeriodos("Fpago" & strVar) & "  " _
                    & rPeriodos("NumdocumentoMovimientoCaja" & strVar) & "  " & _
                    rPeriodos("BancodocumentoMovimientoCaja" & strVar) & "  " & _
                    rPeriodos("fechaChequeMovimientoCaja" & strVar) & "  " & _
                    Format(rPeriodos("MontoCheque" & strVar), "#,##0.00") & vbCrLf)
                Next
                txt(0) = IIf(strFPago = "", "Pago en Efectivo....", strFPago)
                '
                With FlexFacturas(2)
                    .Rows = rPeriodos.RecordCount + 1
                    .ColWidth(1) = IIf(rPeriodos.RecordCount > 19, 2000, 2300)
                End With
        '
                I = 1: rPeriodos.MoveFirst
                Do 'Hacer hasta que sea fin de archivo
            
                    With FlexFacturas(2)    'Llena esta cuadrícula con la informacion de este pago
            '
                       If Not IsNull(rPeriodos("DP")) Then
                            .TextMatrix(I, 0) = IIf(IsNull(rPeriodos("PP")), "", " " & _
                            rPeriodos("PP"))
                            .TextMatrix(I, 1) = rPeriodos("DP")
                            .TextMatrix(I, 2) = Format(rPeriodos("MP"), "#,##0.00")
                        End If
                       'verifica si tiene deducciones el período
                        If Not IsNull(rPeriodos("cgd")) Then
                            .AddItem (""): I = I + 1
                            .TextMatrix(I, 0) = rPeriodos("CGD")
                            .TextMatrix(I, 1) = rPeriodos("DD")
                            .TextMatrix(I, 2) = Format(rPeriodos("MD"), "-#,##0.00")
                        End If
                        rPeriodos.MoveNext
                        I = I + 1
            '
                    End With
            '
                Loop Until rPeriodos.EOF
                
            Else
                FlexFacturas(2).Rows = 2
                Call rtnLimpiar_Grid(FlexFacturas(2))
            End If
            '
        rPeriodos.Close
        Set rPeriodos = Nothing
        'Set rPago = Nothing
        '
    End With
    '
    End Sub
    
    'Rev.23/08/2002---------Rutina que maneja los eventos al selecionar una de las fichas---------
    Private Sub RtnclickFicha(booVisible As Boolean, intFlex As Integer)
    '---------------------------------------------------------------------------------------------
    Call RtnVisible(booVisible)
    If Dat(0) <> "" And Dat(2) <> "" Then
        If booVisible = True Then
            Call RtnFlexPagos
        Else
            If Text1(11) > IntMMora Then FlexFacturas(0).Visible = False
            Call RtnFlex(Dat(2).Text, FlexFacturas(0), IntMMora, IntHonoMorosidad, _
                6, txt(9), cnnPropietario)
            txt(11) = Format(CDbl(txt(9)) + CDbl((txt(10))), "#,##0.00")
        End If
    
    End If
    Dat(0).SetFocus
    '
    End Sub

    Private Sub Print_Pagos()
    'variables locales
    Dim strSQL As String, strLocal As String
    Dim errLocal As Long
    Dim sReport As ctlReport
    '
    strSQL = "SELECT MC.IDTaquilla, MC.IDRecibo, MC.InmuebleMovimientoCaja, I.Nombre, I.Caja, C" _
    & ".DescripCaja, MC.AptoMovimientoCaja, MC.CuentaMovimientoCaja, MC.DescripcionMovimientoCa" _
    & "ja, MC.MontoMovimientoCaja, P.CodGasto, P.Descripcion, P.Periodo, P.Monto, D.CodGasto, D" _
    & ".Titulo, D.Monto, MC.FechaMovimientoCaja, MC.NumDocumentoMovimientoCaja, MC.NumDocumento" _
    & "MovimientoCaja1, MC.NumDocumentoMovimientoCaja2, MC.BancoDocumentoMovimientoCaja, MC.Ban" _
    & "coDocumentoMovimientoCaja1, MC.BancoDocumentoMovimientoCaja2, MC.FechaChequeMovimientoCa" _
    & "ja, MC.FechaChequeMovimientoCaja1, MC.FechaChequeMovimientoCaja2, MC.FormaPagoMovimiento" _
    & "Caja, MC.MontoCheque, MC.MontoCheque1, MC.MontoCheque2, MC.EfectivoMovimientoCaja, MC.FP" _
    & "ago, MC.FPago1, MC.FPago2, 0, I.FondoAct, P.Facturado FROM (((Caja AS C INNER JO" _
    & "IN inmueble AS I ON C.CodigoCaja = I.Caja) INNER JOIN MovimientoCaja AS MC ON I.CodInm =" _
    & "MC.InmuebleMovimientoCaja) LEFT JOIN Periodos AS P ON MC.IDRecibo = P.IDRecibo) LEFT JOI" _
    & "N Deducciones AS D ON P.IDPeriodos = D.IDPeriodos Where MC.InmuebleMovimientoCaja = '" & _
    Dat(0) & "' And MC.AptoMovimientoCaja = '" & Dat(2) & "' ORDER BY MC.FechaMovimientoCaja, " _
    & "Mc.Hora;"
    strLocal = gcPath & "\SAC.MDB"
    Call rtnGenerator(strLocal, strSQL, "QdfPagos")
    '
    Set sReport = New ctlReport
    With sReport
        .Reporte = gcReport + "edoctapro_pagos.rpt"
        .OrigenDatos(0) = gcPath + "\sac.mdb"
        .Formulas(0) = "Inmueble='" & Dat(0) & " - " & Dat(1) & "'"
        .Formulas(1) = "Propietario='" & Dat(2) & " - " & Dat(3) & "'"
        .TituloVentana = "Estado de Pagos " & Dat(0) & " / " & Dat(2)
        .Salida = crPantalla
        .Imprimir
        Call rtnBitacora("Impresión Pagos " & Dat(0) & "/" & Dat(2))
    End With
    Set sReport = Nothing
    '
    End Sub


    
    
