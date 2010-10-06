VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form FrmRemesa 
   Caption         =   "Registro de Remesa"
   ClientHeight    =   15
   ClientLeft      =   1650
   ClientTop       =   1140
   ClientWidth     =   2520
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   15
   ScaleWidth      =   2520
   WindowState     =   2  'Maximized
   Begin ComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   870
      Left            =   0
      TabIndex        =   19
      Top             =   0
      Width           =   2520
      _ExtentX        =   4445
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
            Object.Tag             =   "1"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Enabled         =   0   'False
            Key             =   "Previous"
            Object.ToolTipText     =   "Registro Anterior"
            Object.Tag             =   "2"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Enabled         =   0   'False
            Key             =   "Next"
            Object.ToolTipText     =   "Siguiente Registro"
            Object.Tag             =   "3"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Enabled         =   0   'False
            Key             =   "End"
            Object.ToolTipText     =   "Último Registro"
            Object.Tag             =   "4"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button5 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "New"
            Object.ToolTipText     =   "Nuevo Registro"
            Object.Tag             =   "5"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button6 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Save"
            Object.ToolTipText     =   "Guardar Registro"
            Object.Tag             =   "6"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button7 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Find"
            Object.ToolTipText     =   "Buscar Registro"
            Object.Tag             =   "7"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button8 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Undo"
            Object.ToolTipText     =   "Cancelar Registro"
            Object.Tag             =   "8"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button9 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Delete"
            Object.ToolTipText     =   "Eliminar Registro"
            Object.Tag             =   "9"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button10 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Edit1"
            Object.ToolTipText     =   "Editar Registro"
            Object.Tag             =   "10"
            ImageIndex      =   10
         EndProperty
         BeginProperty Button11 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Print"
            Object.ToolTipText     =   "Imprimir"
            Object.Tag             =   "11"
            ImageIndex      =   11
         EndProperty
         BeginProperty Button12 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Close"
            Object.ToolTipText     =   "Salir"
            Object.Tag             =   "12"
            ImageIndex      =   12
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin VB.Timer Timer1 
      Left            =   5340
      Top             =   4020
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   6765
      Left            =   105
      TabIndex        =   17
      Top             =   585
      Width           =   11010
      _ExtentX        =   19420
      _ExtentY        =   11933
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
      TabCaption(0)   =   "      Datos Generales     "
      TabPicture(0)   =   "FrmRemesa.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "fraRemesa(0)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "          Lista          "
      TabPicture(1)   =   "FrmRemesa.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "fraRemesa(2)"
      Tab(1).Control(1)=   "flexRemesa(1)"
      Tab(1).ControlCount=   2
      Begin VB.Frame fraRemesa 
         Caption         =   "Ordenar por:"
         Height          =   1890
         Index           =   2
         Left            =   -65745
         TabIndex        =   25
         Top             =   840
         Width           =   1485
         Begin VB.OptionButton Opt 
            Caption         =   "Fecha"
            Height          =   225
            Index           =   3
            Left            =   165
            TabIndex        =   29
            Tag             =   "3"
            Top             =   1440
            Width           =   1170
         End
         Begin VB.OptionButton Opt 
            Caption         =   "Cargado"
            Height          =   225
            Index           =   2
            Left            =   165
            TabIndex        =   28
            Tag             =   "2"
            Top             =   1080
            Width           =   1170
         End
         Begin VB.OptionButton Opt 
            Caption         =   "Servicio"
            Height          =   225
            Index           =   1
            Left            =   165
            TabIndex        =   27
            Tag             =   "1"
            Top             =   720
            Width           =   1170
         End
         Begin VB.OptionButton Opt 
            Caption         =   "Nº Remesa"
            Height          =   225
            Index           =   0
            Left            =   165
            TabIndex        =   26
            Tag             =   "0"
            Top             =   360
            Value           =   -1  'True
            Width           =   1170
         End
      End
      Begin VB.Frame fraRemesa 
         Height          =   6090
         Index           =   0
         Left            =   240
         TabIndex        =   18
         Top             =   480
         Width           =   10500
         Begin VB.Frame fraRemesa 
            BorderStyle     =   0  'None
            ClipControls    =   0   'False
            Enabled         =   0   'False
            Height          =   1950
            Index           =   1
            Left            =   270
            TabIndex        =   20
            Top             =   195
            Width           =   10140
            Begin VB.CommandButton cmdRemesa 
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   440
               Index           =   4
               Left            =   9615
               Picture         =   "FrmRemesa.frx":0038
               Style           =   1  'Graphical
               TabIndex        =   34
               ToolTipText     =   "Pegar Copia"
               Top             =   780
               Width           =   440
            End
            Begin VB.TextBox txtRemesa 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00FFFFFF&
               DataField       =   "Codigo"
               DataSource      =   "AdoRemesa"
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
               Height          =   315
               Index           =   6
               Left            =   3690
               Locked          =   -1  'True
               TabIndex        =   33
               Text            =   "0,00"
               Top             =   1290
               Width           =   1335
            End
            Begin VB.TextBox txtRemesa 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00FFFFFF&
               DataField       =   "Codigo"
               DataSource      =   "AdoRemesa"
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
               Height          =   315
               Index           =   5
               Left            =   1110
               TabIndex        =   31
               Text            =   "0,00"
               ToolTipText     =   "Monto Total de la remesa"
               Top             =   1290
               Width           =   1335
            End
            Begin MSMask.MaskEdBox mskRemesa 
               Height          =   315
               Left            =   1110
               TabIndex        =   5
               Top             =   540
               Width           =   1335
               _ExtentX        =   2355
               _ExtentY        =   556
               _Version        =   393216
               MaxLength       =   7
               Format          =   "mm-yyyy"
               Mask            =   "##-####"
               PromptChar      =   "_"
            End
            Begin VB.CommandButton cmdRemesa 
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   440
               Index           =   3
               Left            =   9180
               Picture         =   "FrmRemesa.frx":013A
               Style           =   1  'Graphical
               TabIndex        =   6
               ToolTipText     =   "Pegar Copia"
               Top             =   780
               Width           =   440
            End
            Begin VB.CommandButton cmdRemesa 
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   440
               Index           =   2
               Left            =   8745
               Picture         =   "FrmRemesa.frx":0284
               Style           =   1  'Graphical
               TabIndex        =   24
               ToolTipText     =   "Copiar"
               Top             =   780
               Width           =   440
            End
            Begin VB.CommandButton cmdRemesa 
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   440
               Index           =   1
               Left            =   8310
               Picture         =   "FrmRemesa.frx":03CE
               Style           =   1  'Graphical
               TabIndex        =   23
               ToolTipText     =   "Agregar linea"
               Top             =   780
               Width           =   440
            End
            Begin VB.CommandButton cmdRemesa 
               Caption         =   "&Validar"
               Height          =   440
               Index           =   0
               Left            =   8325
               TabIndex        =   22
               ToolTipText     =   "Validar Remesa"
               Top             =   1215
               Width           =   1725
            End
            Begin VB.TextBox txtRemesa 
               DataField       =   "Titulo"
               DataSource      =   "AdoRemesa"
               Height          =   315
               Index           =   1
               Left            =   3705
               TabIndex        =   3
               Text            =   " "
               Top             =   165
               Width           =   6330
            End
            Begin VB.TextBox txtRemesa 
               BackColor       =   &H00FFFFFF&
               DataField       =   "Codigo"
               DataSource      =   "AdoRemesa"
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
               Left            =   1110
               MaxLength       =   5
               TabIndex        =   1
               Top             =   165
               Width           =   1335
            End
            Begin VB.OptionButton optRemesa 
               Caption         =   "Agua"
               Height          =   210
               Index           =   1
               Left            =   5535
               TabIndex        =   8
               Tag             =   "Inmueble"
               Top             =   585
               Width           =   1395
            End
            Begin VB.OptionButton optRemesa 
               Caption         =   "Teléfono"
               Height          =   210
               Index           =   2
               Left            =   6930
               TabIndex        =   9
               Tag             =   "Inmueble"
               Top             =   585
               Width           =   1395
            End
            Begin VB.OptionButton optRemesa 
               Caption         =   "Luz Eléctrica"
               Height          =   210
               Index           =   0
               Left            =   3840
               TabIndex        =   7
               Tag             =   "IDCS"
               Top             =   585
               Value           =   -1  'True
               Width           =   1395
            End
            Begin VB.TextBox txtRemesa 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00FFFFFF&
               DataField       =   "Codigo"
               DataSource      =   "AdoRemesa"
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
               Height          =   315
               Index           =   2
               Left            =   1110
               Locked          =   -1  'True
               TabIndex        =   11
               Text            =   "0,00"
               Top             =   915
               Width           =   1335
            End
            Begin VB.TextBox txtRemesa 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00FFFFFF&
               DataField       =   "Codigo"
               DataSource      =   "AdoRemesa"
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
               Height          =   315
               Index           =   3
               Left            =   3690
               Locked          =   -1  'True
               TabIndex        =   13
               Text            =   "0,00"
               Top             =   915
               Width           =   1335
            End
            Begin VB.TextBox txtRemesa 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00FFFFFF&
               DataField       =   "Codigo"
               DataSource      =   "AdoRemesa"
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
               Height          =   315
               Index           =   4
               Left            =   6525
               Locked          =   -1  'True
               TabIndex        =   15
               Text            =   "0,00"
               Top             =   915
               Width           =   1335
            End
            Begin VB.Label lblRemesa 
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
               Index           =   8
               Left            =   2580
               TabIndex        =   32
               Top             =   1342
               Width           =   1020
            End
            Begin VB.Label lblRemesa 
               Caption         =   "Remesa Bs.:"
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
               Left            =   45
               TabIndex        =   30
               Top             =   1335
               Width           =   1035
            End
            Begin VB.Label lblRemesa 
               Caption         =   "Nº &Remesa:"
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
               Left            =   45
               TabIndex        =   0
               Top             =   210
               Width           =   1035
            End
            Begin VB.Label lblRemesa 
               Caption         =   "&Cargado:"
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
               Left            =   45
               TabIndex        =   4
               Top             =   585
               Width           =   1035
            End
            Begin VB.Label lblRemesa 
               AutoSize        =   -1  'True
               Caption         =   "&Descripción :"
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
               Left            =   2580
               TabIndex        =   2
               Top             =   210
               Width           =   1035
            End
            Begin VB.Label lblRemesa 
               AutoSize        =   -1  'True
               Caption         =   "Tipo Servicio:"
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
               Left            =   2580
               TabIndex        =   21
               Top             =   585
               Width           =   1095
            End
            Begin VB.Label lblRemesa 
               Caption         =   "Monto:"
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
               Left            =   45
               TabIndex        =   10
               Top             =   960
               Width           =   1035
            End
            Begin VB.Label lblRemesa 
               AutoSize        =   -1  'True
               Caption         =   "IDB:"
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
               Left            =   2580
               TabIndex        =   12
               Top             =   967
               Width           =   345
            End
            Begin VB.Label lblRemesa 
               AutoSize        =   -1  'True
               Caption         =   "Monto Neto:"
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
               Left            =   5370
               TabIndex        =   14
               Top             =   960
               Width           =   1050
            End
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid flexRemesa 
            Height          =   3615
            Index           =   0
            Left            =   240
            TabIndex        =   16
            Top             =   2250
            Width           =   10080
            _ExtentX        =   17780
            _ExtentY        =   6376
            _Version        =   393216
            Cols            =   9
            FixedCols       =   0
            BackColorFixed  =   -2147483646
            ForeColorFixed  =   -2147483639
            BackColorSel    =   65280
            ForeColorSel    =   -2147483646
            BackColorBkg    =   -2147483636
            FocusRect       =   2
            HighLight       =   0
            MergeCells      =   1
            AllowUserResizing=   1
            FormatString    =   "Codigo |<Cliente |ID.Cliente |^Periodo  |Monto |I.D.B. |Neto |Recl. |Desd. "
            BandDisplay     =   1
            BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            _NumberOfBands  =   1
            _Band(0).Cols   =   9
            _Band(0).GridLinesBand=   1
            _Band(0).TextStyleBand=   0
            _Band(0).TextStyleHeader=   0
         End
         Begin VB.Image imgRemesa 
            Enabled         =   0   'False
            Height          =   480
            Index           =   1
            Left            =   0
            Top             =   0
            Visible         =   0   'False
            Width           =   480
         End
         Begin VB.Image imgRemesa 
            Enabled         =   0   'False
            Height          =   480
            Index           =   0
            Left            =   100
            Top             =   100
            Visible         =   0   'False
            Width           =   480
         End
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid flexRemesa 
         Bindings        =   "FrmRemesa.frx":0518
         CausesValidation=   0   'False
         Height          =   5115
         Index           =   1
         Left            =   -74460
         TabIndex        =   35
         Top             =   825
         Width           =   5325
         _ExtentX        =   9393
         _ExtentY        =   9022
         _Version        =   393216
         Cols            =   5
         FixedCols       =   0
         BackColorFixed  =   -2147483646
         ForeColorFixed  =   -2147483639
         BackColorSel    =   65280
         ForeColorSel    =   -2147483646
         BackColorBkg    =   -2147483636
         GridColor       =   -2147483633
         AllowBigSelection=   0   'False
         FocusRect       =   2
         HighLight       =   0
         GridLinesFixed  =   0
         GridLinesUnpopulated=   4
         MergeCells      =   1
         AllowUserResizing=   1
         FormatString    =   "<Remesa |<Titulo |^Cargado |^Fecha |^Hora"
         BandDisplay     =   1
         BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _NumberOfBands  =   1
         _Band(0).Cols   =   5
         _Band(0).TextStyleBand=   0
         _Band(0).TextStyleHeader=   0
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
            Picture         =   "FrmRemesa.frx":052D
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmRemesa.frx":06AF
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmRemesa.frx":0831
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmRemesa.frx":09B3
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmRemesa.frx":0B35
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmRemesa.frx":0CB7
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmRemesa.frx":0E39
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmRemesa.frx":0FBB
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmRemesa.frx":113D
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmRemesa.frx":12BF
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmRemesa.frx":1441
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmRemesa.frx":15C3
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "FrmRemesa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
    '---------------------------------------------------------------------------------------------
    'SINAI TECH, C.A Módulo cuentas por pagar. Ventana Registrar Remesa
    'Variables Públicas a nivel de módulo 20/01/2003
    '---------------------------------------------------------------------------------------------
    Dim rstRemesa(3) As New ADODB.Recordset
    Private Enum scRem
        scRemesa
        scInm
        scApto
        scDetalle
    End Enum
    '
    Dim i%, j%
    Dim byServ As Byte
    Dim mlEdit As Boolean
    Const scPAGADO$ = "PAGADO"
    Const scASIGNADO$ = "ASIGNADO"
    '
    Private Sub cmdRemesa_Click(Index As Integer)
    'variables locales
    Dim numFichero As Integer
    Dim strArchivo As String
    '
    Me.MousePointer = vbArrowHourglass
    Select Case Index
        '
        Case 0  'Validar Remesa
            Call Validar_Remesa
            
        Case 1  'Agregar Linea
            With flexRemesa(0)
                .AddItem ("")
                Call iniciar_check(.Rows - 1)
                If .Rows > 12 Then .TopRow = .Rows - 12
            End With
        
        Case 2  'guarda un temp de las facturas introducidas hasta ahora
        '
            For i = optRemesa.LBound To optRemesa.UBound
                If optRemesa(i).Value = True Then Exit For
            Next
            With flexRemesa(0)
                numFichero = FreeFile
                strArchivo = gcPath & "\TEMPREM00" & i & ".log"
                Open strArchivo For Output As numFichero
                    j = 1
                    Do
                        If .TextMatrix(j, 3) <> "" Or .TextMatrix(j, 4) <> "" Then
                        Write #numFichero, j, .TextMatrix(j, 3), .TextMatrix(j, 4)
                        End If
                        j = j + 1
                    Loop Until j > .Rows - 1
                Close numFichero
            End With
            Call rtnBitacora("Guardar Temp. Remesa Nº " & txtRemesa(0))
            
        Case 3  'Recuperar
            Call Recuperar_Rem
            Call rtnBitacora("Recuperar Remesa..")
            '
    End Select
    Me.MousePointer = vbDefault
    '
    End Sub
    
    Private Sub flexremesa_Click(Index As Integer)
    'VARIABLES LOCALES
    Select Case Index
    '
        Case 0
            '
            With flexRemesa(0)
            
                If .RowSel > 0 And .ColSel = 7 Or .RowSel > 0 And .ColSel = 8 Then
                    Call Checked(.ColSel, .RowSel)
                End If
                
            End With
            '
        End Select
    '
    End Sub


    Private Sub flexRemesa_DblClick(Index As Integer)
    '
    Dim X%, i%
    Dim curTotales(4 To 6) As Currency
    '
    Select Case Index
        Case 1
        '
            flexRemesa(1).MousePointer = vbArrowHourglass
            If flexRemesa(1).RowSel > 0 And _
            Not flexRemesa(1).TextMatrix(flexRemesa(1).RowSel, 0) = "" Then
            
                txtRemesa(0) = flexRemesa(1).TextMatrix(flexRemesa(1).RowSel, 0)
                txtRemesa(1) = flexRemesa(1).TextMatrix(flexRemesa(1).RowSel, 1)
                mskRemesa.PromptInclude = False
                mskRemesa.Text = flexRemesa(1).TextMatrix(flexRemesa(1).RowSel, 2)
                mskRemesa.PromptInclude = True
                '
                rstRemesa(scDetalle).Open "SELECT CodInm,Apto,Nis,Periodo,Bruto,IDB,Monto,Recla" _
                & "mo,Desdomi FROM DetRemesa WHERE IDRemesa=" & _
                flexRemesa(1).TextMatrix(flexRemesa(1).RowSel, 0) & ";", cnnConexion, _
                adOpenStatic, adLockReadOnly, adCmdText
                '
                With rstRemesa(scDetalle)
                    If .RecordCount > 0 Then
                        
                        .MoveFirst
                        X = 1
                        flexRemesa(0).Rows = .RecordCount + 1
                        Do
                            For i = 0 To flexRemesa(0).Cols - 1
                                
                                If i >= 4 And i <= 6 Then curTotales(i) = _
                                curTotales(i) + .Fields(i)
                                If i = 7 Or i = 8 Then
                                    flexRemesa(0).Col = i
                                    flexRemesa(0).Row = X
                                    Set flexRemesa(0).CellPicture = _
                                    IIf(.Fields(i), imgRemesa(1), imgRemesa(0))
                                    flexRemesa(0).CellPictureAlignment = flexAlignCenterCenter
                                Else
                                    flexRemesa(0).TextMatrix(X, i) = _
                                    IIf(i >= 4 And i <= 6, Format(.Fields(i), "#,##0.00"), _
                                    IIf(IsNull(.Fields(i)), "", .Fields(i)))
                                    flexRemesa(0).TextMatrix(X, 1) = _
                                        Buscar_Cliente(IIf(IsNull(.Fields(1)), "", .Fields(1)), .Fields(0))
                                End If
                            Next
                            .MoveNext
                            X = X + 1
                        Loop Until .EOF
                        Call Ampliar_Flex
                        For i = 2 To 4
                            txtRemesa(i) = Format(curTotales(i + 2), "#,##0.00")
                        Next
                    End If
                    .Close
                End With
                '
            End If
            flexRemesa(1).MousePointer = vbDefault
    End Select
    '
    End Sub

    Private Sub flexRemesa_EnterCell(Index%): Call Marcar_Linea(flexRemesa(Index), vbGreen)
    End Sub

    
    Private Sub flexremesa_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    
    If Index = 0 Then
        If flexRemesa(0).Row > 0 And (flexRemesa(0).Col >= 3 And flexRemesa(0).Col <= 4) Then
            If KeyCode = 46 Then flexRemesa(0).Text = ""
            If Shift = 2 Then
            '
                If KeyCode = 67 Then
                    Clipboard.Clear
                    Clipboard.SetText flexRemesa(0).Text
                ElseIf KeyCode = 86 Then
                    flexRemesa(0).Text = Clipboard.GetText
                End If
                '
            End If
            '
        End If
        '
    End If
    '
    End Sub

    Private Sub flexremesa_KeyPress(Index As Integer, KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    Select Case Index
        Case 0
            '
            With flexRemesa(0)
                '
                Select Case .ColSel
                
                    Case 3  'Periodo
                        Call Validacion(KeyAscii, "0123456789/")
                        If KeyAscii = 8 And Len(.Text) > 0 Then .Text = Left(.Text, Len(.Text) - 1)
                        If KeyAscii > 26 Then
                            If KeyAscii = Asc("/") Then
                                If InStr(.Text, "/") = 0 Then .Text = .Text & Chr(KeyAscii)
                            Else
                            .Text = .Text & Chr(KeyAscii)
                            End If
                        End If
                        If KeyAscii = Asc("/") Then
                            If Len(.Text) = 2 Then .Text = Format(Left(.Text, 1), "00") & "/"
                            If Left(.Text, 2) > 12 Then .Text = "12/"
                        End If
                        If KeyAscii = 13 Then .Col = 4
                        
                    Case 2  'ID.CLiente
                        Call Validacion(KeyAscii, "1234567890.")
                        If KeyAscii = 8 And Len(.Text) > 0 Then .Text = Left(.Text, Len(.Text) - 1)
                        If KeyAscii > 26 Then .Text = .Text & Chr(KeyAscii)
                        If KeyAscii = 13 Then .Col = 3
                        
                    Case 4  'Monto
                        If KeyAscii = Asc(".") Then KeyAscii = Asc(",")
                        Call Validacion(KeyAscii, "-0123456789,")
                        If KeyAscii = 8 Then .Text = Left(.Text, Len(.Text) - 1)
                        If KeyAscii > 26 Then .Text = .Text & Chr(KeyAscii)
                        If KeyAscii = 13 Then
                            If .Text <> "" Then
                                .Text = Format(.Text, "#,##0.00")
                                If Not IsNumeric(.Text) Then
                                    MsgBox "Monto no válido [" & .Text & "]", vbExclamation, _
                                    App.ProductName
                                    Exit Sub
                                End If
                            End If
                            .Col = 3
                        End If
                        '
                End Select
            End With
        
        Case 1
        
    End Select
    '
    End Sub


    Private Sub flexremesa_LeaveCell(Index As Integer)
    '
    'Variables locales
    Dim Mes%, Año%, Pos%
    '
    Call Marcar_Linea(flexRemesa(Index), vbWhite)
    Select Case Index
    
        Case 0
            With flexRemesa(0)
                If flexRemesa(0).ColSel = 3 And .Text <> "" Then
                    If IsDate("01-" & .Text) Then
                        Pos = InStr(.Text, "/")
                        Mes = Left(.Text, Pos - 1)
                        Año = Right(.Text, Len(.Text) - Pos)
                        If Mes = 0 Or Mes > 12 Or 1975 > Año Or Año > Year(Date) + 1 Then
                            MsgBox "Verifica la fecha introducida...", vbExclamation, App.ProductName
                            .SetFocus
                        End If
                    Else
                        MsgBox "Intrudujo una fecha inválida", vbInformation, App.ProductName
                        .SetFocus
                    End If
                ElseIf flexRemesa(0).ColSel = 4 And Not .Text = "" Then
                    If Not IsNumeric(.Text) Then
                        MsgBox "Monto no válido [" & .Text & "]", vbExclamation, App.ProductName
                    End If
                
                End If
                .Col = 4
                .Row = .RowSel
                .Text = Format(.Text, "#,##0.00")
                
            End With
        Case 1
        
    End Select
    '
    End Sub

    Private Sub flexRemesa_LostFocus(Index As Integer)
    '
    With flexRemesa(Index)
        '
        .Col = 0
        .Row = 1
        .ColSel = .Cols - 1
        .RowSel = .Rows - 1
        .FillStyle = flexFillRepeat
        .CellBackColor = vbWhite
        .CellFontBold = False
        .FillStyle = flexFillSingle
        '
    End With
    '
    End Sub


    '
    Private Sub Form_Load()
    '
    imgRemesa(0).Picture = LoadResPicture("Unchecked", vbResBitmap)
    imgRemesa(1).Picture = LoadResPicture("Checked", vbResBitmap)
    
    For i = 0 To 1
        Set flexRemesa(i).FontFixed = LetraTitulo(LoadResString(527), 9, , True)
        Set flexRemesa(i).Font = LetraTitulo(LoadResString(528), 9)
    Next i
    flexRemesa(0).RowHeight(0) = 315
    flexRemesa(1).RowHeight(0) = 315
    '
    rstRemesa(scRemesa).Open "SELECT IDRemesa,Titulo,Format(Cargado,'mm/yyyy') AS Periodo,Forma" _
    & "t(Fecha,'Short Date'),Format(Hora,'hh:mm:ss') FROM Remesa ORDER BY Fecha DESC;", _
    cnnConexion, adOpenKeyset, adLockOptimistic, adCmdText
    '
    Call RtnEstado(6, Toolbar1)
    
    Set flexRemesa(1).DataSource = rstRemesa(scRemesa)
    
    If rstRemesa(scRemesa).RecordCount > 0 Then
        For i = 1 To 4: Toolbar1.Buttons(i).Enabled = True
        Next
        Toolbar1.Buttons(11).Enabled = False
    Else
        flexRemesa(1).Rows = 2
    End If
    '
    flexRemesa(1).FormatString = "<Remesa |<Titulo |^Cargado |^Fecha |^Hora"
    '
    For i = optRemesa.LBound To optRemesa.UBound
        If optRemesa(i).Value = True Then txtRemesa(1) = "Servicio de " & optRemesa(i).Caption
        rstRemesa(scInm).Open "SELECT * FROM Servicios WHERE TipoServ=" & i, _
        cnnConexion, adOpenStatic, adLockReadOnly, adCmdText
        rstRemesa(scApto).Open "SELECT * FROM ServiciosApto WHERE TipoServ=" _
        & i, cnnConexion, adOpenStatic, adLockReadOnly, adCmdText
        Exit For
    Next
    '
    If rstRemesa(scRemesa).EOF Or rstRemesa(scRemesa).BOF Then
        For i = 1 To 12
            If i <> 5 And i <> 12 Then Toolbar1.Buttons(i).Enabled = False
        Next
    End If
    Call Presentar_Grid(0)
    Call Presentar_Grid(1)
    If gcNivel > nuSUPERVISOR Then
        For i = 6 To 10: If Not i = 8 Then Toolbar1.Buttons(i).Visible = False
        Next
    End If
    '
    End Sub

    Private Sub Form_Resize()
    On Error Resume Next
    If WindowState <> vbMinimized Then
        SSTab1.Left = ScaleLeft + 100
        SSTab1.Top = ScaleTop + Toolbar1.Height + 100
        SSTab1.Width = ScaleWidth - SSTab1.Left
        SSTab1.Height = ScaleHeight - SSTab1.Top
        SSTab1.tab = 0
        fraRemesa(0).Width = SSTab1.Width - (fraRemesa(0).Left * 2)
        fraRemesa(0).Height = SSTab1.Height - (fraRemesa(0).Top * 2)
        flexRemesa(0).Width = fraRemesa(0).Width - (flexRemesa(0).Left * 2)
        flexRemesa(0).Height = fraRemesa(0).Height - flexRemesa(0).Top - 200
        SSTab1.tab = 1
        flexRemesa(1).Height = SSTab1.Height - flexRemesa(1).Top - 200
    End If
    End Sub

    Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    For i = 0 To 3: Set rstRemesa(i) = Nothing
    Next
    Set FrmRemesa = Nothing
    End Sub

    Private Sub mskRemesa_GotFocus()
    mskRemesa.SelStart = 0
    mskRemesa.SelLength = 0
    End Sub

    Private Sub Opt_Click(Index As Integer)
    rstRemesa(scRemesa).Sort = rstRemesa(scRemesa).Fields(Index).Name
    flexRemesa(1).MergeCol(Index) = True
    Timer1.Interval = 1
    End Sub

    Private Sub optRemesa_Click(Index As Integer)
    txtRemesa(1) = "Servicio de " & optRemesa(Index).Caption
    Call RtnAdd(5)
    End Sub


    Private Sub SSTab1_Click(PreviousTab As Integer)
    '
    If SSTab1.tab = 0 Then
        Toolbar1.Buttons(11).Enabled = True
        Toolbar1.Buttons("Delete").Enabled = False
    Else
        Toolbar1.Buttons(11).Enabled = False
        Toolbar1.Buttons("Delete").Enabled = True
    End If
    '
    End Sub

    
    Private Sub Timer1_Timer()
    Timer1.Interval = 0
    With flexRemesa(1)
        .TextArray(2) = "Cargado"
        .TextArray(3) = "Fecha"
        .TextArray(4) = "Hora"
    End With
    Call Presentar_Grid(1)
    End Sub

    '
    Private Sub Toolbar1_ButtonClick(ByVal Button As ComctlLib.Button)
    'variables locales
    Dim strSQL$, Cargado$, strMensaje$, Remesa&
    Dim rpReporte As ctlReport
    '
    Select Case UCase(Button.Key)
        
        Case "FIRST", "NEXT", "PREVIOUS", "END" 'recorrido
        '--------------------
            Call Moverse(Button.Index)
        
        Case "NEW", "UNDO" 'Agregar Remesa y cancelar
        '--------------------
            Call RtnEstado(Button.Index, Toolbar1)
            Call RtnAdd(Button.Index)
                        
        Case "SAVE" 'Guarda Remesa
        '--------------------
            
            If Not validado Then
                MousePointer = vbHourglass
                Call RtnConfigUtility(True, "SAC", "Iniciando el Proceso...", "Guardar Remesa")
                Call Guardar_Remesa
                Unload FrmUtility
                SSTab1.TabEnabled(1) = True
                MousePointer = vbDefault
            End If
            
        Case "PRINT"    'Imprimir Reportes
        '--------------------
            If txtRemesa(0) = "" Then Exit Sub
            'Genera una consulta del número de remesa
            
            Me.MousePointer = vbHourglass
           strSQL = "SELECT R.*, D.*, I.Nombre, I.TipoCta FROM Remesa as R RIGHT JOIN (DetRemes" _
            & "a as D LEFT JOIN Inmueble as I ON D.CodInm = I.CodInm) ON R.IDRemesa = D.IDRemes" _
            & "a WHERE (((D.IDRemesa)=" & txtRemesa(0) & "));"
            '
            Call rtnGenerator(gcPath & "\sac.mdb", strSQL, "rptRemesa")
            '
            Cargado = "01-" & Left(mskRemesa, 2) & "-" & Right(mskRemesa, 4)
            'Call clear_Crystal(FrmAdmin.rptReporte)
            '
            Set rpReporte = New ctlReport
            With rpReporte
                '
                .OrigenDatos(0) = gcPath & "\sac.mdb"
                .Reporte = gcReport + "remesa_report.rpt"
                .Formulas(0) = "Cargado='" & UCase(Format(Cargado, "mmm/yyyy")) & "'"
                .Formulas(1) = "Titulo='" & txtRemesa(1) & "'"
                .TituloVentana = txtRemesa(1)
                .Imprimir
                Call rtnBitacora("Impresión Remesa Nº " & txtRemesa(0))
                '
            End With
            Set rpReporte = Nothing
            Me.MousePointer = vbDefault
            '
        Case "FIND" 'Buscar
        '--------------------
            SSTab1.tab = 1
            
        Case "DELETE"   'Eliminar Remesa
        '--------------------
            If gcNivel < nuSUPERVISOR Then
                Remesa = flexRemesa(1).TextMatrix(flexRemesa(1).RowSel, 0)
                strMensaje = "Eliminará toda la información asociada con la remesa Nº ," & Remesa _
                & vbCrLf & "Cuentas por Pagar, Facturación, etc. ¿Desea Continuar?"
                'Usuario desea eliminar la remesa
                If Respuesta(strMensaje) Then Call Eliminar_Remesa(Remesa)
                
            Else    'no tiene el nivel necesario para efectuar esta operación
                MsgBox "Usted no tiene los permisos necesarios para llevar adelante esta transa" _
                & "cción" & bvcrlf & "Consulte con el administrador del sistema", vbExclamation, _
                App.ProductName
                'si dispone de RealpoPup lo utiliza como medio para solicitar la autorización
                'a un supervisor
                If Dir("C:\Archivos de programa\RealPopup\RealPopup.exe") <> "" Then
                    strMensaje = "Intento: Eliminar remesa Nº" & Remesa
                    m = Shell("C:\Archivos de programa\RealPopup\RealPopup -send archivo " & _
                    Chr(34) & strMensaje & Chr(34) & " -NOACTIVATE")
                End If
                Call rtnBitacora(strMensaje)
                '
            End If
             '
        Case "EDIT1"
        '--------------------
            MsgBox "Opción No Disponible....Por ahora.", vbInformation, App.ProductName
            '
        Case "CLOSE"
        '--------------------
            Unload Me
            '
    End Select
    '
    End Sub

    Private Sub iniciar_check(ByVal intFil As Integer)
    '
    With flexRemesa(0)
    '
        For i = 7 To 8
            .Col = i
            .Row = intFil
            Set .CellPicture = imgRemesa(0)
            .CellPictureAlignment = flexAlignCenterCenter
        Next
        '
    End With
    '
    End Sub
    
    Private Sub Checked(Col As Integer, Fil As Integer)
    With flexRemesa(0)
        .Col = Col
        .Row = Fil
        Set .CellPicture = IIf(.CellPicture = imgRemesa(0), imgRemesa(1), imgRemesa(0))
        .CellPictureAlignment = flexAlignCenterCenter
    End With
    End Sub
    
    '---------------------------------------------------------------------------------------------
    '   Rutina:     Validar_Resema
    '
    '---------------------------------------------------------------------------------------------
    Private Sub Validar_Remesa()
    'variables locales
    Dim strInm$, StrApto$, m$
    Dim curMonto(3 To 5) As Currency
    Dim booDesDomi$, booRecl$
    Dim rstFD As New ADODB.Recordset
    Dim strCriterio As String
    '
    If txtRemesa(5) = "" Then txtRemesa(5) = "0,00"
    If CCur(txtRemesa(5)) = 0 Then  'veirifca el monto de la remesasa
        MsgBox "Debe introducir el monto de la remesa antes de validar", vbInformation, _
        App.ProductName
        Exit Sub
    ElseIf Not IsDate(mskRemesa) Then   'valor de datosç período
        MsgBox "Período no valido", vbCritical, App.ProductName
        mskRemesa.SetFocus
        Exit Sub
    End If
    '--------------------------------------------
    Me.MousePointer = vbArrowHourglass
    For A = optRemesa.LBound To optRemesa.UBound
        If optRemesa(A).Value = True Then byServ = A: Exit For
    Next

    j = 1: m = Left(mskRemesa, 2) & "/" & Right(mskRemesa, 4)
    '
    Do  'Valida linea por linea
        
        'si el período y/o el monto estan vacios desdomicilia la factura
        If flexRemesa(0).TextMatrix(j, 4) = "" Or flexRemesa(0).TextMatrix(j, 3) = "" Then
            strInm = Desdomiciliar(j)
        Else
        
            With rstRemesa(scInm)  'Buscar en las cuentas de serv. de los inmueble
            
                If Not .EOF Or Not .BOF Then
                    .MoveFirst
                    .Find "IDCS='" & flexRemesa(0).TextMatrix(j, 2) & "'"
                    If Not .EOF Then
                        strInm = IIf(Facturado(!Inmueble, "01/" & m), Desdomiciliar(j), !Inmueble)
                        StrApto = ""
                        If strInm = "" Then GoTo 20
                    Else
                        GoTo 10
                    End If
                Else    'si no hay coincidencia
10                    With rstRemesa(scApto) 'Busca en los apartamentos domiciliados
                        If Not .EOF Or Not .BOF Then
                            .MoveFirst
                            .Find "IDCS='" & flexRemesa(0).TextMatrix(j, 2) & "'"
                            If Not .EOF Then
                                strInm = IIf(Facturado(!Inmueble, "01/" & m), Desdomiciliar(j), !Inmueble)
                                StrApto = !apto
                            Else    'No lo encuentra en el archivo
                                strInm = Desdomiciliar(j)
                                StrApto = strInm
                            End If
                        Else
                            strInm = Desdomiciliar(j)
                            StrApto = strInm
                        End If
                        '
                    End With
                    '
                End If
                '
            End With
            '
            'veirfica si la factura esta duplicada
            rstFD.Open "SELECT * FROM Cargado WHERE Detalle LIKE '%ID%" & _
            flexRemesa(0).TextMatrix(j, 2) & "%" & flexRemesa(0).TextMatrix(j, 3) & "%'", _
            cnnOLEDB & gcPath & "\" & strInm & "\inm.mdb", adOpenKeyset, adLockOptimistic, _
            adCmdText
            '
            If Not rstFD.EOF And Not rstFD.BOF Then
            
                MsgBox "Facura ya procesada:" & vbCrLf & rstFD("Detalle") & vbCrLf & "Fecha: " _
                & rstFD("Fecha") & vbCrLf & "Usuario: " & rstFD("Usuario") & vbCrLf & "Se desdo" _
                & "miciliará", vbInformation, App.ProductName
                strInm = Desdomiciliar(j)
                'End If
            
            Else
                rstFD.Close
                '
                strCriterio = Format(DateAdd("M", -1, "01/" & flexRemesa(0).TextMatrix(j, 3)), _
                "MM/YYYY")
                '
                'verifica ahora que la factura anterior a esta este recibida
                rstFD.Open "SELECT * FROM Cargado WHERE Detalle LIKE '%ID%" & _
                flexRemesa(0).TextMatrix(j, 2) & "%" & strCriterio & "%'", cnnOLEDB & gcPath _
                & "\" & strInm & "\inm.mdb", adOpenKeyset, adLockOptimistic, adCmdText
                '
                If rstFD.EOF And rstFD.BOF Then
                    rstFD.Close
                    'verificamos que no sea la primera factura domiciliada
                    rstFD.Open "SELECT * FROM Cargado WHERE Detalle LIKE '%ID%" & _
                    flexRemesa(0).TextMatrix(j, 2) & "%'", cnnOLEDB + gcPath & "\" & strInm & _
                    "\inm.mdb", adOpenKeyset, adLockOptimistic, adCmdText
                    
                    If Not rstFD.EOF And Not rstFD.BOF Then 'existen otras facturas registradas
                    
                        MsgBox "No se ha recibido la factura anterior del Inm:" & strInm & vbCrLf & _
                        "ID: '" & flexRemesa(0).TextMatrix(j, 2) & "'" & vbCrLf & "Se desdomiciliar" _
                        & "á", vbInformation, App.ProductName
                        strInm = Desdomiciliar(j)
                        
                    End If
                    
                End If
                '
        End If
        '
        rstFD.Close
            '
            With flexRemesa(0)
                .TextMatrix(j, 5) = Format(0, "#,##0.00")
                .TextMatrix(j, 6) = flexRemesa(0).TextMatrix(j, 4)
                If Not strInm = "" Then
                    With flexRemesa(0)
                        If .TextMatrix(j, 0) = "" Then
                            .TextMatrix(j, 0) = strInm
                            .TextMatrix(j, 1) = Buscar_Cliente("", strInm)
                        End If
                    End With
                    .Row = j
                    .Col = 8
                    If .CellPicture = imgRemesa(1) Then Set .CellPicture = imgRemesa(0)
                    .Col = 1
                    .CellAlignment = flexAlignLeftCenter
                    If optRemesa(0) And Cta_Separada(strInm) And gnIDB > 0 Then 'aplicar IDB
                        curIDB = CCur(.TextMatrix(j, 4)) * gnIDB / 100
                        curNeto = CCur(.TextMatrix(j, 4)) + curIDB
                        .TextMatrix(j, 5) = Format(curIDB, "#,##0.00")
                        .TextMatrix(j, 6) = Format(curNeto, "#,##0.00")
                    End If
                    '
                End If
                '
            End With
            '
            For G = 3 To 5  'Indica Monto Bruto,IDB,Neto Total
                curMonto(G) = curMonto(G) + CCur(flexRemesa(0).TextMatrix(j, G + 1))
                txtRemesa(G - 1) = Format(curMonto(G), "#,##0.00")
            Next
            '
            End If
            
20         j = j + 1
            '
        Loop Until j = flexRemesa(0).Rows
        
        If IsNumeric(txtRemesa(2)) And IsNumeric(txtRemesa(5)) Then
            txtRemesa(6) = Format(CCur(txtRemesa(2)) - CCur(txtRemesa(5)), "#,##0.00")
            If CCur(txtRemesa(6)) <> 0 Then
                MsgBox "Esta remesa no cuadra..", vbExclamation, App.ProductName
            End If
        End If
        Set rstFD = Nothing
    '
        If byServ = 0 Then Call Ampliar_Flex
        Toolbar1.Buttons(6).Enabled = True
        Me.MousePointer = vbDefault
        Call rtnBitacora("Validar Remesa Nº " & txtRemesa(0))
    '
    End Sub

    
    '---------------------------------------------------------------------------------------------
    '   Rutina:     ampliar_flex
    '
    '   Muestra el contenido de todas las columnas del grid
    '---------------------------------------------------------------------------------------------
    Private Sub Ampliar_Flex()
    '
    With flexRemesa(0)
        For i = 0 To .Cols - 1
            If .ColWidth(i) = 0 Then .ColWidth(i) = LoadResString(i + 510)
        Next
        .Width = 10100
        Call Color_Fondo(flexRemesa(0))
        .Col = 0
        .Row = 1
        Call Marcar_Linea(flexRemesa(0), vbGreen)
    End With
    
    '
    End Sub

    '---------------------------------------------------------------------------------------------
    '
    '   Funcion:    Buscar_Cliente
    '
    '   Devuelve el nombre del inmueble o el código del apto. que corresponde
    '---------------------------------------------------------------------------------------------
    Private Function Buscar_Cliente(strCliente$, strCliente1$) As String
    'variables locales
    Dim cnnApto As ADODB.Connection
    Dim rstApto As New ADODB.Recordset
    '
    If strCliente = "" Then
        With FrmAdmin.objRst
            If .State = 0 Then .Open
            .MoveFirst
            .Find "CodInm='" & strCliente1 & "'"
            If Not .EOF Then
                Buscar_Cliente = !Nombre
            Else
                Buscar_Cliente = "NO REGISTRADO"
            End If
        End With
    Else
        'crea una nueva instancia del objeto
        Set cnnApto = New ADODB.Connection
        Set rstApto = New ADODB.Recordset
        '
        cnnApto.Open cnnOLEDB & gcPath & "\" & strCliente1 & "\inm.mdb"
        rstApto.Open "SELECT * FROM Propietarios WHERE Codigo='" & strCliente & "'", _
        cnnApto, adOpenStatic, adLockReadOnly, adCmdText
        Buscar_Cliente = rstApto!Codigo & "--" & rstApto!Nombre
        'los cierra y los descarga de memoria
        rstApto.Close
        cnnApto.Close
        Set rstApto = Nothing
        Set cnnApto = Nothing
        '
    End If
    '
    End Function
    

    '---------------------------------------------------------------------------------------------
    '
    '   Funcion:    Cta_Separada
    '
    '   Entrada:    variable en cadena strCI, contiene el codigo del inmueble
    '
    '   Salida:     Devuelve valor true si la modalidad es cta. separada y False
    '               si es lo contrario
    '---------------------------------------------------------------------------------------------
    Private Function Cta_Separada(strCI$) As Boolean
    '
    With FrmAdmin.objRst
        .MoveFirst
        .Find "CodInm ='" & strCI & "'"
        If !tipoCta = "PARTICULAR" Then
            Cta_Separada = True
        Else
            Cta_Separada = False
        End If
        If strCI = sysCodInm Then Cta_Separada = False
    End With
    '
    End Function
    
    '---------------------------------------------------------------------------------------------
    '
    '   Rutina:     Guardar_Remesa
    '
    '   Procedimiento que guarda el maestro y el detalle de la remesa, previamente
    '   validad, llama la rutina actualizar_Cpp, Procesamiento por lotes
    '
    '---------------------------------------------------------------------------------------------
    Private Sub Guardar_Remesa()
    'variables locales
    Dim strDesDomi$, strRecl$, StrApto$
    Dim intC%
    '
    cnnConexion.BeginTrans  'comienza la transacción
    On Error GoTo RtnError
    mskRemesa.PromptInclude = True  'Guarda la inf. básica de la remesa
    
    Call RtnProUtility("Guardando Encabezado...", 15)
    cnnConexion.Execute "INSERT INTO Remesa (IDRemesa,Titulo,Cargado,Usuario,Fecha,Hora) VALUES" _
    & "(" & txtRemesa(0) & ",'" & txtRemesa(1) & "','01/" & mskRemesa & "','" & gcUsuario & _
    "',Date(),Time());"
    mskRemesa.PromptInclude = False
    '
    For i = 1 To flexRemesa(0).Rows - 1
        Call RtnProUtility("Detalle..." & flexRemesa(0).TextMatrix(i, 2), _
        6000 / flexRemesa(0).Rows * i)
        
        If flexRemesa(0).TextMatrix(i, 3) <> "" And flexRemesa(0).TextMatrix(i, 4) <> "" Then
        strDesDomi = YesNo(i, 7)
        strRecl = YesNo(i, 8)
        
        With flexRemesa(0)
            'Asignar valor variables locales
            intC = InStr(.TextMatrix(i, 1), "--")
            If intC = 0 Then
                StrApto = ""
            Else
                StrApto = Left(.TextMatrix(i, 1), intC - 1)
            End If
            'StrApto = IIf(intC = 0, "", Left(.TextMatrix(I, 1), intC - 1))
            'Agrega el detalle de la remesa
            cnnConexion.Execute "INSERT INTO DetRemesa(IDRemesa,CodInm,Apto,Nis,Periodo,Bruto,I" _
            & "DB,Monto,Ndoc,Reclamo,DesDomi) VALUES(" & txtRemesa(0) & ",'" & .TextMatrix(i, 0) _
            & "','" & StrApto & "','" & .TextMatrix(i, 2) & "','" & .TextMatrix(i, 3) & "','" & _
            .TextMatrix(i, 4) & "','" & .TextMatrix(i, 5) & "','" & .TextMatrix(i, 6) & "','0'," _
            & strDesDomi & "," & strRecl & ");"
            '
        End With
        '
        End If
    '
    Next
    '
RtnError:
    If Err.Number = 0 Then
        cnnConexion.CommitTrans
        Call RtnProUtility("Guardando Cuentas x Pagar...", 0)
        Call Actualizar_Cpp
    Else
        cnnConexion.RollbackTrans
        Call rtnBitacora("Error Remesa: " & txtRemesa(0) & Err.Description)
        MsgBox "Ha ocurrio un error durante el proceso, consulte al administrador del sistema" _
        & vbCrLf & Err.Description, vbInformation, App.ProductName
    End If
    '
    End Sub

    Private Function YesNo(ByVal intFil As Integer, ByVal intCol As Integer) As String
    With flexRemesa(0)
        .Col = intCol
        .Row = intFil
        YesNo = IIf(.CellPicture = imgRemesa(0), "False", "True")
    End With
    End Function

    '---------------------------------------------------------------------------------------------
    '
    '   Rutina:     Actualizar_Cpp
    '
    '   Procedimiento que genera una factura con estatus 'asignado' a la cuenta
    '   de c/condominio según lo relacionado en el grid
    '
    '---------------------------------------------------------------------------------------------
    Private Sub Actualizar_Cpp()
    'Variables locales
    Dim rstCpp As New ADODB.Recordset
    Dim rstTipo As New ADODB.Recordset
    Dim Time1$, strN$, strEstatus$, strProveedor$, strDoc$
    Dim curCheque@
    Dim datV As Date
    Dim strCpp(2) As String
    '
    On Error GoTo SalirCPP
    Call RtnProUtility("Buscanco Detalle de la Remesa...", 15)
    '
    rstCpp.CursorLocation = adUseClient
    rstCpp.Open "DetRemesa", cnnConexion, adOpenKeyset, adLockOptimistic, adCmdTable
    rstCpp.Filter = "IDRemesa=" & txtRemesa(0) & " AND Reclamo=False and Desdomi=false"
    
    If rstCpp.RecordCount > 0 Then rstCpp.MoveFirst
    '
    'Selecciona Codigo de Proveedor, Beneficiario,Descripcion de la factura
    
    rstTipo.Open "SELECT ServiciosTipo.*,Proveedores.* FROM ServiciosTipo INNER JOIN Proveedore" _
    & "s ON Proveedores.Codigo=ServiciosTipo.IDProveedor WHERE ServiciosTipo.IDServicio = " & _
    byServ, cnnConexion, adOpenStatic, adLockReadOnly, adCmdText
    '
    strCpp(0) = rstTipo!IDProveedor
    strCpp(1) = rstTipo!NombProv
    rstTipo.Close
    Set rstTipo = Nothing
    '
    With rstCpp
        '
        Time1 = Trim(CStr(Time))
        If byServ = 0 Then  'el servicio es luz eléctrica
        '
            Do  'Hacer
            '
                Call RtnProUtility("Agregando Cta. x Pagar IDS: " & !NIS, _
                6000 / .RecordCount * .AbsolutePosition)
                
                If Cta_Separada(!CodInm) Then
                    strEstatus = scASIGNADO
                    strProveedor = sysEmpresa
                Else
                    strEstatus = scPAGADO
                    strProveedor = strCpp(1)
                End If
                strCpp(2) = UCase(optRemesa(byServ).Caption & " ID: " & !NIS & " mes: " _
                & !Periodo)
                strDoc = Agregar_Cpp(rstCpp, strCpp(0), strProveedor, strCpp(2), Time1, _
                strEstatus, !Periodo)
                .Update "Ndoc", strDoc
                curCheque = curCheque + !Bruto
                .MoveNext
                '
            Loop Until .EOF 'Hasta fin de archivo
            datV = DateAdd("d", 30, Date)
            strN = FrmFactura.FntStrDoc
            strCpp(2) = UCase(optRemesa(byServ).Caption & " remesa nº " & txtRemesa(0))
            'Agrega una factura por el total de la factura
            cnnConexion.Execute "INSERT INTO Cpp(Tipo,Ndoc,Fact,CodProv,Benef,Detalle,Monto,Ivm" _
            & ",Total,FRecep,Fecr,Fven,CodInm,Moneda,Estatus,Usuario,Freg,Campo1) VALUES('RM','" _
            & strN & "','" & txtRemesa(0) & "','" & strCpp(0) & "','" & strCpp(1) & "','" & _
            strCpp(2) & "','" & curCheque & "',0,'" & curCheque & "',Date(),Date(),'" & _
            datV & "','" & sysCodInm & "','BS','ASIGNADO','" & gcUsuario & "',Date(),'" & Time1 & "')"
            'Agrega la informacion al cargado del cheque
            
        '
        Else    'servicio es teléfono o agua
        
            Do  'hacer
                Call RtnProUtility("Agregando Cta. x Pagar IDS: " & !NIS, 6000 / .RecordCount)
                strCpp(2) = UCase(optRemesa(byServ).Caption & " ID: " & !NIS & " mes: " _
                & !Periodo)
                .Update "NDoc", Agregar_Cpp(rstCpp, strCpp(0), strCpp(1), strCpp(2), Time1, _
                "ASIGNADO", !Periodo)
                .MoveNext
            Loop Until .EOF 'hasta fin de archivo
            
        End If
        .Close
    End With
    Set rstCpp = Nothing
SalirCPP:
    If Err.Number <> 0 Or Actualizar_Facturacion(Time1) Then
        cnnConexion.Execute "DELETE FROM Cpp WHERE Usuario='" & gcUsuario & "' AND Campo2='" _
        & Time1 & "';"
        cnnConexion.Execute "DELETE FROM Remesa WHERE IDRemesa=" & txtRemesa(0)
        Call rtnBitacora("Cpp.- Error Remesa: " & txtRemesa(0) & Err.Description)
    MsgBox "Ha ocurrido un error durante el proceso de Actulización de Cuentas Por Pagar. " & _
        Err.Description, vbExclamation, App.ProductName
    End If
    '
    End Sub

    '---------------------------------------------------------------------------------------------
    '   Funcion:     Actualizar_Facturacion
    '
    '   Procedimiento que actualiza las tablas AsignaGasto,Cargado y GastoNoComun
    '   con la información de las facturas recibidas en la remesa
    '   Procesamiento por lotes
    '---------------------------------------------------------------------------------------------
    Private Function Actualizar_Facturacion(strH$) As Boolean
    '
    Call RtnProUtility("Actualizando Facturación...", 0)
    
    Dim rstRR As New ADODB.Recordset  'variables locales
    Dim rstReintegro As New ADODB.Recordset
    Dim strCG$, strCGNC$, strMsg$, StrDetalle$
    
    rstRR.Open "SELECT * FROM Cpp WHERE Freg=Date() AND Usuario='" & gcUsuario & "' AND Campo1=" _
    & "'" & strH & "' AND Tipo IN ('FC','RM');", cnnConexion, adOpenKeyset, adLockOptimistic, _
    adCmdText
    '
    mskRemesa.PromptInclude = True
    If Not rstRR.EOF Or Not rstRR.BOF Then
        cnnConexion.BeginTrans
        On Error GoTo SalirFact
        '
        With rstRR
            .MoveFirst
            strCGNC = ""
            strMsg = "Por favor ingrese el código de gasto del servicio de " & _
            optRemesa(byServ).Caption & ", correspondiente a los gastos no comunes"
            Do
            'Agregar el registro al cargado
            Call RtnProUtility("Agregando Cargado..Inm: " & !CodInm, _
            6015 / .RecordCount * .AbsolutePosition)
            '
            If !Campo2 = "" Then
                strCG = Buscar_CG(!Campo3)
                'Agregar el registro a AsignaGasto
                cnnConexion.Execute "INSERT INTO AsignaGasto(Ndoc,CodGasto,Cargado,Descripcion," _
                & "Comun,Alicuota,Monto,Usuario,Fecha,Hora) IN '" & gcPath & "\" & !CodInm _
                & "\inm.mdb' VALUES ('" & !NDoc & "','" & strCG & "','01/" & mskRemesa & "','" _
                & !Detalle & "',-1,-1,'" & !Total & "','" & gcUsuario & "',Date(),Time())"
                'Call Reintegro_Presupuestado(!CodInm, !Campo3, !Campo4, !Ndoc)
            Else
                'Asigna el registro a GastoNoComun
              If !Tipo <> "RM" Then   'Si no es la factura global entonces
                If strCGNC = "" Then strCGNC = InputBox(strMsg, "Gasto No Común..")
                cnnConexion.Execute "INSERT INTO GastoNoComun(CodApto,CodGasto,Concepto,Monto," _
                & "Periodo,Fecha,Hora,Usuario) IN '" & gcPath & "\" & !CodInm _
                & "\inm.mdb' VALUES ('" & !Campo2 & "','" & strCGNC & "','" & !Detalle & "','" _
                & !Monto & "','01/" & mskRemesa & "',Date(),Time(),'" & gcUsuario & "');"
              End If
            End If
            '
            cnnConexion.Execute "INSERT INTO Cargado(Ndoc,CodGasto,Detalle,Periodo,Monto,Fecha," _
            & "Hora,Usuario) IN '" & gcPath & "\" & !CodInm & "\inm.mdb' VALUES('" & !NDoc _
            & "','" & strCG & "','" & !Detalle & "','01/" & mskRemesa & "','" & !Total _
            & "',Date(),Time(),'" & gcUsuario & "');"
            
            .MoveNext
            Loop Until .EOF
            '
        End With
    Else
        Actualizar_Facturacion = MsgBox("No se consiguen los registros en Cuentas por Pagar par" _
        & "a" & vbCrLf & "actualizar la Facturación.", vbInformation, App.ProductName)
        Exit Function
    End If
    
    '
SalirFact:
    If Err.Number = 0 Then
        cnnConexion.Execute "UPDATE Cpp SET Campo1='',Campo2='',Campo3='',Campo4='' WHERE Freg=" _
        & "Date() AND Usuario='" & gcUsuario & "' AND Campo1='" & strH & "';"
        cnnConexion.CommitTrans
        Call RtnEstado(6, Toolbar1)
        Call rtnBitacora("Remesa Registrada Nº " & txtRemesa(0))
        If Dir(gcPath & "\TEMPREM00" & byServ & ".log") <> "" Then _
        Kill gcPath & "\TEMPREM00" & byServ & ".log"
        MsgBox "Remesa registrada con éxito...", vbInformation, App.ProductName
        
    Else
        cnnConexion.RollbackTrans
        Call rtnBitacora("AFACT -Error Remesa: " & txtRemesa(0) & Err.Description)
        Actualizar_Facturacion = MsgBox("Ha ocurrido un error durante el proceso de Actualizació" _
        & "n de Facturación. " & Err.Description)
    End If
    mskRemesa.PromptInclude = False
    '
    End Function

    
    Private Function Desdomiciliar(ByVal W As Integer) As String
    flexRemesa(0).Row = W
    flexRemesa(0).Col = 8
    Set flexRemesa(0).CellPicture = imgRemesa(1)
    Desdomiciliar = ""
    End Function

    '---------------------------------------------------------------------------------------------
    '
    '   Funcion:    Buscar_CG
    '
    '   Entrada:    Identificacion de cuenta del cliente
    '
    '---------------------------------------------------------------------------------------------
    Private Function Buscar_CG(IDCS$) As String
    '
    Dim rstCG As New ADODB.Recordset  'variables locales
    '
    '
    rstCG.Open "SELECT * FROM Servicios WHERE IDCS='" & IDCS & "';", _
    cnnConexion, adOpenStatic, adLockReadOnly
    If rstCG.EOF Or rstCG.BOF Then
        Buscar_CG = InputBox("No hay ningún código de gasto asociado con el Servicio de " _
        & optRemesa(byServ).Caption & ", Introduzcalo por favor..")
    Else
        Buscar_CG = rstCG!codGasto
    End If
    rstCG.Close
    Set rstCG = Nothing
    '
    End Function

    '---------------------------------------------------------------------------------------------
    '
    '   Funcion:     Agregar_Cpp
    '
    '   Agregar un registro a la tabla cuentas por pagar
    '---------------------------------------------------------------------------------------------
    Private Function Agregar_Cpp(rst As ADODB.Recordset, StrP$, strB$, strD$, strH$, _
    strE$, Periodo$) As String
    '
    Dim strND$ 'variables locales
    Dim datVen As Date
    '
    With rst
        '
        If Not !DesDomi Or Not !Reclamo Then
            datVen = DateAdd("d", 30, Date) 'Fecha de Vencimiento suma 30 días a la fecha
            strND = FrmFactura.FntStrDoc   'Consecutivo Cpp
            'Guardar registro en Cpp
            cnnConexion.Execute "INSERT INTO Cpp(Tipo,Ndoc,Fact,CodProv,Benef,Detalle,Monto" _
            & ",Ivm,Total,FRecep,Fecr,Fven,CodInm,Moneda,Estatus,Usuario,Freg,Campo1,Campo2,Cam" _
            & "po3,Campo4) VALUES('FC','" & strND & "','" & Right(!NIS, 7) & "','" & StrP & "'," _
            & "'" & strB & "','" & strD & "','" & !Bruto & "','" & !IDB & "','" & !Monto & "',D" _
            & "ate(),Date(),'" & datVen & "','" & !CodInm & "','BS','" & strE & "','" & _
            gcUsuario & "'," & "Date(),'" & strH & "','" & !apto & "','" & !NIS & "','" & _
            Periodo & "');"
        End If
        '
    End With
    Agregar_Cpp = strND
    '
    End Function

    '---------------------------------------------------------------------------------------------
    '   Rutina= RtnAdd
    '
    '
    '---------------------------------------------------------------------------------------------
    Private Sub RtnAdd(intBoton As Integer)
    '
    txtRemesa(0) = ""
    If mskRemesa.PromptInclude Then mskRemesa.PromptInclude = False
    mskRemesa.Text = ""
    mskRemesa.PromptInclude = True
    For i = 2 To 4
        txtRemesa(i) = "0,00"
    Next
    'Configura la presentación del grid
    flexRemesa(0).Rows = 2
    Call rtnLimpiar_Grid(flexRemesa(0))
    Call Presentar_Grid(0)
    If intBoton = 5 Then 'Boton Agregar Registro
        For j = 0 To 2
            If optRemesa(j).Value = True Then Exit For
        Next
        If SSTab1.tab = 1 Then SSTab1.tab = 0
        Dim rstAdd As New ADODB.Recordset
        rstAdd.Open "SELECT Inmueble,Apto,IDCS, '' as Nom FROM ServiciosApto WHERE TipoServ=" & j & " UNIO" _
        & "N SELECT Servicios.Inmueble,'' as Apto,Servicios.IDCS,Inmueble.Nombre FROM Servicios " _
        & " INNER JOIN Inmueble ON Servicios.Inmueble = Inmueble.CodInm WHERE TipoServ=" & j, _
        cnnConexion, adOpenStatic, adLockReadOnly
        rstAdd.Sort = Me.optRemesa(j).Tag
        With rstAdd
            If .RecordCount > 0 Then
                rstRemesa(scInm).Close
                rstRemesa(scInm).Open "SELECT * FROM Servicios WHERE TipoServ=" & j
                .MoveFirst
                flexRemesa(0).Rows = .RecordCount + 1
                j = 0
                Do
                    j = j + 1
                    flexRemesa(0).TextMatrix(j, 0) = !Inmueble
                    flexRemesa(0).TextMatrix(j, 1) = IIf(!apto = "", !Nom, Buscar_Cliente(!apto, !Inmueble))
                    flexRemesa(0).TextMatrix(j, 2) = !IDCS
                    Call iniciar_check(j)
                    .MoveNext
                Loop Until .EOF
                flexRemesa(0).Col = 0
                flexRemesa(0).Row = 1
                Call Marcar_Linea(flexRemesa(0), vbGreen)
            Else
                flexRemesa(0).Rows = 2
                Call rtnLimpiar_Grid(flexRemesa(0))
            End If
            .Close
        End With
        Set rstAdd = Nothing
        SSTab1.TabEnabled(1) = False
        fraRemesa(1).Enabled = True
    Else
        SSTab1.TabEnabled(1) = True
        fraRemesa(1).Enabled = False
    End If
    '
    End Sub
    
    '---------------------------------------------------------------------------------------------
    '   Rutina:     presentar_grid
    '
    '---------------------------------------------------------------------------------------------
    Private Sub Presentar_Grid(ByVal E As Integer)
    '
    Dim intAncho%, intU%
    '
    intU = IIf(E = 0, 510, 519)
    With flexRemesa(E)
        For i = 0 To .Cols - 1
            .Row = 0
            .Col = i
            .ColWidth(i) = LoadResString(i + intU)
            .CellAlignment = flexAlignCenterCenter
            If i > 4 And i <= 6 Then .ColWidth(i) = 0
            intAncho = intAncho + LoadResString(i + intU)
        Next
        .Width = intAncho + 400
        If E = 0 Then
            Call iniciar_check(1)
        Else
            .Row = 1
            .Col = 0
            Call Marcar_Linea(flexRemesa(E), vbGreen)
        End If
    End With
    '
    End Sub

    
    Private Sub Moverse(IntAccion As Integer)
    '
    With rstRemesa(scRemesa)
        '
        .Find "IDremesa=" & flexRemesa(1).TextMatrix(flexRemesa(1).RowSel, 0)
        If .EOF Then
            .MoveFirst
            .Find "IDremesa=" & flexRemesa(1).TextMatrix(flexRemesa(1).RowSel, 0)
        End If
        Select Case IntAccion
            Case 1  'Ir al primer registro
                .MoveFirst
            Case 2  'avanzar al siguiente registro
                .MovePrevious
                If .BOF Then .MoveLast
            Case 3  'apuntar al registro anterior
                .MoveNext
                If .EOF Then .MoveFirst
            Case 4  'ir al ùltimo registro
                .MoveLast
        End Select
    End With
    '
    With flexRemesa(1)
        For i = 1 To .Rows - 1
            If .TextMatrix(i, 0) = rstRemesa(scRemesa)!IDRemesa Then
                Call Color_Fondo(flexRemesa(1))
                .Row = i
                Call Marcar_Linea(flexRemesa(1), vbGreen)
                Exit For
            End If
        Next
    End With
    '
    End Sub

    Private Sub Color_Fondo(ctlGrid As Control)
    '
    With ctlGrid
        .Col = 0
        .Row = 1
        .RowSel = .Rows - 1
        .ColSel = .Cols - 1
        .FillStyle = flexFillRepeat
        .CellBackColor = vbWhite
        .CellFontBold = False
        .FillStyle = flexFillSingle
    End With
    '
    End Sub

    
    '---------------------------------------------------------------------------------------------
    '
    '   Rutina:     Recuperar_rem
    '
    '   Devuelve los últimos registros salvados por el usuario
    '---------------------------------------------------------------------------------------------
    Sub Recuperar_Rem()
    On Error Resume Next
    Dim numFichero%
    Dim strArchivo$
    '
    numFichero = FreeFile
    For i = optRemesa.LBound To optRemesa.UBound
        If optRemesa(i).Value = True Then Exit For
    Next
    strArchivo = gcPath & "\TEMPREM00" & i & ".log"
    '
    Open strArchivo For Input As numFichero
        Do
            Input #numFichero, intL, strDate, curMonto
            With flexRemesa(0)
                .TextMatrix(intL, 3) = strDate
                .TextMatrix(intL, 4) = curMonto
            End With
        Loop Until EOF(numFichero)
    Close numFichero
    '
    End Sub

    
    '---------------------------------------------------------------------------------------------
    '   Rutina: Eliminar_Remesa
    '
    '   Entrada:    NRemesa (ID de la remesa a eliminar)
    '
    '   Revierte todo el proceso de guardar_remesa. Elimina todos los registros
    '---------------------------------------------------------------------------------------------
    Private Sub Eliminar_Remesa(NRemesa As Long)
    'variables locales
    Dim strSQL$, Progres%
    Dim rstRem As New ADODB.Recordset
    '
    On Error GoTo finalizar
    
    Call RtnConfigUtility(True, "Usuario:" & gcUsuario, "Iniciando el proceso...", "Eliminar Remesa Nº" & NRemesa)
    'Selecciona los registros
    strSQL = "SELECT DetRemesa.CodInm, DetRemesa.Ndoc,Remesa.Cargado FROM Remesa INNER JOIN Det" _
    & "Remesa ON Remesa.IDRemesa = DetRemesa.IDRemesa Where Remesa.IDRemesa =" & NRemesa & _
    " ORDER BY DetRemesa.CodInm;"
    '
    With rstRem
        .Open strSQL, cnnConexion, adOpenStatic, adLockReadOnly, adCmdText
        Call RtnProUtility("Seleccionado los registros...", 0)
        .MoveFirst
        cnnConexion.BeginTrans
        Do
            '
            strSQL = gcPath + "\" + !CodInm + "\inm.mdb"
            If Facturado(!CodInm, !Cargado, True) Then Err.Raise 501, , "Período Facturado"
            Progres = .AbsolutePosition * 6015 / .RecordCount
            Call RtnProUtility("Eliminando Gasto Inm:" & !CodInm, Progres)
            'Elimina la asignacion del gasto
            '----------------------------------------------
            cnnConexion.Execute "DELETE FROM AsignaGasto IN '" & strSQL & "' WHERE Ndoc='" _
            & !NDoc & "';"  'AsignaGasto
            '----------------------------------------------
            cnnConexion.Execute "DELETE FROM Cargado IN '" & strSQL & "' WHERE Ndoc='" _
            & !NDoc & "';"  'Cargado
            'elimina la factura de Cpp
            Call RtnProUtility("Eliminando Cpp " & !CodInm & "/" & !NDoc, Progres)
            cnnConexion.Execute "DELETE FROM Cpp WHERE Ndoc='" & !NDoc & "';"
            .MoveNext
        Loop Until .EOF
        .Close
        Call RtnProUtility("Finalizando Proceso...", 6015)
        'elimina la remesa y su detalle
        cnnConexion.Execute "DELETE FROM Remesa WHERE IDRemesa=" & NRemesa
        
finalizar:
        If Err.Number = 0 Then
            cnnConexion.CommitTrans
            flexRemesa(1).RemoveItem (flexRemesa(1).RowSel)
            MsgBox "Finalizado con éxito", vbInformation, App.ProductName
            Call rtnBitacora("Eliminar Remesa Nº" & NRemesa & ",Exitoso....")
        Else
            MsgBox Err.Description, vbExclamation, App.ProductName
            cnnConexion.RollbackTrans
            Call rtnBitacora("Eliminar Remesa Nº" & NRemesa & ",Fallido....")
        End If
        Unload FrmUtility
        '
    End With
    Set rstRem = Nothing
    '
    End Sub

    '---------------------------------------------------------------------------------------------
    '
    '   Función:    Facturado
    '
    '   Entradas Inm(Código del Inmueble), Mes(Período al que se carga la remesa)
    '
    '   Devuelve verdadero si ese perído ya está facturado, de lo contrario retorna
    '   false
    '---------------------------------------------------------------------------------------------
    Private Function Facturado(Inm$, Mes As Date, Optional Del As Boolean) As Boolean
    'variables locales
    Dim rstFact As New ADODB.Recordset, strSQL$
    '
    strSQL = "SELECT MAX(Periodo) FROM Factura WHERE Fact Not Like 'CHD%';"
    '
    With rstFact
        .Open strSQL, cnnOLEDB + gcPath + "\" + Inm + "\inm.mdb", adOpenStatic, adLockReadOnly, _
        adCmdText
    '
        If Mes <= IIf(IsNull(.Fields(0)), "01/01/1975", .Fields(0)) Then
        
            If Del Then
                Facturado = MsgBox("Imposible eliminar esta remesa, ya se ha facturado el inm." _
                & Inm, vbExclamation, App.ProductName)
            Else
                Facturado = MsgBox("Inm:" & Inm & ". Imposible cargar la factura. Período ya fa" _
                & "cturado...será marcada como 'desdomiciliar'", vbInformation, _
                "Período Facturado")
            End If
        
        End If
        .Close  'cierra el objeto
    End With
    'lo descarga
    Set rstFact = Nothing
    
    
    '
    End Function

    '---------------------------------------------------------------------------------------------
    '   Funcion:    Validado
    '
    '   Devuelve True si falta algún valor necesario para procesar la remesa
    '   de lo contrario retorna false
    '---------------------------------------------------------------------------------------------
    Function validado() As Boolean
    'variables locales
    Dim strMsg$, strTitulo, boton As VbMsgBoxStyle, i%
    Dim ctl As Control
    '
    boton = vbInformation
    strTitulo = App.ProductName
    'Valida los datos básicos de la remesa
    If txtRemesa(1) = "" Then
        strMsg = "Falta Descripción de la Remesa.."
        Set ctl = txtRemesa(1)
    ElseIf Left(mskRemesa, 2) > 12 Or Right(mskRemesa, 4) < Year(Date) Then
        strMsg = "Introdujo un período no válido"
        Set ctl = mskRemesa
    ElseIf txtRemesa(0) = "" Then
        strMsg = "Debe introducir un número de identificación de la remesa"
        Set ctl = txtRemesa(0)
    ElseIf CCur(txtRemesa(2)) = 0 Then
        strMsg = "Debe validar para poder guardar..."
        Set ctl = txtRemesa(2)
    ElseIf txtRemesa(5) = "" Then
        strMsg = "Introduzca el monto de la remesa antes de guardar...."
        Set ctl = txtRemesa(5)
    ElseIf txtRemesa(5) <= 0 Then
        strMsg = "Introduzca el monto de la remesa antes de guardar...."
        Set ctl = txtRemesa(5)
    ElseIf CCur(txtRemesa(6)) <> 0 Then
        strMsg = "Esta remesa no cuadra!...."
        Set ctl = txtRemesa(6)
    End If
    If strMsg <> "" Then validado = MsgBox(strMsg, boton, strTitulo)
    '
    If validado Then
        With ctl
            For i = 0 To 500
                .BackColor = IIf(i Mod 2 = 0, vbActiveTitleBar, &HFFFFFF)
                .Refresh
            Next i
            .BackColor = &HFFFFFF
            .Refresh
            .SelStart = 0
            .SelLength = Len(.Text)
        End With
    End If
    '
    End Function


    Private Sub txtRemesa_KeyPress(Index As Integer, KeyAscii As Integer)
    '
    Select Case Index
        '
        Case 5
            If KeyAscii = Asc(".") Then KeyAscii = Asc(",")
            Call Validacion(KeyAscii, "0123456789,")
            If KeyAscii = 13 Then txtRemesa(5) = Format(txtRemesa(5), "#,##0.00")
    End Select
    End Sub
