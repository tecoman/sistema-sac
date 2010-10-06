VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Begin VB.Form frmFicha 
   Caption         =   "Ficha de Centros y Asociaciones de Emigrantes Españoles"
   ClientHeight    =   10290
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15240
   ControlBox      =   0   'False
   Icon            =   "frmFicha.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   10290
   ScaleWidth      =   15240
   WindowState     =   2  'Maximized
   Begin MSAdodcLib.Adodc adoCentro 
      Height          =   375
      Left            =   150
      Top             =   9900
      Visible         =   0   'False
      Width           =   3000
      _ExtentX        =   5292
      _ExtentY        =   661
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   2
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
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=W:\Proyecto\Centros\centros.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=W:\Proyecto\Centros\centros.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "FichaCentro"
      Caption         =   "Registros Principales"
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
   Begin MSComctlLib.ImageList iml 
      Index           =   0
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   20
      ImageHeight     =   20
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   12
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFicha.frx":0442
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFicha.frx":05C4
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFicha.frx":0746
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFicha.frx":08C8
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFicha.frx":0A4A
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFicha.frx":0BCC
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFicha.frx":0CDE
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFicha.frx":0DF0
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFicha.frx":0F72
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFicha.frx":1084
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFicha.frx":1206
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFicha.frx":1318
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList iml 
      Index           =   1
      Left            =   720
      Top             =   15
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   20
      ImageHeight     =   20
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   12
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFicha.frx":149A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFicha.frx":161C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFicha.frx":179E
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFicha.frx":1920
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFicha.frx":1AA2
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFicha.frx":1C26
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFicha.frx":1D3A
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFicha.frx":1E4E
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFicha.frx":1FD0
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFicha.frx":20E4
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFicha.frx":2268
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFicha.frx":237C
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar bHerramientas 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   15240
      _ExtentX        =   26882
      _ExtentY        =   741
      ButtonWidth     =   714
      ButtonHeight    =   688
      AllowCustomize  =   0   'False
      Appearance      =   1
      Style           =   1
      ImageList       =   "iml(0)"
      HotImageList    =   "iml(1)"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   12
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "FIRST"
            Object.ToolTipText     =   "Ir al primer registro"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "BACK"
            Object.ToolTipText     =   "Ir al registro anterior"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "NEXT"
            Object.ToolTipText     =   "Ir al registro siguiente"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "LAST"
            Object.ToolTipText     =   "Ir al último registro"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "NEW"
            Object.ToolTipText     =   "Agregar nuevo registro"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "SAVE"
            Object.ToolTipText     =   "Guardar cambios"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "FIND"
            Object.ToolTipText     =   "Buscar"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "UNDO"
            Object.ToolTipText     =   "Deshacer cambios"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "DELETE"
            Object.ToolTipText     =   "Eliminar"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "EDIT"
            Object.ToolTipText     =   "Editar"
            ImageIndex      =   10
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "PRINT"
            Object.ToolTipText     =   "Imprimir"
            ImageIndex      =   11
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "CLOSE"
            Object.ToolTipText     =   "Cerrar"
            ImageIndex      =   12
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   9030
      Left            =   315
      TabIndex        =   1
      Top             =   1080
      Width           =   14610
      _ExtentX        =   25770
      _ExtentY        =   15928
      _Version        =   393216
      Tabs            =   8
      Tab             =   3
      TabsPerRow      =   4
      TabHeight       =   847
      ShowFocusRect   =   0   'False
      TabCaption(0)   =   "Datos Identificativos"
      TabPicture(0)   =   "frmFicha.frx":2500
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "fra(0)"
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Normativa"
      TabPicture(1)   =   "frmFicha.frx":251C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "fra(1)"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Medios de Financiacion"
      TabPicture(2)   =   "frmFicha.frx":2538
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "fra(2)"
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "Subvenciones Recibidas de Organismos Públicos"
      TabPicture(3)   =   "frmFicha.frx":2554
      Tab(3).ControlEnabled=   -1  'True
      Tab(3).Control(0)=   "fra(9)"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).ControlCount=   1
      TabCaption(4)   =   "Instalaciones y Locales"
      TabPicture(4)   =   "frmFicha.frx":2570
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "fra(5)"
      Tab(4).ControlCount=   1
      TabCaption(5)   =   "Organos de Dirección"
      TabPicture(5)   =   "frmFicha.frx":258C
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "fra(4)"
      Tab(5).ControlCount=   1
      TabCaption(6)   =   "Notas de Interés"
      TabPicture(6)   =   "frmFicha.frx":25A8
      Tab(6).ControlEnabled=   0   'False
      Tab(6).Control(0)=   "fra(6)"
      Tab(6).ControlCount=   1
      TabCaption(7)   =   "Lista de Centros"
      TabPicture(7)   =   "frmFicha.frx":25C4
      Tab(7).ControlEnabled=   0   'False
      Tab(7).ControlCount=   0
      Begin VB.Frame fra 
         Enabled         =   0   'False
         Height          =   7230
         Index           =   0
         Left            =   -74600
         TabIndex        =   108
         Top             =   1350
         Width           =   13635
         Begin VB.TextBox txt 
            DataField       =   "Docimilio"
            DataSource      =   "adoCentro"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Index           =   1
            Left            =   3795
            MaxLength       =   249
            TabIndex        =   119
            Top             =   1515
            Width           =   9240
         End
         Begin VB.TextBox txt 
            DataField       =   "Demarcacion"
            DataSource      =   "adoCentro"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Index           =   2
            Left            =   3795
            MaxLength       =   249
            TabIndex        =   118
            Top             =   3915
            Width           =   9240
         End
         Begin VB.TextBox txt 
            DataField       =   "n_sociose"
            DataSource      =   "adoCentro"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Index           =   3
            Left            =   11550
            TabIndex        =   117
            Top             =   5100
            Width           =   1455
         End
         Begin VB.TextBox txt 
            DataField       =   "n_socios"
            DataSource      =   "adoCentro"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Index           =   4
            Left            =   11550
            TabIndex        =   116
            Top             =   5685
            Width           =   1455
         End
         Begin VB.TextBox txt 
            DataField       =   "Nombre"
            DataSource      =   "adoCentro"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Index           =   0
            Left            =   3795
            MaxLength       =   249
            TabIndex        =   115
            Top             =   915
            Width           =   9225
         End
         Begin VB.TextBox txt 
            DataField       =   "Ciudad"
            DataSource      =   "adoCentro"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Index           =   32
            Left            =   3795
            MaxLength       =   249
            TabIndex        =   114
            Top             =   2115
            Width           =   4920
         End
         Begin VB.TextBox txt 
            DataField       =   "Estado"
            DataSource      =   "adoCentro"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Index           =   33
            Left            =   3795
            MaxLength       =   249
            TabIndex        =   113
            Top             =   2715
            Width           =   4920
         End
         Begin VB.TextBox txt 
            DataField       =   "Telefono"
            DataSource      =   "adoCentro"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Index           =   34
            Left            =   3795
            MaxLength       =   249
            TabIndex        =   112
            Top             =   5100
            Width           =   2655
         End
         Begin VB.TextBox txt 
            DataField       =   "email"
            DataSource      =   "adoCentro"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Index           =   35
            Left            =   3795
            MaxLength       =   249
            TabIndex        =   111
            Top             =   5685
            Width           =   2655
         End
         Begin VB.TextBox txt 
            DataField       =   "Pais"
            DataSource      =   "adoCentro"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Index           =   36
            Left            =   3795
            MaxLength       =   249
            TabIndex        =   110
            Top             =   3285
            Width           =   4920
         End
         Begin VB.TextBox txt 
            DataField       =   "NIF"
            DataSource      =   "adoCentro"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Index           =   37
            Left            =   11565
            TabIndex        =   109
            Top             =   4515
            Width           =   1455
         End
         Begin MSMask.MaskEdBox mskFecha 
            DataField       =   "fecha_fundacion"
            DataSource      =   "adoCentro"
            Height          =   360
            Index           =   0
            Left            =   3795
            TabIndex        =   120
            Top             =   4515
            Width           =   1410
            _ExtentX        =   2487
            _ExtentY        =   635
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   10
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   "dd/mm/yyyy"
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "Nombre de la Entidad:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   0
            Left            =   1185
            TabIndex        =   132
            Top             =   990
            Width           =   2025
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "Domicilio Social:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   1
            Left            =   1755
            TabIndex        =   131
            Top             =   1590
            Width           =   1455
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "Demarcación Consular:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   2
            Left            =   1140
            TabIndex        =   130
            Top             =   3990
            Width           =   2070
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "Fecha de su fundación:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   3
            Left            =   1110
            TabIndex        =   129
            Top             =   4590
            Width           =   2100
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "Números de socios españoles:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   4
            Left            =   8250
            TabIndex        =   128
            Top             =   5175
            Width           =   2715
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "Número total de socios:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   5
            Left            =   8790
            TabIndex        =   127
            Top             =   5760
            Width           =   2175
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "Ciudad:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   49
            Left            =   2520
            TabIndex        =   126
            Top             =   2190
            Width           =   690
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "Estado:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   50
            Left            =   2520
            TabIndex        =   125
            Top             =   2790
            Width           =   690
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "Teléfono /Fax:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   51
            Left            =   1875
            TabIndex        =   124
            Top             =   5175
            Width           =   1320
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "email:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   52
            Left            =   2685
            TabIndex        =   123
            Top             =   5760
            Width           =   525
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "País:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   53
            Left            =   2790
            TabIndex        =   122
            Top             =   3360
            Width           =   420
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "N.I.F. / C.I.F.:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   54
            Left            =   9750
            TabIndex        =   121
            Top             =   4590
            Width           =   1215
         End
      End
      Begin VB.Frame fra 
         Enabled         =   0   'False
         Height          =   7230
         Index           =   1
         Left            =   -74600
         TabIndex        =   79
         Top             =   1365
         Width           =   13635
         Begin VB.TextBox txt 
            DataField       =   "Nota1"
            DataSource      =   "adoCentro"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1215
            Index           =   5
            Left            =   6660
            MultiLine       =   -1  'True
            TabIndex        =   99
            Top             =   2580
            Width           =   6585
         End
         Begin VB.TextBox txt 
            DataField       =   "Normativa"
            DataSource      =   "adoCentro"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   720
            Index           =   6
            Left            =   750
            MaxLength       =   249
            MultiLine       =   -1  'True
            TabIndex        =   98
            Top             =   4560
            Width           =   9225
         End
         Begin VB.TextBox txt 
            DataField       =   "Nota2"
            DataSource      =   "adoCentro"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1155
            Index           =   7
            Left            =   6660
            MultiLine       =   -1  'True
            TabIndex        =   97
            Top             =   5970
            Width           =   6585
         End
         Begin VB.Frame fra 
            BorderStyle     =   0  'None
            Height          =   810
            Index           =   10
            Left            =   1440
            TabIndex        =   94
            Top             =   555
            Width           =   2940
            Begin VB.OptionButton opt 
               Caption         =   "SI"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Index           =   0
               Left            =   180
               TabIndex        =   96
               Tag             =   "Estatutos_si                        "
               Top             =   105
               Width           =   1215
            End
            Begin VB.OptionButton opt 
               Caption         =   "NO"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Index           =   1
               Left            =   180
               TabIndex        =   95
               Tag             =   "Estututos_no"
               Top             =   540
               Width           =   1215
            End
         End
         Begin VB.Frame fra 
            BorderStyle     =   0  'None
            Height          =   1410
            Index           =   11
            Left            =   1290
            TabIndex        =   86
            Top             =   2580
            Width           =   4965
            Begin VB.OptionButton opt 
               Caption         =   "Hospital"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Index           =   5
               Left            =   195
               TabIndex        =   93
               Tag             =   "hospital"
               Top             =   1095
               Width           =   2520
            End
            Begin VB.OptionButton opt 
               Caption         =   "Sociedad Soc. Mutuos"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Index           =   4
               Left            =   195
               TabIndex        =   92
               Tag             =   "soc_soc_mut"
               Top             =   780
               Width           =   2520
            End
            Begin VB.OptionButton opt 
               Caption         =   "Fundación"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Index           =   3
               Left            =   195
               TabIndex        =   91
               Tag             =   "fundacion"
               Top             =   465
               Width           =   2520
            End
            Begin VB.OptionButton opt 
               Caption         =   "Asociación"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Index           =   2
               Left            =   195
               TabIndex        =   90
               Tag             =   "Asociacion"
               Top             =   150
               Width           =   2520
            End
            Begin VB.OptionButton opt 
               Caption         =   "Otros"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Index           =   8
               Left            =   3180
               TabIndex        =   89
               Tag             =   "Caracter_Otras"
               Top             =   780
               Width           =   2520
            End
            Begin VB.OptionButton opt 
               Caption         =   "Club Deportivo"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Index           =   7
               Left            =   3180
               TabIndex        =   88
               Tag             =   "ClubDeportivo"
               Top             =   465
               Width           =   2520
            End
            Begin VB.OptionButton opt 
               Caption         =   "ONG"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Index           =   6
               Left            =   3180
               TabIndex        =   87
               Tag             =   "ONG"
               Top             =   150
               Width           =   2520
            End
         End
         Begin VB.Frame fra 
            BorderStyle     =   0  'None
            Height          =   1110
            Index           =   12
            Left            =   1365
            TabIndex        =   85
            Top             =   6015
            Width           =   5115
         End
         Begin VB.CheckBox chk 
            Caption         =   "Cultural"
            DataField       =   "Cultural"
            DataSource      =   "adoCentro"
            Height          =   315
            Index           =   0
            Left            =   1500
            TabIndex        =   84
            Top             =   6000
            Width           =   1215
         End
         Begin VB.CheckBox chk 
            Caption         =   "Recreativa"
            DataField       =   "Recreativa"
            DataSource      =   "adoCentro"
            Height          =   315
            Index           =   1
            Left            =   1500
            TabIndex        =   83
            Top             =   6390
            Width           =   1215
         End
         Begin VB.CheckBox chk 
            Caption         =   "Deportiva"
            DataField       =   "Deportiva"
            DataSource      =   "adoCentro"
            Height          =   315
            Index           =   2
            Left            =   1500
            TabIndex        =   82
            Top             =   6780
            Width           =   1215
         End
         Begin VB.CheckBox chk 
            Caption         =   "SocioCultural"
            DataField       =   "SocioCultural"
            DataSource      =   "adoCentro"
            Height          =   315
            Index           =   3
            Left            =   4605
            TabIndex        =   81
            Top             =   6000
            Width           =   1260
         End
         Begin VB.CheckBox chk 
            Caption         =   "Otras"
            DataField       =   "Actividad_Otras"
            DataSource      =   "adoCentro"
            Height          =   315
            Index           =   4
            Left            =   4605
            TabIndex        =   80
            Top             =   6390
            Width           =   1215
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Estatutos:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   6
            Left            =   795
            TabIndex        =   107
            Top             =   330
            Width           =   960
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Adjuntas copia de los estatutos compulsada por Autoridad Local."
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
            Left            =   765
            TabIndex        =   106
            Top             =   1620
            Width           =   5295
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Carácter de la Entidad:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   8
            Left            =   795
            TabIndex        =   105
            Top             =   2250
            Width           =   2085
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Nota Aclaratoria:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   9
            Left            =   6720
            TabIndex        =   104
            Top             =   2250
            Width           =   1545
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Normativa por la que se rige:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   10
            Left            =   795
            TabIndex        =   103
            Top             =   4200
            Width           =   2640
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Actividad de la Entidad:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   11
            Left            =   825
            TabIndex        =   102
            Top             =   5670
            Width           =   2175
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Nota Aclaratoria:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   12
            Left            =   6720
            TabIndex        =   101
            Top             =   5670
            Width           =   1545
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "(Legislación del país en base a la que está constituida la Entidad)"
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
            Left            =   3525
            TabIndex        =   100
            Top             =   4215
            Width           =   5280
         End
         Begin VB.Line lin 
            BorderColor     =   &H00808080&
            BorderWidth     =   2
            Index           =   0
            X1              =   150
            X2              =   13290
            Y1              =   2055
            Y2              =   2055
         End
         Begin VB.Line lin 
            BorderColor     =   &H00808080&
            BorderWidth     =   2
            Index           =   1
            X1              =   135
            X2              =   13275
            Y1              =   4080
            Y2              =   4080
         End
         Begin VB.Line lin 
            BorderColor     =   &H00808080&
            BorderWidth     =   2
            Index           =   2
            X1              =   150
            X2              =   13290
            Y1              =   5490
            Y2              =   5490
         End
         Begin VB.Line lin 
            BorderColor     =   &H00FFFFFF&
            Index           =   4
            X1              =   165
            X2              =   13305
            Y1              =   2055
            Y2              =   2055
         End
         Begin VB.Line lin 
            BorderColor     =   &H00FFFFFF&
            Index           =   5
            X1              =   165
            X2              =   13305
            Y1              =   4080
            Y2              =   4080
         End
         Begin VB.Line lin 
            BorderColor     =   &H00FFFFFF&
            Index           =   6
            X1              =   165
            X2              =   13305
            Y1              =   5490
            Y2              =   5490
         End
      End
      Begin VB.Frame fra 
         Enabled         =   0   'False
         Height          =   7245
         Index           =   2
         Left            =   -74600
         TabIndex        =   53
         Top             =   1365
         Width           =   13635
         Begin VB.ComboBox cmb 
            DataField       =   "Peridicidad"
            DataSource      =   "adoCentro"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            ItemData        =   "frmFicha.frx":25E0
            Left            =   2925
            List            =   "frmFicha.frx":25F3
            TabIndex        =   65
            Top             =   1080
            Width           =   2490
         End
         Begin VB.OptionButton opt 
            Caption         =   "SI"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   9
            Left            =   2205
            TabIndex        =   64
            Tag             =   "AyudaSi"
            Top             =   6015
            Width           =   1215
         End
         Begin VB.OptionButton opt 
            Caption         =   "NO"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   10
            Left            =   2205
            TabIndex        =   63
            Tag             =   "AyudaNo"
            Top             =   6435
            Width           =   1215
         End
         Begin VB.TextBox txt 
            Alignment       =   1  'Right Justify
            CausesValidation=   0   'False
            DataField       =   "Importe"
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "#,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   8202
               SubFormatType   =   0
            EndProperty
            DataSource      =   "adoCentro"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Index           =   8
            Left            =   2925
            TabIndex        =   62
            Text            =   "0,00"
            Top             =   1560
            Width           =   2490
         End
         Begin VB.TextBox txt 
            DataField       =   "Clase"
            DataSource      =   "adoCentro"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Index           =   9
            Left            =   2925
            TabIndex        =   61
            Top             =   2040
            Width           =   2490
         End
         Begin VB.TextBox txt 
            DataField       =   "año1"
            DataSource      =   "adoCentro"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Index           =   10
            Left            =   2925
            TabIndex        =   60
            Top             =   3585
            Width           =   1050
         End
         Begin VB.TextBox txt 
            DataField       =   "año2"
            DataSource      =   "adoCentro"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Index           =   11
            Left            =   2925
            TabIndex        =   59
            Top             =   4065
            Width           =   1050
         End
         Begin VB.TextBox txt 
            DataField       =   "año3"
            DataSource      =   "adoCentro"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Index           =   12
            Left            =   2925
            TabIndex        =   58
            Top             =   4545
            Width           =   1050
         End
         Begin VB.TextBox txt 
            Alignment       =   1  'Right Justify
            CausesValidation=   0   'False
            DataField       =   "cuantia1"
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "#,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   2058
               SubFormatType   =   0
            EndProperty
            DataSource      =   "adoCentro"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Index           =   13
            Left            =   5400
            TabIndex        =   57
            Text            =   "0,00"
            Top             =   3585
            Width           =   2490
         End
         Begin VB.TextBox txt 
            Alignment       =   1  'Right Justify
            CausesValidation=   0   'False
            DataField       =   "cuantia2"
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "#,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   2058
               SubFormatType   =   0
            EndProperty
            DataSource      =   "adoCentro"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Index           =   14
            Left            =   5400
            TabIndex        =   56
            Text            =   "0,00"
            Top             =   4065
            Width           =   2490
         End
         Begin VB.TextBox txt 
            Alignment       =   1  'Right Justify
            CausesValidation=   0   'False
            DataField       =   "cuantia3"
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "#,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   2058
               SubFormatType   =   0
            EndProperty
            DataSource      =   "adoCentro"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Index           =   15
            Left            =   5400
            TabIndex        =   55
            Text            =   "0,00"
            Top             =   4545
            Width           =   2490
         End
         Begin VB.TextBox txt 
            Alignment       =   1  'Right Justify
            CausesValidation=   0   'False
            DataField       =   "cuantia"
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "#,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   2058
               SubFormatType   =   0
            EndProperty
            DataSource      =   "adoCentro"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Index           =   16
            Left            =   5400
            TabIndex        =   54
            Text            =   "0,00"
            Top             =   5992
            Width           =   2490
         End
         Begin VB.Line lin 
            BorderColor     =   &H00808080&
            BorderWidth     =   2
            Index           =   9
            X1              =   405
            X2              =   13545
            Y1              =   5130
            Y2              =   5130
         End
         Begin VB.Line lin 
            BorderColor     =   &H00808080&
            BorderWidth     =   2
            Index           =   7
            X1              =   405
            X2              =   13545
            Y1              =   2745
            Y2              =   2745
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "Cuotas de los socios:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   14
            Left            =   345
            TabIndex        =   78
            Top             =   510
            Width           =   1920
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "Periodicidad:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   15
            Left            =   1470
            TabIndex        =   77
            Top             =   1155
            Width           =   1170
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "Importe:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   16
            Left            =   1830
            TabIndex        =   76
            Top             =   1635
            Width           =   810
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "Clases de Cuotas:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   17
            Left            =   1035
            TabIndex        =   75
            Top             =   2115
            Width           =   1605
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "Donaciones de los trés últimos años:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   18
            Left            =   435
            TabIndex        =   74
            Top             =   3060
            Width           =   3360
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "Año 1:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   19
            Left            =   2205
            TabIndex        =   73
            Top             =   3660
            Width           =   615
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "Cuantía 1:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   20
            Left            =   4380
            TabIndex        =   72
            Top             =   3660
            Width           =   945
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "Año 2:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   21
            Left            =   2205
            TabIndex        =   71
            Top             =   4140
            Width           =   615
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "Cuantía 2:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   22
            Left            =   4395
            TabIndex        =   70
            Top             =   4140
            Width           =   945
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "Año 3:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   23
            Left            =   2205
            TabIndex        =   69
            Top             =   4620
            Width           =   615
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "Cuantía 3:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   24
            Left            =   4380
            TabIndex        =   68
            Top             =   4620
            Width           =   945
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "Ayudas o Subvenciones privadas del último año:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   25
            Left            =   510
            TabIndex        =   67
            Top             =   5475
            Width           =   4440
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "Cuantía:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   26
            Left            =   4335
            TabIndex        =   66
            Top             =   6067
            Width           =   765
         End
         Begin VB.Line lin 
            BorderColor     =   &H00FFFFFF&
            Index           =   3
            X1              =   420
            X2              =   13560
            Y1              =   2745
            Y2              =   2745
         End
         Begin VB.Line lin 
            BorderColor     =   &H00FFFFFF&
            Index           =   8
            X1              =   420
            X2              =   13560
            Y1              =   5130
            Y2              =   5130
         End
      End
      Begin VB.Frame fra 
         Enabled         =   0   'False
         Height          =   7230
         Index           =   4
         Left            =   -74610
         TabIndex        =   29
         Top             =   1305
         Width           =   13635
         Begin VB.TextBox txt 
            DataField       =   "Presidente"
            DataSource      =   "adoCentro"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Index           =   22
            Left            =   3240
            MaxLength       =   249
            TabIndex        =   38
            Top             =   240
            Width           =   9225
         End
         Begin VB.TextBox txt 
            DataField       =   "Secretario"
            DataSource      =   "adoCentro"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Index           =   23
            Left            =   3240
            MaxLength       =   249
            TabIndex        =   37
            Top             =   1140
            Width           =   9240
         End
         Begin VB.TextBox txt 
            DataField       =   "vPresidente"
            DataSource      =   "adoCentro"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Index           =   24
            Left            =   3240
            MaxLength       =   249
            TabIndex        =   36
            Top             =   690
            Width           =   9240
         End
         Begin VB.TextBox txt 
            DataField       =   "Tesorero"
            DataSource      =   "adoCentro"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Index           =   25
            Left            =   3240
            MaxLength       =   249
            TabIndex        =   35
            Top             =   1590
            Width           =   9225
         End
         Begin VB.TextBox txt 
            DataField       =   "Organos_Otros"
            DataSource      =   "adoCentro"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Index           =   27
            Left            =   3240
            MaxLength       =   249
            TabIndex        =   34
            Top             =   2055
            Width           =   9240
         End
         Begin VB.TextBox txt 
            DataField       =   "nota4"
            DataSource      =   "adoCentro"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1215
            Index           =   28
            Left            =   5775
            MultiLine       =   -1  'True
            TabIndex        =   33
            Top             =   4560
            Width           =   6585
         End
         Begin VB.TextBox txt 
            DataField       =   "nota3"
            DataSource      =   "adoCentro"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1215
            Index           =   26
            Left            =   5775
            MultiLine       =   -1  'True
            TabIndex        =   32
            Top             =   2850
            Width           =   6720
         End
         Begin VB.TextBox txt 
            DataField       =   "fIncripcion"
            DataSource      =   "adoCentro"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Index           =   30
            Left            =   3180
            MaxLength       =   249
            TabIndex        =   31
            Top             =   6735
            Width           =   9225
         End
         Begin VB.TextBox txt 
            DataField       =   "Federacion"
            DataSource      =   "adoCentro"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Index           =   29
            Left            =   3180
            MaxLength       =   249
            TabIndex        =   30
            Top             =   6300
            Width           =   9225
         End
         Begin MSMask.MaskEdBox mskFecha 
            DataField       =   "fuEleccion"
            DataSource      =   "adoCentro"
            Height          =   360
            Index           =   1
            Left            =   3225
            TabIndex        =   39
            Top             =   2535
            Width           =   1410
            _ExtentX        =   2487
            _ExtentY        =   635
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   10
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   "dd/mm/yyyy"
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox mskFecha 
            DataField       =   "fpEleccion"
            DataSource      =   "adoCentro"
            Height          =   360
            Index           =   2
            Left            =   3240
            TabIndex        =   40
            Top             =   4200
            Width           =   1410
            _ExtentX        =   2487
            _ExtentY        =   635
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   10
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   "dd/mm/yyyy"
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "Presidente:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   37
            Left            =   1815
            TabIndex        =   52
            Top             =   315
            Width           =   1035
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "Vicepresidente:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   38
            Left            =   1455
            TabIndex        =   51
            Top             =   765
            Width           =   1395
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "Secretario:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   39
            Left            =   1860
            TabIndex        =   50
            Top             =   1215
            Width           =   990
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "Tesorero:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   40
            Left            =   1995
            TabIndex        =   49
            Top             =   1665
            Width           =   855
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "Otros Cargos:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   41
            Left            =   1590
            TabIndex        =   48
            Top             =   2130
            Width           =   1260
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "Nota aclaratoria:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   45
            Left            =   5820
            TabIndex        =   47
            Top             =   4260
            Width           =   1515
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "Fecha Próxima Renovación:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   44
            Left            =   330
            TabIndex        =   46
            Top             =   4275
            Width           =   2490
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "Nota aclaratoria:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   43
            Left            =   5820
            TabIndex        =   45
            Top             =   2550
            Width           =   1515
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "Fecha de la última elección:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   42
            Left            =   375
            TabIndex        =   44
            Top             =   2595
            Width           =   2475
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "Fecha(s) de inscripción:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   47
            Left            =   585
            TabIndex        =   43
            Top             =   6810
            Width           =   2115
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "Nombre:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   46
            Left            =   1950
            TabIndex        =   42
            Top             =   6405
            Width           =   765
         End
         Begin VB.Line lin 
            BorderColor     =   &H00FFFFFF&
            Index           =   16
            X1              =   300
            X2              =   13440
            Y1              =   6015
            Y2              =   6015
         End
         Begin VB.Line lin 
            BorderColor     =   &H00808080&
            BorderWidth     =   2
            Index           =   17
            X1              =   315
            X2              =   13455
            Y1              =   6015
            Y2              =   6015
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "  Federación a la que Pertenece la Entidad :  "
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   56
            Left            =   540
            TabIndex        =   41
            Top             =   5880
            Width           =   3885
         End
      End
      Begin VB.Frame fra 
         Enabled         =   0   'False
         Height          =   7230
         Index           =   9
         Left            =   405
         TabIndex        =   15
         Top             =   1300
         Width           =   13635
         Begin VB.Frame fra 
            Height          =   7230
            Index           =   7
            Left            =   -15
            TabIndex        =   16
            Top             =   150
            Width           =   13635
            Begin VB.Frame FrmBusca1 
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   990
               Left            =   5685
               TabIndex        =   22
               Top             =   6135
               Width           =   5640
               Begin VB.TextBox TxtTot 
                  Alignment       =   2  'Center
                  BackColor       =   &H00FFFFFF&
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
                  Left            =   1395
                  Locked          =   -1  'True
                  TabIndex        =   25
                  Top             =   600
                  Width           =   540
               End
               Begin VB.TextBox TxtBus 
                  BackColor       =   &H00FFFFFF&
                  Height          =   285
                  Left            =   990
                  TabIndex        =   24
                  Top             =   285
                  Width           =   4170
               End
               Begin VB.CommandButton BotBusca 
                  Height          =   330
                  Left            =   4785
                  Picture         =   "frmFicha.frx":2629
                  Style           =   1  'Graphical
                  TabIndex        =   23
                  ToolTipText     =   "Buscar"
                  Top             =   600
                  Width           =   375
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  Caption         =   "Total Registros"
                  Height          =   240
                  Index           =   55
                  Left            =   120
                  TabIndex        =   27
                  Top             =   615
                  Width           =   1080
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
                  TabIndex        =   26
                  Top             =   315
                  Width           =   630
               End
            End
            Begin VB.Frame FrmBusca 
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
               Height          =   990
               Left            =   1485
               TabIndex        =   17
               Top             =   6135
               Width           =   3915
               Begin VB.OptionButton OptBusca 
                  Caption         =   "Nombre Centro"
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
                  TabIndex        =   21
                  Top             =   615
                  Width           =   1575
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
                  Left            =   105
                  TabIndex        =   20
                  Top             =   345
                  Value           =   -1  'True
                  Width           =   1575
               End
               Begin VB.OptionButton OptBusca 
                  Caption         =   "Estado"
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
                  Left            =   1875
                  TabIndex        =   19
                  Top             =   630
                  Width           =   1335
               End
               Begin VB.OptionButton OptBusca 
                  Caption         =   "Ciudad"
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
                  Left            =   1875
                  TabIndex        =   18
                  Top             =   360
                  Width           =   1440
               End
            End
            Begin MSDataGridLib.DataGrid Grid 
               Bindings        =   "frmFicha.frx":272B
               Height          =   5370
               Left            =   375
               TabIndex        =   28
               Top             =   495
               Width           =   12750
               _ExtentX        =   22490
               _ExtentY        =   9472
               _Version        =   393216
               AllowUpdate     =   0   'False
               ColumnHeaders   =   -1  'True
               HeadLines       =   2
               RowHeight       =   15
               RowDividerStyle =   5
               FormatLocked    =   -1  'True
               BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Tahoma"
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
               Caption         =   "LISTADO CENSO CENTROS Y ASOCIACIONES DE EMIGRANTES ESPAÑOLES"
               ColumnCount     =   7
               BeginProperty Column00 
                  DataField       =   "IDCentro"
                  Caption         =   "Código"
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
                  DataField       =   "Nombre"
                  Caption         =   "Nombre Centro"
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
                  DataField       =   "Ciudad"
                  Caption         =   "Ciudad"
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
                  DataField       =   "Estado"
                  Caption         =   "Estado"
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
                  DataField       =   "fecha_fundacion"
                  Caption         =   "Fundación"
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
               BeginProperty Column05 
                  DataField       =   "n_sociose"
                  Caption         =   "Socios Españoles"
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
               BeginProperty Column06 
                  DataField       =   "n_socios"
                  Caption         =   "Nº Socios"
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
                     ColumnWidth     =   645,165
                  EndProperty
                  BeginProperty Column01 
                     ColumnWidth     =   4919,811
                  EndProperty
                  BeginProperty Column02 
                     ColumnWidth     =   1769,953
                  EndProperty
                  BeginProperty Column03 
                     ColumnWidth     =   2069,858
                  EndProperty
                  BeginProperty Column04 
                     ColumnWidth     =   959,811
                  EndProperty
                  BeginProperty Column05 
                     Alignment       =   2
                     ColumnWidth     =   959,811
                  EndProperty
                  BeginProperty Column06 
                     Alignment       =   2
                     ColumnWidth     =   870,236
                  EndProperty
               EndProperty
            End
         End
      End
      Begin VB.Frame fra 
         Enabled         =   0   'False
         Height          =   7230
         Index           =   5
         Left            =   -74600
         TabIndex        =   5
         Top             =   1335
         Width           =   13635
         Begin VB.TextBox txt 
            DataField       =   "Local"
            DataSource      =   "adoCentro"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1215
            Index           =   20
            Left            =   690
            MultiLine       =   -1  'True
            TabIndex        =   12
            Top             =   1305
            Width           =   12465
         End
         Begin VB.TextBox txt 
            DataField       =   "Instalacion"
            DataSource      =   "adoCentro"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1215
            Index           =   21
            Left            =   6525
            MultiLine       =   -1  'True
            TabIndex        =   11
            Top             =   4290
            Width           =   6585
         End
         Begin VB.Frame fra 
            BorderStyle     =   0  'None
            Height          =   1950
            Index           =   17
            Left            =   660
            TabIndex        =   6
            Top             =   3855
            Width           =   2250
            Begin VB.OptionButton opt 
               Alignment       =   1  'Right Justify
               Caption         =   "Otros:"
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
               Index           =   24
               Left            =   0
               TabIndex        =   10
               Tag             =   "Local_Otros"
               Top             =   1305
               Width           =   1890
            End
            Begin VB.OptionButton opt 
               Alignment       =   1  'Right Justify
               Caption         =   "Propiedad:"
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
               Index           =   23
               Left            =   0
               TabIndex        =   9
               Tag             =   "Propiedad"
               Top             =   870
               Width           =   1890
            End
            Begin VB.OptionButton opt 
               Alignment       =   1  'Right Justify
               Caption         =   "Cesión:"
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
               Index           =   22
               Left            =   0
               TabIndex        =   8
               Tag             =   "Cesion"
               Top             =   435
               Width           =   1890
            End
            Begin VB.OptionButton opt 
               Alignment       =   1  'Right Justify
               Caption         =   "Alquiler:"
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
               Index           =   21
               Left            =   0
               TabIndex        =   7
               Tag             =   "Alquiler"
               Top             =   0
               Width           =   1890
            End
         End
         Begin VB.Line lin 
            BorderColor     =   &H00808080&
            BorderWidth     =   2
            Index           =   15
            X1              =   270
            X2              =   13410
            Y1              =   3210
            Y2              =   3210
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "Breve descripción de las instalaciones:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   35
            Left            =   345
            TabIndex        =   14
            Top             =   750
            Width           =   3465
         End
         Begin VB.Line lin 
            BorderColor     =   &H00FFFFFF&
            Index           =   14
            X1              =   285
            X2              =   13425
            Y1              =   3210
            Y2              =   3210
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Nota Aclaratoria:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   36
            Left            =   6585
            TabIndex        =   13
            Top             =   3915
            Width           =   1545
         End
      End
      Begin VB.Frame fra 
         Enabled         =   0   'False
         Height          =   7230
         Index           =   6
         Left            =   -74535
         TabIndex        =   2
         Top             =   1410
         Width           =   13635
         Begin VB.TextBox txt 
            DataField       =   "Notas"
            DataSource      =   "adoCentro"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   3255
            Index           =   31
            Left            =   750
            MultiLine       =   -1  'True
            TabIndex        =   3
            Top             =   1110
            Width           =   12315
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "Notas de Interés:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   48
            Left            =   345
            TabIndex        =   4
            Top             =   540
            Width           =   1590
         End
      End
   End
End
Attribute VB_Name = "frmFicha"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub BotBusca_Click()
'variables locales
Dim strCriterio As String
'
If OptBusca(0) Then
    strCriterio = "IDCentro ='" & Trim(TxtBus) & "'"
ElseIf OptBusca(1) Then
    strCriterio = "Nombre Like '*" & Trim(TxtBus) & "*'"
ElseIf OptBusca(2) Then
    strCriterio = "Estado Like '*" & Trim(TxtBus) & "*'"
ElseIf OptBusca(3) Then
    strCriterio = "Ciudad ='" & Trim(TxtBus) & "'"
End If
With adoCentro.Recordset
    .Find strCriterio
    If .EOF Then
        .MoveFirst
        .Find strCriterio
        If .EOF Then
            MsgBox "No existe coincidencia " & strCriterio, vbInformation, App.ProductName
            .MoveFirst
        End If
    End If
End With
'
End Sub

Private Sub adoCentro_WillMove(ByVal adReason As ADODB.EventReasonEnum, _
adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
'
On Error Resume Next

With adoCentro.Recordset
    For i = 0 To opt.UBound
        opt(i).Value = .Fields(opt(i).Tag)
    Next
End With
'
End Sub


Private Sub bHerramientas_ButtonClick(ByVal Button As MSComctlLib.Button)
'variables locales
Dim i%
Dim msg As String
Dim Respuesta As Long
'
With adoCentro.Recordset

    Select Case Button.Key
    
        Case "FIRST"
            .MoveFirst
        
        Case "BACK"
            .MovePrevious
            If .BOF Then .MoveLast
        
        Case "NEXT"
            .MoveNext
            If .EOF Then .MoveFirst
        
        Case "LAST"
            .MoveLast
            
        Case "NEW"
        
            For i = 0 To 6: fra(i).Enabled = True
            Next
            .AddNew
            For i = 0 To opt.UBound: opt(i).Value = False
            Next
            SSTab1.Tab = 0
            txt(0).SetFocus
            Call Estado_BarraHerramientas(bHerramientas, "NEW", False)
            
        Case "SAVE"
            'valida los campos mínimos requerios
            If validar_guardar Then Exit Sub
            
            For i = 0 To 6: fra(i).Enabled = False
            Next
            For i = 0 To mskFecha.UBound: mskFecha(i).PromptInclude = True
            Next
            Set Grid.DataSource = Nothing
            For i = 0 To opt.UBound: .Fields(Trim(opt(i).Tag)) = opt(i).Value
            Next
            
            If .EditMode = adEditAdd Then !IDCentro = NCentro
            !Usuario = gcUsuario
            !Freg = Date
            
            .Update
            
            Set Grid.DataSource = adoCentro
            MsgBox "Registro procesdo con éxito", vbInformation, Me.Caption
            For i = 0 To mskFecha.UBound: mskFecha(i).PromptInclude = False
            Next
            Call Estado_BarraHerramienta(bHerramientas, "SAVE", Not .EOF And Not .BOF)
            
        Case "EDIT"
            For i = 0 To 6: fra(i).Enabled = True
            Next
            Call Estado_BarraHerramienta(bHerramientas, "EDIT", False)
        
        Case "UNDO"
        
            For i = 0 To 6: fra(i).Enabled = False
            Next
            For i = 0 To 2: mskFecha(i).PromptInclude = True
            Next
            Set Grid.DataSource = Nothing
            .CancelUpdate
            For i = 0 To 2: mskFecha(i).PromptInclude = False
            Next
            Set Grid.DataSource = adoCentro
            MsgBox "Cambios cancelados..", vbInformation, Me.Caption
            Call Estado_BarraHerramienta(bHerramientas, "UNDO", False)
        
        Case "DELETE"
            msg = "Desea elmininar el centro: " & .Fields("Nombre")
            Respuesta = MsgBox(msg, vbYesNo + vbQuestion, Me.Caption)
            If Respuesta = vbYes Then
                .Delete
                .Requery
                MsgBox "Centro Eliminado", vbInformation, Me.Caption
            End If
            Call Estado_BarraHerramienta(bHerramientas, "DELETE", Not .EOF And Not .BOF)
                
        Case "PRINT"
            MsgBox "Opción no disponible. Consulte con el administrador del sistema", _
            vbInformation, Me.Caption
            
        Case "CLOSE"
            Unload Me
        
        Case "FIND"
            SSTab1.Tab = 7
            
    End Select
    '
End With
'
End Sub

Function NCentro() As Long
'variables locales
Dim rstNum As New ADODB.Recordset
'
rstNum.Open "SELECT Max(IDCentro) FROM FichaCentro", adoCentro.ConnectionString, adOpenKeyset, _
adLockOptimistic, adCmdText
'
'If IsNull(rstNum.Fields(0)) Then

'    NCentro = 1
'Else

    NCentro = 1 + rstNum.Fields(0)
'End If
'
rstNum.Close
Set rstNum = Nothing
'
End Function

Private Sub Form_Load()
'CONFIGURA EL CONTROL ADO'S
With Me.adoCentro
    .CursorLocation = adUseClient
    .Mode = adModeShareDenyNone
    .ConnectionString = strProvider & gcPath & "\centros.mdb"
    .LockType = adLockOptimistic
    .CommandType = adCmdTxt
    .RecordSource = "SELECT * FROM FichaCentro WHERE Pais='" & gcPais & "' ORDER BY IDCentro"
    .Refresh
    
End With
If Not adoCentro.Recordset.BOF And Not adoCentro.Recordset.EOF Then
'
    For i = 1 To 4: bHerramientas.Buttons(i).Enabled = True
    Next
    bHerramientas.Buttons("DELETE").Enabled = True
    bHerramientas.Buttons("EDIT").Enabled = True
    bHerramientas.Buttons("PRINT").Enabled = True
    '
End If
'
End Sub

Private Sub mskFecha_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii > 26 And InStr("0123456789", Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub

Private Sub SSTab1_DblClick()
If SSTab1.Tab = 7 Then
   Set Grid.DataSource = Nothing
   Set Grid.DataSource = adoCentro
   Grid.Refresh
   TxtTot = adoCentro.Recordset.RecordCount
   TxtTot.Refresh
End If


End Sub

Private Sub SSTab1_GotFocus()
If SSTab1.Tab = 7 Then
   Set Grid.DataSource = Nothing
   Set Grid.DataSource = adoCentro
   Grid.Refresh
   TxtTot = adoCentro.Recordset.RecordCount
   TxtTot.Refresh
End If
End Sub

Private Sub txt_GotFocus(Index As Integer)
    If IsNumeric(txt(Index)) Then txt(Index) = "": txt(Index).Refresh
    
End Sub

Private Sub txt_KeyPress(Index As Integer, KeyAscii As Integer)
If Index = 35 Then
    KeyAscii = Asc(LCase(Chr(KeyAscii)))
Else
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End If
Select Case Index

    Case 3, 4, 10, 11, 12
    
        If KeyAscii > 26 And InStr("0123456789", Chr(KeyAscii)) = 0 Then KeyAscii = 0
        
    
    Case 13, 14, 15, 16, 8
        If KeyAscii = Asc(".") Then KeyAscii = Asc(",")
        If KeyAscii > 26 And InStr("0123456789,", Chr(KeyAscii)) = 0 Then KeyAscii = 0
        If KeyAscii = 13 Then
            txt(Index) = Format(txt(Index), "#,##0.00")
        End If
        '
End Select

End Sub


Function validar_guardar() As Boolean
'variables locales
Dim strFalta As String
'
If txt(0) = "" Then
    strFalta = "- Falta el Nombre del Centro"
End If
If txt(1) = "" Then
    strFalta = IIf(strFalta = "", "", strFalta & vbCrLf) & "- Falta Domicilio Social"
End If
If txt(2) = "" Then
    strFalta = IIf(strFalta = "", "", strFalta & vbCrLf) & "- Falta Demarcación Consular"
End If

If strFalta <> "" Then
    strFalta = "No se puede procesar el registro: " & vbCrLf & vbCrLf & strFalta
    validar_guardar = MsgBox(strFalta, vbInformation, Me.Caption)
End If
End Function
