VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "msadodc.ocx"
Begin VB.Form FrmBancos 
   Caption         =   " "
   ClientHeight    =   45
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   4860
   ControlBox      =   0   'False
   Icon            =   "FrmBancos.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   45
   ScaleWidth      =   4860
   Tag             =   "1"
   WindowState     =   2  'Maximized
   Begin ComctlLib.Toolbar tlbBanco 
      Align           =   1  'Align Top
      Height          =   480
      Left            =   0
      TabIndex        =   9
      Top             =   0
      Width           =   4860
      _ExtentX        =   8573
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
            Object.ToolTipText     =   "Imprimir"
            Object.Tag             =   ""
            ImageIndex      =   10
         EndProperty
         BeginProperty Button11 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Print"
            Object.ToolTipText     =   "Salir"
            Object.Tag             =   ""
            ImageIndex      =   11
         EndProperty
         BeginProperty Button12 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Close"
            Object.Tag             =   ""
            ImageIndex      =   12
         EndProperty
      EndProperty
      BorderStyle     =   1
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
               Picture         =   "FrmBancos.frx":000C
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "FrmBancos.frx":018E
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "FrmBancos.frx":0310
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "FrmBancos.frx":0492
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "FrmBancos.frx":0614
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "FrmBancos.frx":0796
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "FrmBancos.frx":0918
               Key             =   ""
            EndProperty
            BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "FrmBancos.frx":0A9A
               Key             =   ""
            EndProperty
            BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "FrmBancos.frx":0C1C
               Key             =   ""
            EndProperty
            BeginProperty ListImage10 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "FrmBancos.frx":0D9E
               Key             =   ""
            EndProperty
            BeginProperty ListImage11 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "FrmBancos.frx":0F20
               Key             =   ""
            EndProperty
            BeginProperty ListImage12 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "FrmBancos.frx":10A2
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
   Begin TabDlg.SSTab sstBanco 
      Height          =   5295
      Left            =   0
      TabIndex        =   2
      Top             =   600
      Width           =   9615
      _ExtentX        =   16960
      _ExtentY        =   9340
      _Version        =   393216
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
      TabCaption(0)   =   "Datos Generales"
      TabPicture(0)   =   "FrmBancos.frx":1224
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "TxtIna"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "fraBanco(1)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "fraBanco(0)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "txtBanco(5)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).ControlCount=   4
      TabCaption(1)   =   "Datos Administrativos"
      TabPicture(1)   =   "FrmBancos.frx":1240
      Tab(1).ControlEnabled=   0   'False
      Tab(1).ControlCount=   0
      TabCaption(2)   =   "Datos Adicionales"
      TabPicture(2)   =   "FrmBancos.frx":125C
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "AdoBancos(0)"
      Tab(2).Control(1)=   "fraBanco(3)"
      Tab(2).Control(2)=   "fraBanco(4)"
      Tab(2).Control(3)=   "dtgBanco"
      Tab(2).ControlCount=   4
      Begin VB.TextBox txtBanco 
         Height          =   975
         Index           =   5
         Left            =   255
         TabIndex        =   31
         Top             =   3960
         Width           =   9075
      End
      Begin MSDataGridLib.DataGrid dtgBanco 
         Bindings        =   "FrmBancos.frx":1278
         Height          =   3000
         Left            =   -74670
         TabIndex        =   21
         Top             =   690
         Width           =   9060
         _ExtentX        =   15981
         _ExtentY        =   5292
         _Version        =   393216
         AllowUpdate     =   0   'False
         BackColor       =   16777215
         BorderStyle     =   0
         Enabled         =   -1  'True
         HeadLines       =   1
         RowHeight       =   16
         RowDividerStyle =   4
         FormatLocked    =   -1  'True
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "LISTADO DE BANCOS"
         ColumnCount     =   3
         BeginProperty Column00 
            DataField       =   "IDBanco"
            Caption         =   "Código"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   "0%"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   "NombreBanco"
            Caption         =   "Banco"
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
         BeginProperty Column02 
            DataField       =   "Telefono"
            Caption         =   "Teléfono"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   "(0000)-###-##-##"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   0
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            BeginProperty Column00 
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   3030,236
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   2849,953
            EndProperty
         EndProperty
      End
      Begin VB.Frame fraBanco 
         Enabled         =   0   'False
         Height          =   1620
         Index           =   0
         Left            =   270
         TabIndex        =   22
         Top             =   330
         Width           =   9045
         Begin VB.TextBox txtBanco 
            DataField       =   "NombreBanco"
            DataSource      =   "AdoBancos(0)"
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
            Index           =   1
            Left            =   1350
            Locked          =   -1  'True
            TabIndex        =   25
            Top             =   630
            Width           =   7095
         End
         Begin VB.TextBox txtBanco 
            DataField       =   "IDBanco"
            DataSource      =   "AdoBancos(0)"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   315
            Index           =   0
            Left            =   1350
            Locked          =   -1  'True
            TabIndex        =   24
            Top             =   210
            Width           =   1725
         End
         Begin VB.TextBox txtBanco 
            DataField       =   "Telefono"
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "(0000)-###-##-##"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   8202
               SubFormatType   =   0
            EndProperty
            DataSource      =   "AdoBancos(0)"
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
            Index           =   2
            Left            =   1350
            TabIndex        =   23
            Text            =   " "
            Top             =   1050
            Width           =   4905
         End
         Begin VB.Label lblBanco 
            Caption         =   "Nombre:"
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
            Left            =   270
            TabIndex        =   28
            Top             =   720
            Width           =   705
         End
         Begin VB.Label Label3 
            Caption         =   "Teléfonos"
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
            Left            =   210
            TabIndex        =   27
            Top             =   1110
            Width           =   810
         End
         Begin VB.Label lblBanco 
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
            Height          =   210
            Index           =   0
            Left            =   210
            TabIndex        =   26
            Top             =   300
            Width           =   675
         End
      End
      Begin VB.Frame fraBanco 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1050
         Index           =   4
         Left            =   -70860
         TabIndex        =   15
         Top             =   3720
         Width           =   4695
         Begin VB.CommandButton BotBusca 
            Height          =   330
            Left            =   4230
            Picture         =   "FrmBancos.frx":1290
            Style           =   1  'Graphical
            TabIndex        =   18
            ToolTipText     =   "Buscar"
            Top             =   645
            Width           =   375
         End
         Begin VB.TextBox txtBanco 
            BackColor       =   &H00FFFFFF&
            Height          =   315
            Index           =   13
            Left            =   795
            TabIndex        =   17
            Top             =   285
            Width           =   3795
         End
         Begin VB.TextBox txtBanco 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   14
            Left            =   1320
            Locked          =   -1  'True
            TabIndex        =   16
            Top             =   660
            Width           =   705
         End
         Begin VB.Label lblBanco 
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
            Index           =   7
            Left            =   165
            TabIndex        =   20
            Top             =   285
            Width           =   630
         End
         Begin VB.Label lblBanco 
            AutoSize        =   -1  'True
            Caption         =   "Total Bancos :"
            Height          =   195
            Index           =   8
            Left            =   135
            TabIndex        =   19
            Top             =   705
            Width           =   1035
         End
      End
      Begin VB.Frame fraBanco 
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
         Height          =   1065
         Index           =   3
         Left            =   -74475
         TabIndex        =   12
         Top             =   3705
         Width           =   3210
         Begin VB.OptionButton optBanco 
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
            Height          =   225
            Index           =   0
            Left            =   1185
            TabIndex        =   14
            Tag             =   "IDBanco"
            Top             =   345
            Value           =   -1  'True
            Width           =   1335
         End
         Begin VB.OptionButton optBanco 
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
            Height          =   240
            Index           =   1
            Left            =   1185
            TabIndex        =   13
            Tag             =   "NombreBanco"
            Top             =   690
            Width           =   1335
         End
      End
      Begin VB.Frame fraBanco 
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
         Height          =   1680
         Index           =   1
         Left            =   255
         TabIndex        =   4
         Top             =   2025
         Width           =   9075
         Begin MSMask.MaskEdBox mskBanco 
            DataField       =   "Telefono"
            DataSource      =   "AdoBancos(0)"
            Height          =   315
            Index           =   0
            Left            =   6150
            TabIndex        =   29
            Top             =   270
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   556
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   16
            Format          =   "(####)-###-##-##"
            Mask            =   "(####)-###-##-##"
            PromptChar      =   "_"
         End
         Begin VB.TextBox txtBanco 
            DataField       =   "Agencia"
            DataSource      =   "AdoBancos(0)"
            Height          =   315
            Index           =   4
            Left            =   1110
            TabIndex        =   11
            Text            =   " "
            Top             =   1050
            Width           =   3165
         End
         Begin VB.TextBox txtBanco 
            DataField       =   "Contacto"
            DataSource      =   "AdoBancos(0)"
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
            Index           =   3
            Left            =   1110
            TabIndex        =   0
            Top             =   195
            Width           =   3165
         End
         Begin MSDataListLib.DataCombo dtcBanco 
            DataField       =   "Cargo"
            DataSource      =   "AdoBancos(0)"
            Height          =   330
            Left            =   1110
            TabIndex        =   1
            Top             =   615
            Width           =   3165
            _ExtentX        =   5583
            _ExtentY        =   582
            _Version        =   393216
            BackColor       =   -2147483643
            ListField       =   "Cargo"
            BoundColumn     =   "Nombre"
            Text            =   ""
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin MSMask.MaskEdBox mskBanco 
            DataField       =   "Fax"
            DataSource      =   "AdoBancos(0)"
            Height          =   315
            Index           =   1
            Left            =   6150
            TabIndex        =   30
            Top             =   675
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   556
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   16
            Format          =   "(####)-###-##-##"
            Mask            =   "(####)-###-##-##"
            PromptChar      =   "_"
         End
         Begin VB.Label lblBanco 
            Caption         =   "Teléfono :"
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
            Left            =   4665
            TabIndex        =   10
            Top             =   795
            Width           =   855
         End
         Begin VB.Label lblBanco 
            Caption         =   "Fax :"
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
            Left            =   4695
            TabIndex        =   8
            Top             =   330
            Width           =   390
         End
         Begin VB.Label lblBanco 
            Caption         =   "Agencia :"
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
            Left            =   120
            TabIndex        =   7
            Top             =   1087
            Width           =   765
         End
         Begin VB.Label lblBanco 
            Caption         =   "Cargo :"
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
            Left            =   120
            TabIndex        =   6
            Top             =   679
            Width           =   585
         End
         Begin VB.Label lblBanco 
            Caption         =   "Contacto :"
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
            Left            =   120
            TabIndex        =   5
            Top             =   270
            Width           =   870
         End
      End
      Begin VB.TextBox TxtIna 
         DataField       =   "Inactivo"
         Height          =   285
         Left            =   7650
         TabIndex        =   3
         Top             =   975
         Visible         =   0   'False
         Width           =   300
      End
      Begin MSAdodcLib.Adodc AdoBancos 
         Height          =   330
         Index           =   0
         Left            =   -74940
         Top             =   4920
         Visible         =   0   'False
         Width           =   2145
         _ExtentX        =   3784
         _ExtentY        =   582
         ConnectMode     =   0
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
   End
End
Attribute VB_Name = "FrmBancos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

    Private Sub AdoBancos_RecordsetChangeComplete(Index As Integer, _
    ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, _
    adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
    '
    If Index = 0 Then
        If adReason = adRsnUpdate Then txtBanco(14) = AdoBancos(0).Recordset.RecordCount
    End If
    '
    End Sub

    Private Sub BotBusca_Click()
    'variables locales
    Dim Criterio As String
    '
    If txtBanco(13) = "" Then Exit Sub
    If optBanco(0).Value = True Then
        Criterio = "IDbanco=" & txtBanco(13)
    Else
        Criterio = "NombreBanco Like '" & txtBanco(13) & "%'"
    End If
    '
    With AdoBancos(0).Recordset
        .Find Criterio
        If .EOF Or .BOF Then
            .MoveFirst
            .Find Criterio
            If .EOF Then MsgBox "Banco no registrado", vbInformation, App.ProductName
        End If
    End With
    '
    End Sub

    Private Sub Form_Load()
    '
    With dtgBanco
        Set .HeadFont = LetraTitulo(LoadResString(527), 7.5, True)
        Set .Font = LetraTitulo(LoadResString(528), 8)
    End With
    '
    AdoBancos(0).CursorLocation = adUseClient
    AdoBancos(0).LockType = adLockOptimistic
    AdoBancos(0).CursorType = adOpenKeyset
    AdoBancos(0).ConnectionString = cnnOLEDB & mcDatos
    AdoBancos(0).CommandType = adCmdText
    AdoBancos(0).RecordSource = "SELECT * FROM Bancos"
    AdoBancos(0).Refresh
    '
    Call RtnEstado(6, tlbBanco, (AdoBancos(0).Recordset.EOF And AdoBancos(0).Recordset.BOF))
    
    txtBanco(14) = AdoBancos(0).Recordset.RecordCount
    Set dtgBanco.DataSource = AdoBancos(0)
    '
    End Sub

    Private Sub Form_Resize()
    '
    If WindowState <> vbMinimized Then
        With sstBanco
            .Left = (FrmBancos.ScaleWidth / 2) - (.Width / 2)
            .Top = (FrmBancos.ScaleHeight / 2) - (.Height / 2)
        End With
    End If
    '
    End Sub


    Private Sub optBanco_Click(Index As Integer)
    AdoBancos(0).Recordset.Sort = optBanco(Index).Tag
    End Sub

    Private Sub tlbBanco_ButtonClick(ByVal Button As ComctlLib.Button)
    '
    With AdoBancos(0).Recordset
    
        Select Case UCase(Button.Key)
            Case "FIRST"    'IR AL PRIMER REGISTRO
                .MoveFirst
                If .BOF Then .MoveLast
                
            Case "NEXT"
                .MoveNext
                If .EOF Then .MoveFirst
                
            Case "PREVIOUS"
                .MovePrevious
                If .BOF Then .MoveLast
                
            Case "END"
                .MoveLast
                If .EOF Then .MoveFirst
                
            Case "NEW"
                .AddNew
                fraBanco(0).Enabled = True
                fraBanco(1).Enabled = True
                Call RtnEstado(Button.Index, tlbBanco, True)
                
            Case "UNDO" 'CANCELAR
                .CancelUpdate
                fraBanco(0).Enabled = False
                fraBanco(1).Enabled = False
                MsgBox "Cambios cancelados", vbInformation, App.ProductName
                Call RtnEstado(Button.Index, tlbBanco, .EOF And .BOF)
            
            
            Case "CLOSE"    'CERRAR LA VENTANA
                Unload Me
                Set FrmBancos = Nothing
            
            Case "EDIT"
                fraBanco(0).Enabled = True
                fraBanco(1).Enabled = True
                Call RtnEstado(Button.Index, tlbBanco, (.EOF And .BOF))
            
            Case "SAVE"
                .Update
                
                fraBanco(0).Enabled = False
                fraBanco(1).Enabled = False
                Call RtnEstado(Button.Index, tlbBanco, (.EOF And .BOF))
                MsgBox "Cambios efectuado con éxito", vbInformation, App.ProductName
                
        End Select
        
    End With
    End Sub

Private Sub txtBanco_KeyPress(Index%, KeyAscii%): KeyAscii = Asc(UCase(Chr(KeyAscii)))
If optBanco(0).Value = True And Index = 13 Then Call Validacion(KeyAscii, "0123456789")
End Sub

