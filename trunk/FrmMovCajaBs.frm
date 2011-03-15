VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form FrmMovCajaBs 
   Caption         =   "Cobranza por Caja [Bolívares]"
   ClientHeight    =   45
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   2475
   ControlBox      =   0   'False
   Icon            =   "FrmMovCajaBs.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   11460
   ScaleWidth      =   18960
   WindowState     =   2  'Maximized
   Begin ComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   480
      Left            =   0
      TabIndex        =   97
      Top             =   0
      Width           =   18960
      _ExtentX        =   33443
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
            Key             =   "New"
            Object.ToolTipText     =   "Nuevo (Ctrl + N)"
            Object.Tag             =   ""
            ImageIndex      =   5
         EndProperty
         BeginProperty Button6 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Save"
            Object.ToolTipText     =   "Guardar (Ctl + G)"
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
      MousePointer    =   99
      MouseIcon       =   "FrmMovCajaBs.frx":000C
      Begin VB.TextBox Txt 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   8202
            SubFormatType   =   0
         EndProperty
         DataSource      =   "ADOcontrol(0)"
         Enabled         =   0   'False
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
         Index           =   14
         Left            =   5235
         TabIndex        =   152
         Text            =   "Para guardar una operación puede presionar Ctl + G"
         Top             =   90
         Width           =   5145
      End
   End
   Begin MSComCtl2.MonthView MntCalendar 
      Height          =   2310
      Left            =   4125
      TabIndex        =   16
      Tag             =   "0"
      Top             =   5130
      Visible         =   0   'False
      Width           =   2355
      _ExtentX        =   4154
      _ExtentY        =   4075
      _Version        =   393216
      ForeColor       =   -2147483630
      BackColor       =   -2147483633
      BorderStyle     =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ShowToday       =   0   'False
      StartOfWeek     =   84803585
      TitleBackColor  =   -2147483646
      TitleForeColor  =   16777215
      CurrentDate     =   37319
   End
   Begin MSAdodcLib.Adodc ADOcontrol 
      Height          =   330
      Index           =   0
      Left            =   180
      Top             =   8055
      Visible         =   0   'False
      Width           =   2700
      _ExtentX        =   4763
      _ExtentY        =   582
      ConnectMode     =   3
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   1
      LockType        =   3
      CommandType     =   1
      CursorOptions   =   0
      CacheSize       =   20
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
      Caption         =   "AdoCajaDia"
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
   Begin TabDlg.SSTab SSTab1 
      Height          =   6420
      Left            =   75
      TabIndex        =   20
      Top             =   495
      Width           =   11070
      _ExtentX        =   19526
      _ExtentY        =   11324
      _Version        =   393216
      Tabs            =   4
      TabsPerRow      =   4
      TabHeight       =   520
      ShowFocusRect   =   0   'False
      MouseIcon       =   "FrmMovCajaBs.frx":0326
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Datos Generales     "
      TabPicture(0)   =   "FrmMovCajaBs.frx":0342
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "FmeCuentas"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame1"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Lista"
      TabPicture(1)   =   "FrmMovCajaBs.frx":035E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "DataGrid1"
      Tab(1).Control(1)=   "Frame2"
      Tab(1).Control(2)=   "FrmBusca1"
      Tab(1).Control(3)=   "FrmBusca"
      Tab(1).ControlCount=   4
      TabCaption(2)   =   "Deducciones"
      TabPicture(2)   =   "FrmMovCajaBs.frx":037A
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "frame3(3)"
      Tab(2).Control(1)=   "frame3(1)"
      Tab(2).Control(2)=   "frame3(2)"
      Tab(2).Control(3)=   "Command3(0)"
      Tab(2).Control(4)=   "Command3(1)"
      Tab(2).Control(5)=   "AdoDeducciones"
      Tab(2).Control(6)=   "Winsock1"
      Tab(2).Control(7)=   "dtc"
      Tab(2).Control(8)=   "ADOcontrol(4)"
      Tab(2).Control(9)=   "Label16(0)"
      Tab(2).Control(10)=   "Label16(2)"
      Tab(2).ControlCount=   11
      TabCaption(3)   =   "Lista Cheques"
      TabPicture(3)   =   "FrmMovCajaBs.frx":0396
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "frame3(4)"
      Tab(3).ControlCount=   1
      Begin VB.Frame Frame1 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         ForeColor       =   &H00400000&
         Height          =   3495
         Left            =   180
         TabIndex        =   49
         Top             =   1950
         Width           =   7890
         Begin MSFlexGridLib.MSFlexGrid FlexFacturas 
            Height          =   3330
            Left            =   15
            TabIndex        =   50
            Tag             =   "Caja"
            Top             =   120
            Width           =   7860
            _ExtentX        =   13864
            _ExtentY        =   5874
            _Version        =   393216
            Cols            =   9
            FixedCols       =   0
            BackColorFixed  =   -2147483646
            ForeColorFixed  =   -2147483639
            BackColorSel    =   65280
            ForeColorSel    =   0
            GridColor       =   -2147483645
            AllowBigSelection=   0   'False
            FocusRect       =   2
            HighLight       =   2
            ScrollBars      =   2
            AllowUserResizing=   1
            MousePointer    =   99
            FormatString    =   "Factura |^Periodo |Facturado |Abonado |Saldo |Acumulado |Cancelar |Ded.|Bs"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MouseIcon       =   "FrmMovCajaBs.frx":03B2
         End
      End
      Begin VB.Frame frame3 
         BackColor       =   &H80000002&
         BorderStyle     =   0  'None
         Height          =   2610
         Index           =   3
         Left            =   -72930
         TabIndex        =   26
         Top             =   2385
         Visible         =   0   'False
         Width           =   6165
         Begin VB.Timer TimDemonio 
            Left            =   5400
            Top             =   315
         End
         Begin VB.Label Label16 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Haga Clik Aqui para cancelar peticion:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000E&
            Height          =   195
            Index           =   15
            Left            =   3150
            MouseIcon       =   "FrmMovCajaBs.frx":0514
            MousePointer    =   99  'Custom
            TabIndex        =   27
            Top             =   2130
            Width           =   2700
         End
         Begin VB.Label Label16 
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   $"FrmMovCajaBs.frx":0666
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000009&
            Height          =   2310
            Index           =   14
            Left            =   135
            TabIndex        =   28
            Top             =   165
            Width           =   5835
         End
      End
      Begin VB.Frame FrmBusca 
         Caption         =   "Ordenar y Buscar por:"
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
         Height          =   1800
         Left            =   -74670
         TabIndex        =   57
         Top             =   4410
         Width           =   2730
         Begin VB.OptionButton OptBusca 
            Caption         =   "Fecha y Hora"
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
            Left            =   210
            TabIndex        =   135
            Tag             =   "IndiceMovimientoCaja"
            Top             =   330
            Value           =   -1  'True
            Width           =   2370
         End
         Begin VB.OptionButton OptBusca 
            Caption         =   "Caja"
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
            Index           =   2
            Left            =   195
            TabIndex        =   128
            Tag             =   "Caja"
            Top             =   1095
            Width           =   2370
         End
         Begin VB.OptionButton OptBusca 
            Caption         =   "Inmueble"
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
            Left            =   210
            TabIndex        =   59
            Tag             =   "InmuebleMovimientoCaja"
            Top             =   1485
            Width           =   2370
         End
         Begin VB.OptionButton OptBusca 
            Caption         =   "Forma de Pago"
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
            Left            =   210
            TabIndex        =   58
            Tag             =   "FormaPagoMovimientoCaja"
            Top             =   720
            Width           =   2370
         End
      End
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
         Height          =   1800
         Left            =   -71880
         TabIndex        =   52
         Top             =   4425
         Width           =   5520
         Begin VB.TextBox Text3 
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
            Height          =   315
            Left            =   1305
            TabIndex        =   137
            Top             =   1365
            Visible         =   0   'False
            Width           =   1515
         End
         Begin VB.TextBox Text2 
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
            Height          =   315
            Left            =   150
            TabIndex        =   136
            Top             =   1365
            Visible         =   0   'False
            Width           =   1005
         End
         Begin VB.TextBox TxtBus 
            BackColor       =   &H00FFFFFF&
            Height          =   285
            Left            =   1020
            TabIndex        =   54
            Top             =   285
            Width           =   4335
         End
         Begin VB.TextBox TxtTot 
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
            ForeColor       =   &H000000FF&
            Height          =   285
            Left            =   1320
            Locked          =   -1  'True
            TabIndex        =   53
            Top             =   825
            Width           =   690
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
            TabIndex        =   56
            Top             =   315
            Width           =   630
         End
         Begin VB.Label Label17 
            AutoSize        =   -1  'True
            Caption         =   "Total Bancos :"
            Height          =   195
            Left            =   165
            TabIndex        =   55
            Top             =   870
            Width           =   1035
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Filtrar Por:"
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
         Height          =   1800
         Left            =   -66285
         TabIndex        =   51
         Top             =   4440
         Width           =   2220
         Begin VB.OptionButton OptBusca 
            Caption         =   "Ninguno"
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
            Index           =   7
            Left            =   225
            TabIndex        =   141
            Top             =   330
            Value           =   -1  'True
            Width           =   1785
         End
         Begin VB.OptionButton OptBusca 
            Caption         =   "Inmueble"
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
            Left            =   225
            TabIndex        =   140
            Tag             =   "InmuebleMovimientoCaja"
            Top             =   675
            Width           =   1785
         End
         Begin VB.OptionButton OptBusca 
            Caption         =   "Forma de Pago"
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
            Index           =   4
            Left            =   225
            TabIndex        =   139
            Tag             =   "FormaPagoMovimientoCaja"
            Top             =   1020
            Width           =   1785
         End
         Begin VB.OptionButton OptBusca 
            Caption         =   "Caja"
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
            Left            =   225
            TabIndex        =   138
            Tag             =   "Caja"
            Top             =   1365
            Width           =   1785
         End
      End
      Begin VB.Frame frame3 
         Caption         =   "Detalle Deducciones"
         Height          =   2910
         Index           =   1
         Left            =   -74790
         TabIndex        =   46
         Top             =   2610
         Width           =   10575
         Begin MSFlexGridLib.MSFlexGrid FlexDeducciones 
            Height          =   2295
            Left            =   270
            TabIndex        =   47
            Tag             =   "1280|6850|1800|0"
            Top             =   360
            Width           =   10170
            _ExtentX        =   17939
            _ExtentY        =   4048
            _Version        =   393216
            Cols            =   4
            FixedCols       =   0
            BackColorFixed  =   -2147483646
            ForeColorFixed  =   -2147483639
            BackColorBkg    =   -2147483636
            FocusRect       =   2
            HighLight       =   2
            MergeCells      =   2
            BorderStyle     =   0
            Appearance      =   0
            FormatString    =   "^Cod.Gasto |<Descripcion  |>Monto|"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
      End
      Begin VB.Frame frame3 
         Height          =   1870
         Index           =   2
         Left            =   -74790
         TabIndex        =   31
         Top             =   495
         Width           =   10560
         Begin VB.Frame Frame4 
            Caption         =   "Detalle Factura:"
            Height          =   1455
            Index           =   0
            Left            =   7560
            TabIndex        =   32
            Top             =   200
            Width           =   2775
            Begin VB.Label Label16 
               BackColor       =   &H80000005&
               BorderStyle     =   1  'Fixed Single
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Index           =   13
               Left            =   1080
               TabIndex        =   36
               Top             =   960
               Width           =   1365
            End
            Begin VB.Label Label16 
               BackColor       =   &H80000005&
               BorderStyle     =   1  'Fixed Single
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Index           =   12
               Left            =   1080
               TabIndex        =   35
               Top             =   360
               Width           =   1365
            End
            Begin VB.Label Label16 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Periodo:"
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
               Left            =   195
               TabIndex        =   34
               Top             =   960
               Width           =   675
            End
            Begin VB.Label Label16 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "N°:"
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
               Left            =   600
               TabIndex        =   33
               Top             =   400
               Width           =   270
            End
         End
         Begin VB.Label Label16 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Caja:"
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
            Left            =   855
            TabIndex        =   45
            Top             =   1365
            Width           =   390
         End
         Begin VB.Label Label16 
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Index           =   8
            Left            =   2280
            TabIndex        =   44
            Top             =   1315
            Width           =   2685
         End
         Begin VB.Label Label16 
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Index           =   7
            Left            =   1320
            TabIndex        =   43
            Top             =   1315
            Width           =   885
         End
         Begin VB.Label Label16 
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Index           =   6
            Left            =   2280
            TabIndex        =   42
            Top             =   815
            Width           =   4965
         End
         Begin VB.Label Label16 
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Index           =   5
            Left            =   1320
            TabIndex        =   41
            Top             =   815
            Width           =   885
         End
         Begin VB.Label Label16 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Propietario:"
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
            Left            =   240
            TabIndex        =   40
            Top             =   860
            Width           =   1000
         End
         Begin VB.Label Label16 
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Index           =   3
            Left            =   2280
            TabIndex        =   39
            Top             =   315
            Width           =   4965
         End
         Begin VB.Label LblCodInm 
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   1320
            TabIndex        =   38
            Top             =   315
            Width           =   885
         End
         Begin VB.Label Label16 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Inmueble :"
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
            Left            =   240
            TabIndex        =   37
            Top             =   360
            Width           =   1000
         End
      End
      Begin VB.CommandButton Command3 
         Caption         =   "&Salir"
         Height          =   570
         Index           =   0
         Left            =   -73125
         TabIndex        =   30
         Top             =   5670
         Width           =   1500
      End
      Begin VB.CommandButton Command3 
         Caption         =   "A&ceptar"
         Enabled         =   0   'False
         Height          =   570
         Index           =   1
         Left            =   -74775
         TabIndex        =   29
         Top             =   5670
         Width           =   1500
      End
      Begin VB.Frame frame3 
         Height          =   6600
         Index           =   4
         Left            =   -74805
         TabIndex        =   21
         Top             =   435
         Width           =   10710
         Begin VB.CommandButton Command3 
            Caption         =   "Reimprimir Voucher"
            Enabled         =   0   'False
            Height          =   795
            Index           =   5
            Left            =   9480
            Style           =   1  'Graphical
            TabIndex        =   147
            Top             =   5700
            Width           =   1020
         End
         Begin VB.CommandButton Command3 
            Caption         =   "&Depositar"
            Height          =   795
            Index           =   2
            Left            =   4980
            Style           =   1  'Graphical
            TabIndex        =   121
            Top             =   5700
            Width           =   1500
         End
         Begin VB.CommandButton Command3 
            Caption         =   "&Actualizar Depósito"
            Height          =   795
            Index           =   4
            Left            =   7980
            Style           =   1  'Graphical
            TabIndex        =   127
            Top             =   5700
            Width           =   1500
         End
         Begin VB.CommandButton Command3 
            Caption         =   "&Consultar Depósito"
            Height          =   795
            Index           =   3
            Left            =   6480
            Style           =   1  'Graphical
            TabIndex        =   122
            Top             =   5700
            Width           =   1500
         End
         Begin VB.Frame frame3 
            Caption         =   "Detalle Deposito"
            Height          =   5280
            Index           =   7
            Left            =   4980
            TabIndex        =   104
            Top             =   270
            Width           =   5520
            Begin VB.TextBox Txt 
               Alignment       =   1  'Right Justify
               BeginProperty DataFormat 
                  Type            =   1
                  Format          =   "00"
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
               ForeColor       =   &H000000C0&
               Height          =   315
               Index           =   4
               Left            =   4170
               TabIndex        =   112
               Top             =   480
               Width           =   1200
            End
            Begin MSComCtl2.DTPicker DTPicker1 
               Height          =   315
               Left            =   1335
               TabIndex        =   105
               Top             =   900
               Width           =   1695
               _ExtentX        =   2990
               _ExtentY        =   556
               _Version        =   393216
               CalendarTitleBackColor=   -2147483646
               CalendarTitleForeColor=   -2147483643
               Format          =   84803585
               CurrentDate     =   37417
            End
            Begin MSDataListLib.DataCombo Dat 
               Height          =   315
               Index           =   9
               Left            =   90
               TabIndex        =   110
               Top             =   480
               Width           =   2175
               _ExtentX        =   3836
               _ExtentY        =   556
               _Version        =   393216
               MatchEntry      =   -1  'True
               ListField       =   "NumCuenta"
               BoundColumn     =   "NombreBanco"
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
            Begin MSDataListLib.DataCombo Dat 
               Height          =   315
               Index           =   10
               Left            =   2280
               TabIndex        =   113
               Top             =   480
               Width           =   1875
               _ExtentX        =   3307
               _ExtentY        =   556
               _Version        =   393216
               MatchEntry      =   -1  'True
               ListField       =   "NombreBanco"
               BoundColumn     =   "NumCuenta"
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
            Begin MSFlexGridLib.MSFlexGrid GridCheques 
               Height          =   3015
               Index           =   1
               Left            =   90
               TabIndex        =   115
               Tag             =   "900|1400|1200|1900"
               Top             =   1500
               Width           =   5295
               _ExtentX        =   9340
               _ExtentY        =   5318
               _Version        =   393216
               Rows            =   3
               Cols            =   4
               FixedCols       =   0
               BackColorFixed  =   -2147483646
               ForeColorFixed  =   -2147483643
               BackColorSel    =   65280
               ForeColorSel    =   -2147483640
               HighLight       =   2
               MergeCells      =   2
               FormatString    =   "Cheque Nº |Banco |Monto|<Cód. Cuenta"
            End
            Begin VB.Label Label16 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Total Depósito:"
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
               Index           =   32
               Left            =   1530
               TabIndex        =   126
               Top             =   4800
               Width           =   1635
            End
            Begin VB.Label Label16 
               Alignment       =   1  'Right Justify
               BackColor       =   &H80000005&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "0,00"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Index           =   31
               Left            =   3225
               TabIndex        =   125
               Top             =   4755
               Width           =   1650
            End
            Begin VB.Label Label16 
               Alignment       =   1  'Right Justify
               BackColor       =   &H80000005&
               BorderStyle     =   1  'Fixed Single
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
               Index           =   30
               Left            =   3945
               TabIndex        =   124
               Top             =   878
               Width           =   1425
            End
            Begin VB.Label Label16 
               Alignment       =   2  'Center
               Caption         =   "Cuenta:"
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
               Index           =   21
               Left            =   210
               TabIndex        =   111
               Top             =   270
               Width           =   1575
            End
            Begin VB.Label Label16 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Efectivo:"
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
               Index           =   23
               Left            =   3135
               TabIndex        =   109
               Top             =   930
               Width           =   795
            End
            Begin VB.Label Label16 
               BackStyle       =   0  'Transparent
               Caption         =   "Fecha:"
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
               Index           =   20
               Left            =   720
               TabIndex        =   108
               Top             =   930
               Width           =   675
            End
            Begin VB.Label Label16 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "Banco:"
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
               Index           =   22
               Left            =   1830
               TabIndex        =   107
               Top             =   270
               Width           =   2010
            End
            Begin VB.Label Label16 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "Deposito Nº:"
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
               Index           =   19
               Left            =   4200
               TabIndex        =   106
               Top             =   270
               Width           =   1200
            End
         End
         Begin VB.Frame frame3 
            Caption         =   "Efectivo - Cheques Recibidos:"
            Height          =   4620
            Index           =   6
            Left            =   225
            TabIndex        =   103
            Top             =   1725
            Width           =   4620
            Begin MSFlexGridLib.MSFlexGrid GridCheques 
               Height          =   3735
               Index           =   0
               Left            =   195
               TabIndex        =   114
               Tag             =   "1000|1500|1400"
               Top             =   375
               Width           =   4215
               _ExtentX        =   7435
               _ExtentY        =   6588
               _Version        =   393216
               Rows            =   3
               Cols            =   3
               FixedCols       =   0
               BackColorFixed  =   -2147483646
               ForeColorFixed  =   -2147483643
               BackColorSel    =   65280
               ForeColorSel    =   -2147483640
               FocusRect       =   0
               HighLight       =   2
               SelectionMode   =   1
               FormatString    =   "Cheque Nº |Banco |Monto"
            End
            Begin VB.Label Label16 
               BackStyle       =   0  'Transparent
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000002&
               Height          =   210
               Index           =   26
               Left            =   1290
               TabIndex        =   119
               Top             =   4245
               Width           =   480
            End
            Begin VB.Label Label16 
               BackStyle       =   0  'Transparent
               Caption         =   "Cheq. Caja:"
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
               Index           =   24
               Left            =   210
               TabIndex        =   118
               Top             =   4245
               Width           =   1110
            End
            Begin VB.Label Label16 
               BackStyle       =   0  'Transparent
               Caption         =   "0"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000002&
               Height          =   210
               Index           =   27
               Left            =   3840
               TabIndex        =   117
               Top             =   4245
               Width           =   480
            End
            Begin VB.Label Label16 
               BackStyle       =   0  'Transparent
               Caption         =   "Total Cheq. Recibidos:"
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
               Index           =   25
               Left            =   1935
               TabIndex        =   116
               Top             =   4245
               Width           =   1905
            End
         End
         Begin VB.Frame frame3 
            Caption         =   "Seleccion Caja"
            Height          =   1275
            Index           =   5
            Left            =   225
            TabIndex        =   22
            Top             =   270
            Width           =   4635
            Begin MSDataListLib.DataCombo Dat 
               Height          =   315
               Index           =   7
               Left            =   825
               TabIndex        =   23
               Top             =   375
               Width           =   915
               _ExtentX        =   1614
               _ExtentY        =   556
               _Version        =   393216
               MatchEntry      =   -1  'True
               Style           =   2
               ListField       =   "Codigo"
               BoundColumn     =   "Descripcion"
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
            Begin MSDataListLib.DataCombo Dat 
               Height          =   315
               Index           =   8
               Left            =   1740
               TabIndex        =   24
               Top             =   375
               Width           =   2685
               _ExtentX        =   4736
               _ExtentY        =   556
               _Version        =   393216
               MatchEntry      =   -1  'True
               Style           =   2
               ListField       =   "Descripcion"
               BoundColumn     =   "Codigo"
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
            Begin VB.Label Label16 
               Alignment       =   1  'Right Justify
               BackColor       =   &H80000005&
               BorderStyle     =   1  'Fixed Single
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
               Index           =   29
               Left            =   2940
               TabIndex        =   123
               Top             =   788
               Width           =   1470
            End
            Begin VB.Label Label16 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Efectivo:"
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
               Index           =   28
               Left            =   2115
               TabIndex        =   120
               Top             =   840
               Width           =   795
            End
            Begin VB.Label Label16 
               BackStyle       =   0  'Transparent
               Caption         =   "Caja:"
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
               Index           =   18
               Left            =   255
               TabIndex        =   25
               Top             =   427
               Width           =   480
            End
         End
      End
      Begin MSAdodcLib.Adodc AdoDeducciones 
         Height          =   330
         Left            =   -67455
         Top             =   6210
         Visible         =   0   'False
         Width           =   2820
         _ExtentX        =   4974
         _ExtentY        =   582
         ConnectMode     =   3
         CursorLocation  =   3
         IsolationLevel  =   -1
         ConnectionTimeout=   15
         CommandTimeout  =   30
         CursorType      =   2
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
         Caption         =   "Adodc1"
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
         Bindings        =   "FrmMovCajaBs.frx":0738
         Height          =   3855
         Left            =   -74910
         TabIndex        =   48
         Top             =   480
         Width           =   10830
         _ExtentX        =   19103
         _ExtentY        =   6800
         _Version        =   393216
         AllowUpdate     =   0   'False
         AllowArrows     =   -1  'True
         Appearance      =   0
         BackColor       =   16777215
         BorderStyle     =   0
         ForeColor       =   -2147483630
         HeadLines       =   1
         RowHeight       =   14
         TabAcrossSplits =   -1  'True
         TabAction       =   2
         WrapCellPointer =   -1  'True
         RowDividerStyle =   6
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
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "LISTIN DE OPERACIONES DE CAJA"
         ColumnCount     =   7
         BeginProperty Column00 
            DataField       =   "NumeroMovimientoCaja"
            Caption         =   "Op."
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   "00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   8202
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   "InmuebleMovimientoCaja"
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
            DataField       =   "Nombre"
            Caption         =   "Nombre Inmueble"
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
            DataField       =   "AptoMovimientoCaja"
            Caption         =   "Apto."
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
            DataField       =   "Caja"
            Caption         =   "Caja Inm."
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
            DataField       =   "FormaPagoMovimientoCaja"
            Caption         =   "Pago"
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
            DataField       =   "MontoMovimientoCaja"
            Caption         =   "Monto"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   "#,##0.00  "
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
               Alignment       =   2
               ColumnWidth     =   585,071
            EndProperty
            BeginProperty Column01 
               Alignment       =   2
               ColumnWidth     =   734,74
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   3000,189
            EndProperty
            BeginProperty Column03 
               Alignment       =   2
               ColumnWidth     =   1305,071
            EndProperty
            BeginProperty Column04 
               Alignment       =   2
               ColumnWidth     =   824,882
            EndProperty
            BeginProperty Column05 
            EndProperty
            BeginProperty Column06 
               Alignment       =   1
               ColumnWidth     =   1590,236
            EndProperty
         EndProperty
      End
      Begin VB.Frame FmeCuentas 
         Enabled         =   0   'False
         Height          =   5970
         Left            =   75
         TabIndex        =   62
         Top             =   360
         Width           =   10815
         Begin VB.CheckBox CHK 
            Alignment       =   1  'Right Justify
            Caption         =   "En&viar al Edificio:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   8670
            TabIndex        =   146
            Top             =   5055
            Width           =   1950
         End
         Begin VB.Frame frame3 
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   360
            Index           =   8
            Left            =   6840
            TabIndex        =   129
            Top             =   300
            Width           =   3870
            Begin VB.CommandButton Command 
               DisabledPicture =   "FrmMovCajaBs.frx":0754
               Enabled         =   0   'False
               Height          =   255
               Index           =   0
               Left            =   3510
               Picture         =   "FrmMovCajaBs.frx":089E
               Style           =   1  'Graphical
               TabIndex        =   131
               Top             =   45
               Width           =   255
            End
            Begin VB.TextBox Txt 
               BackColor       =   &H80000002&
               DataField       =   "NumeroMovimientoCaja"
               BeginProperty DataFormat 
                  Type            =   1
                  Format          =   "00"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   8202
                  SubFormatType   =   0
               EndProperty
               DataSource      =   "ADOcontrol(0)"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000009&
               Height          =   285
               Index           =   0
               Left            =   660
               Locked          =   -1  'True
               TabIndex        =   130
               Text            =   " "
               Top             =   15
               Width           =   840
            End
            Begin MSMask.MaskEdBox MskFecha 
               Bindings        =   "FrmMovCajaBs.frx":09E8
               DataField       =   "FechaMovimientoCaja"
               BeginProperty DataFormat 
                  Type            =   1
                  Format          =   "dd/MM/yyyy"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   8202
                  SubFormatType   =   3
               EndProperty
               DataSource      =   "ADOcontrol(0)"
               Height          =   315
               Index           =   0
               Left            =   2370
               TabIndex        =   132
               TabStop         =   0   'False
               Tag             =   "0"
               Top             =   15
               Width           =   1425
               _ExtentX        =   2514
               _ExtentY        =   556
               _Version        =   393216
               BackColor       =   -2147483646
               ForeColor       =   -2147483639
               PromptInclude   =   0   'False
               MaxLength       =   10
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Format          =   "dd/MM/yyyy"
               Mask            =   "##/##/####"
               PromptChar      =   "_"
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Op.N.:"
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
               Left            =   105
               TabIndex        =   134
               Top             =   45
               Width           =   540
            End
            Begin VB.Label Lbl 
               AutoSize        =   -1  'True
               Caption         =   "&Fecha :"
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
               Left            =   1635
               TabIndex        =   133
               Top             =   30
               Width           =   600
            End
         End
         Begin VB.Frame Frame5 
            BorderStyle     =   0  'None
            Height          =   1530
            Left            =   90
            TabIndex        =   68
            Top             =   3510
            Visible         =   0   'False
            Width           =   7995
            Begin VB.CommandButton Command 
               Caption         =   ".."
               Height          =   285
               Index           =   6
               Left            =   7725
               Style           =   1  'Graphical
               TabIndex        =   151
               Top             =   1065
               Width           =   195
            End
            Begin VB.CommandButton Command 
               Caption         =   ".."
               Height          =   285
               Index           =   5
               Left            =   7725
               Style           =   1  'Graphical
               TabIndex        =   150
               Top             =   780
               Width           =   195
            End
            Begin VB.CommandButton Command 
               Caption         =   ".."
               Height          =   285
               Index           =   4
               Left            =   7725
               Style           =   1  'Graphical
               TabIndex        =   149
               Top             =   495
               Width           =   195
            End
            Begin VB.TextBox Txt 
               Alignment       =   1  'Right Justify
               DataField       =   "NumDocumentoMovimientoCaja"
               DataSource      =   "ADOcontrol(0)"
               Height          =   315
               Index           =   6
               Left            =   1395
               MaxLength       =   11
               TabIndex        =   83
               ToolTipText     =   "Nro.Documento 1"
               Top             =   480
               Width           =   1785
            End
            Begin VB.CommandButton Command 
               Height          =   255
               Index           =   1
               Left            =   6075
               Picture         =   "FrmMovCajaBs.frx":0A0A
               Style           =   1  'Graphical
               TabIndex        =   82
               Top             =   510
               Width           =   255
            End
            Begin VB.ComboBox Cmb 
               DataField       =   "BancoDocumentoMovimientoCaja"
               DataSource      =   "ADOcontrol(0)"
               Height          =   315
               Index           =   2
               ItemData        =   "FrmMovCajaBs.frx":0B54
               Left            =   3255
               List            =   "FrmMovCajaBs.frx":0B56
               Sorted          =   -1  'True
               TabIndex        =   81
               ToolTipText     =   "Documento Banco 1"
               Top             =   480
               Width           =   1785
            End
            Begin VB.CommandButton Command 
               Height          =   255
               Index           =   2
               Left            =   6075
               Picture         =   "FrmMovCajaBs.frx":0B58
               Style           =   1  'Graphical
               TabIndex        =   80
               Top             =   825
               Width           =   255
            End
            Begin VB.TextBox Txt 
               Alignment       =   1  'Right Justify
               DataField       =   "NumDocumentoMovimientoCaja1"
               DataSource      =   "ADOcontrol(0)"
               Height          =   315
               Index           =   7
               Left            =   1410
               MaxLength       =   11
               TabIndex        =   79
               Text            =   " "
               ToolTipText     =   "Nro.Documento 2"
               Top             =   795
               Width           =   1785
            End
            Begin VB.ComboBox Cmb 
               CausesValidation=   0   'False
               DataField       =   "BancoDocumentoMovimientoCaja1"
               DataSource      =   "ADOcontrol(0)"
               Height          =   315
               Index           =   3
               ItemData        =   "FrmMovCajaBs.frx":0CA2
               Left            =   3255
               List            =   "FrmMovCajaBs.frx":0CA4
               Sorted          =   -1  'True
               TabIndex        =   78
               ToolTipText     =   "Documento Banco 2"
               Top             =   795
               Width           =   1785
            End
            Begin VB.TextBox Txt 
               Alignment       =   1  'Right Justify
               DataField       =   "MontoCheque"
               BeginProperty DataFormat 
                  Type            =   1
                  Format          =   "#,##0.00"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   8202
                  SubFormatType   =   0
               EndProperty
               DataSource      =   "ADOcontrol(0)"
               Height          =   315
               Index           =   1
               Left            =   6390
               TabIndex        =   77
               ToolTipText     =   "Doc. Monto 1"
               Top             =   480
               Width           =   1335
            End
            Begin VB.TextBox Txt 
               Alignment       =   1  'Right Justify
               DataField       =   "MontoCheque1"
               BeginProperty DataFormat 
                  Type            =   1
                  Format          =   "#,##0.00"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   8202
                  SubFormatType   =   0
               EndProperty
               DataSource      =   "ADOcontrol(0)"
               Height          =   315
               Index           =   2
               Left            =   6390
               TabIndex        =   76
               Text            =   " "
               ToolTipText     =   "Doc.Monto 2"
               Top             =   795
               Width           =   1335
            End
            Begin VB.TextBox Txt 
               Alignment       =   1  'Right Justify
               DataField       =   "MontoCheque2"
               BeginProperty DataFormat 
                  Type            =   1
                  Format          =   "#,##0.00"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   8202
                  SubFormatType   =   0
               EndProperty
               DataSource      =   "ADOcontrol(0)"
               Height          =   315
               Index           =   3
               Left            =   6390
               TabIndex        =   75
               Text            =   " "
               ToolTipText     =   "Doc.Monto 3"
               Top             =   1110
               Width           =   1335
            End
            Begin VB.ComboBox Cmb 
               CausesValidation=   0   'False
               DataField       =   "BancoDocumentoMovimientoCaja2"
               DataSource      =   "ADOcontrol(0)"
               Height          =   315
               Index           =   4
               ItemData        =   "FrmMovCajaBs.frx":0CA6
               Left            =   3255
               List            =   "FrmMovCajaBs.frx":0CA8
               Sorted          =   -1  'True
               TabIndex        =   74
               ToolTipText     =   "Documento Banco 3"
               Top             =   1110
               Width           =   1785
            End
            Begin VB.TextBox Txt 
               Alignment       =   1  'Right Justify
               DataField       =   "NumDocumentoMovimientoCaja2"
               DataSource      =   "ADOcontrol(0)"
               Height          =   315
               Index           =   8
               Left            =   1410
               MaxLength       =   11
               TabIndex        =   73
               Text            =   " "
               ToolTipText     =   "Nro.Documento 3"
               Top             =   1110
               Width           =   1785
            End
            Begin VB.CommandButton Command 
               Height          =   255
               Index           =   3
               Left            =   6075
               Picture         =   "FrmMovCajaBs.frx":0CAA
               Style           =   1  'Graphical
               TabIndex        =   72
               Top             =   1140
               Width           =   255
            End
            Begin VB.ComboBox Cmb 
               DataField       =   "FPago"
               DataSource      =   "ADOcontrol(0)"
               Height          =   315
               Index           =   5
               ItemData        =   "FrmMovCajaBs.frx":0DF4
               Left            =   30
               List            =   "FrmMovCajaBs.frx":0E01
               Sorted          =   -1  'True
               TabIndex        =   71
               ToolTipText     =   "Forma de Pago 1"
               Top             =   480
               Width           =   1305
            End
            Begin VB.ComboBox Cmb 
               DataField       =   "FPago1"
               DataSource      =   "ADOcontrol(0)"
               Height          =   315
               Index           =   6
               ItemData        =   "FrmMovCajaBs.frx":0E26
               Left            =   30
               List            =   "FrmMovCajaBs.frx":0E33
               Sorted          =   -1  'True
               TabIndex        =   70
               ToolTipText     =   "Forma de Pago 2"
               Top             =   795
               Width           =   1305
            End
            Begin VB.ComboBox Cmb 
               DataField       =   "FPago2"
               DataSource      =   "ADOcontrol(0)"
               Height          =   315
               Index           =   7
               ItemData        =   "FrmMovCajaBs.frx":0E58
               Left            =   30
               List            =   "FrmMovCajaBs.frx":0E65
               Sorted          =   -1  'True
               TabIndex        =   69
               ToolTipText     =   "Forma de Pago 3"
               Top             =   1110
               Width           =   1305
            End
            Begin MSMask.MaskEdBox MskFecha 
               Bindings        =   "FrmMovCajaBs.frx":0E8A
               DataField       =   "FechaChequeMovimientoCaja1"
               BeginProperty DataFormat 
                  Type            =   1
                  Format          =   "dd/MM/yyyy"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   8202
                  SubFormatType   =   3
               EndProperty
               DataSource      =   "ADOcontrol(0)"
               Height          =   315
               Index           =   2
               Left            =   5070
               TabIndex        =   84
               TabStop         =   0   'False
               ToolTipText     =   "Documento Fecha 2"
               Top             =   795
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
            Begin MSMask.MaskEdBox MskFecha 
               Bindings        =   "FrmMovCajaBs.frx":0EAC
               DataField       =   "FechaChequeMovimientoCaja2"
               BeginProperty DataFormat 
                  Type            =   1
                  Format          =   "dd/MM/yyyy"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   8202
                  SubFormatType   =   3
               EndProperty
               DataSource      =   "ADOcontrol(0)"
               Height          =   315
               Index           =   3
               Left            =   5070
               TabIndex        =   85
               TabStop         =   0   'False
               ToolTipText     =   "Documento Fecha 3"
               Top             =   1110
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
            Begin MSMask.MaskEdBox MskFecha 
               Bindings        =   "FrmMovCajaBs.frx":0ECE
               DataField       =   "FechaChequeMovimientoCaja"
               BeginProperty DataFormat 
                  Type            =   1
                  Format          =   "dd/MM/yyyy"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   8202
                  SubFormatType   =   3
               EndProperty
               DataSource      =   "ADOcontrol(0)"
               Height          =   315
               Index           =   1
               Left            =   5070
               TabIndex        =   86
               TabStop         =   0   'False
               ToolTipText     =   "Documento Fecha 1"
               Top             =   480
               Width           =   1290
               _ExtentX        =   2275
               _ExtentY        =   556
               _Version        =   393216
               PromptInclude   =   0   'False
               MaxLength       =   10
               Format          =   "dd/mm/yyyy"
               Mask            =   "##/##/####"
               PromptChar      =   "_"
            End
            Begin VB.Label Lbl 
               Alignment       =   2  'Center
               BackColor       =   &H80000002&
               Caption         =   "Documento"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000009&
               Height          =   285
               Index           =   11
               Left            =   1380
               TabIndex        =   91
               Top             =   180
               Width           =   1830
            End
            Begin VB.Label Lbl 
               Alignment       =   2  'Center
               BackColor       =   &H80000002&
               Caption         =   "Banco"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000009&
               Height          =   285
               Index           =   12
               Left            =   3210
               TabIndex        =   90
               Top             =   180
               Width           =   1815
            End
            Begin VB.Label Lbl 
               Alignment       =   2  'Center
               BackColor       =   &H80000002&
               Caption         =   "Fecha"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000009&
               Height          =   285
               Index           =   13
               Left            =   5025
               TabIndex        =   89
               Top             =   180
               Width           =   1335
            End
            Begin VB.Label Lbl 
               Alignment       =   2  'Center
               BackColor       =   &H80000002&
               Caption         =   "Monto"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000009&
               Height          =   285
               Index           =   4
               Left            =   6360
               TabIndex        =   88
               Top             =   180
               Width           =   1545
            End
            Begin VB.Label Lbl 
               Alignment       =   2  'Center
               BackColor       =   &H80000002&
               Caption         =   "Forma de Pago"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000009&
               Height          =   285
               Index           =   5
               Left            =   30
               TabIndex        =   87
               Top             =   180
               Width           =   1365
            End
         End
         Begin VB.CommandButton Command1 
            Caption         =   "&Aplicar"
            Height          =   540
            Index           =   0
            Left            =   8040
            TabIndex        =   10
            ToolTipText     =   "Presione para seleccionar las facturas a cancelar"
            Top             =   3045
            Width           =   1200
         End
         Begin VB.CommandButton Command1 
            Caption         =   "&Distribuir"
            Height          =   540
            Index           =   1
            Left            =   9470
            TabIndex        =   11
            ToolTipText     =   "Presione para distribuir el monto a candelar desde la factura mas antigua"
            Top             =   3030
            Width           =   1200
         End
         Begin VB.Frame frame3 
            Enabled         =   0   'False
            Height          =   1305
            Index           =   0
            Left            =   7995
            TabIndex        =   63
            Top             =   1590
            Width           =   2685
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
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000002&
               Height          =   360
               Index           =   11
               Left            =   1020
               TabIndex        =   65
               Top             =   780
               Width           =   1455
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
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   360
               Index           =   10
               Left            =   1020
               TabIndex        =   64
               Text            =   " "
               Top             =   240
               Width           =   1455
            End
            Begin VB.Label Lbl 
               Alignment       =   1  'Right Justify
               Caption         =   "Deuda + Honorarios:"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   405
               Index           =   18
               Left            =   75
               TabIndex        =   67
               Top             =   735
               Width           =   915
            End
            Begin VB.Label Lbl 
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
               Index           =   17
               Left            =   165
               TabIndex        =   66
               Top             =   270
               Width           =   765
            End
         End
         Begin VB.TextBox Txt 
            Alignment       =   1  'Right Justify
            DataField       =   "EfectivoMovimientoCaja"
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "#,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   8202
               SubFormatType   =   0
            EndProperty
            DataSource      =   "ADOcontrol(0)"
            Height          =   285
            Index           =   12
            Left            =   9180
            TabIndex        =   13
            Text            =   "0"
            Top             =   3750
            Visible         =   0   'False
            Width           =   1455
         End
         Begin VB.ComboBox Cmb 
            DataField       =   "FormaPagoMovimientoCaja"
            DataSource      =   "ADOcontrol(0)"
            Height          =   315
            Index           =   0
            ItemData        =   "FrmMovCajaBs.frx":0EF0
            Left            =   9180
            List            =   "FrmMovCajaBs.frx":0F03
            Sorted          =   -1  'True
            TabIndex        =   9
            ToolTipText     =   "Forma de Pago"
            Top             =   1185
            Width           =   1455
         End
         Begin VB.ComboBox Cmb 
            DataField       =   "TipoMovimientoCaja"
            DataSource      =   "ADOcontrol(0)"
            Height          =   315
            Index           =   1
            ItemData        =   "FrmMovCajaBs.frx":0F39
            Left            =   9180
            List            =   "FrmMovCajaBs.frx":0F43
            Sorted          =   -1  'True
            TabIndex        =   7
            ToolTipText     =   "Tipo de Movimiento"
            Top             =   780
            Width           =   1455
         End
         Begin VB.TextBox Txt 
            Alignment       =   1  'Right Justify
            DataField       =   "MontoMovimientoCaja"
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "#,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   8202
               SubFormatType   =   0
            EndProperty
            DataSource      =   "ADOcontrol(0)"
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
            Index           =   5
            Left            =   9180
            TabIndex        =   15
            Text            =   "0,00"
            Top             =   4155
            Width           =   1455
         End
         Begin VB.TextBox Txt 
            DataField       =   "DescripcionMovimientoCaja"
            DataSource      =   "ADOcontrol(0)"
            Height          =   315
            Index           =   9
            Left            =   1515
            MaxLength       =   254
            TabIndex        =   94
            Text            =   " "
            ToolTipText     =   "Concepto de Caja"
            Top             =   5500
            Width           =   9120
         End
         Begin MSDataListLib.DataCombo Dat 
            Bindings        =   "FrmMovCajaBs.frx":0F58
            DataField       =   "CodGasto"
            DataSource      =   "ADOcontrol(0)"
            Height          =   315
            Index           =   5
            Left            =   1530
            TabIndex        =   92
            ToolTipText     =   "Cod. de Cuenta"
            Top             =   5100
            Width           =   1050
            _ExtentX        =   1852
            _ExtentY        =   556
            _Version        =   393216
            ListField       =   "CodGasto"
            BoundColumn     =   "Cuenta de Ingreso / Egreso"
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
         Begin MSDataListLib.DataCombo Dat 
            Bindings        =   "FrmMovCajaBs.frx":0F74
            DataField       =   "CuentaMovimientoCaja"
            DataSource      =   "ADOcontrol(0)"
            Height          =   315
            Index           =   6
            Left            =   2595
            TabIndex        =   93
            ToolTipText     =   "Descripción de Cuenta"
            Top             =   5100
            Width           =   5295
            _ExtentX        =   9340
            _ExtentY        =   556
            _Version        =   393216
            ListField       =   "Titulo"
            BoundColumn     =   "Descripcion"
            Text            =   " "
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
         Begin VB.Frame Frame6 
            BorderStyle     =   0  'None
            Caption         =   "Frame6"
            Height          =   1380
            Left            =   270
            TabIndex        =   142
            Top             =   135
            Width           =   6585
            Begin MSDataListLib.DataCombo Dat 
               Bindings        =   "FrmMovCajaBs.frx":0F90
               DataField       =   "AptoMovimientoCaja"
               DataSource      =   "ADOcontrol(0)"
               Height          =   315
               Index           =   2
               Left            =   1530
               TabIndex        =   5
               ToolTipText     =   "Codigo del Propietario"
               Top             =   600
               Width           =   1050
               _ExtentX        =   1852
               _ExtentY        =   556
               _Version        =   393216
               ListField       =   "Codigo"
               BoundColumn     =   "Codigo"
               Text            =   " "
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
            Begin MSDataListLib.DataCombo Dat 
               Bindings        =   "FrmMovCajaBs.frx":0FAC
               Height          =   315
               Index           =   4
               Left            =   2580
               TabIndex        =   143
               ToolTipText     =   "Descripción de Caja"
               Top             =   1050
               Width           =   3915
               _ExtentX        =   6906
               _ExtentY        =   556
               _Version        =   393216
               Locked          =   -1  'True
               ListField       =   "DescripCaja"
               BoundColumn     =   "DescripCaja"
               Text            =   " "
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
            Begin MSDataListLib.DataCombo Dat 
               Bindings        =   "FrmMovCajaBs.frx":0FC8
               Height          =   315
               Index           =   3
               Left            =   1515
               TabIndex        =   144
               ToolTipText     =   "Codigo de Caja"
               Top             =   1050
               Width           =   1050
               _ExtentX        =   1852
               _ExtentY        =   556
               _Version        =   393216
               Locked          =   -1  'True
               ListField       =   "Caja"
               BoundColumn     =   "CodigoCaja"
               Text            =   " "
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
            Begin MSDataListLib.DataCombo Dat 
               DataField       =   "InmuebleMovimientoCaja"
               DataSource      =   "ADOcontrol(0)"
               Height          =   315
               Index           =   0
               Left            =   1530
               TabIndex        =   1
               ToolTipText     =   "Codigo del Inmueble"
               Top             =   165
               Width           =   1050
               _ExtentX        =   1852
               _ExtentY        =   556
               _Version        =   393216
               BackColor       =   -2147483643
               ListField       =   "CodInm"
               BoundColumn     =   "CodInm"
               Text            =   " "
               Object.DataMember      =   ""
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
            Begin MSDataListLib.DataCombo Dat 
               Height          =   315
               Index           =   1
               Left            =   2595
               TabIndex        =   3
               ToolTipText     =   "Nombre del Inmueble"
               Top             =   165
               Width           =   3915
               _ExtentX        =   6906
               _ExtentY        =   556
               _Version        =   393216
               ListField       =   "Nombre"
               BoundColumn     =   "Nombre"
               Text            =   " "
               Object.DataMember      =   ""
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
            Begin MSDataListLib.DataCombo DatProp 
               Bindings        =   "FrmMovCajaBs.frx":0FE3
               Height          =   315
               Left            =   2595
               TabIndex        =   6
               ToolTipText     =   "Nombre del Propietario"
               Top             =   600
               Width           =   3915
               _ExtentX        =   6906
               _ExtentY        =   556
               _Version        =   393216
               ListField       =   "Nombre"
               BoundColumn     =   "Nombre"
               Text            =   " "
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
            Begin VB.Label Lbl 
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
               Left            =   45
               TabIndex        =   145
               Top             =   1110
               Width           =   1470
            End
            Begin VB.Label Lbl 
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
               Index           =   3
               Left            =   45
               TabIndex        =   4
               Top             =   660
               Width           =   1470
            End
            Begin VB.Label Lbl 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
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
               Width           =   1470
            End
         End
         Begin VB.Image ImgAceptar 
            Enabled         =   0   'False
            Height          =   480
            Index           =   0
            Left            =   900
            Picture         =   "FrmMovCajaBs.frx":0FFF
            Top             =   870
            Visible         =   0   'False
            Width           =   480
         End
         Begin VB.Label Label16 
            Alignment       =   1  'Right Justify
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0,00"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Index           =   17
            Left            =   9180
            TabIndex        =   102
            Top             =   5040
            Visible         =   0   'False
            Width           =   1455
         End
         Begin VB.Label Label16 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Deducciones:"
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
            Index           =   16
            Left            =   8055
            TabIndex        =   101
            Top             =   5040
            Visible         =   0   'False
            Width           =   1095
         End
         Begin VB.Label LblHono 
            Alignment       =   1  'Right Justify
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0,00"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   9180
            TabIndex        =   100
            Top             =   4635
            Width           =   1455
         End
         Begin VB.Label Lbl 
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
            Left            =   8160
            TabIndex        =   99
            Top             =   4665
            Width           =   975
         End
         Begin VB.Label Lbl 
            AutoSize        =   -1  'True
            Caption         =   "Bs. &Efectivo :"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   8
            Left            =   8145
            TabIndex        =   12
            Top             =   3780
            Visible         =   0   'False
            Width           =   960
         End
         Begin VB.Label Lbl 
            AutoSize        =   -1  'True
            Caption         =   "Cta.Ingr/E&gr :"
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
            Index           =   15
            Left            =   360
            TabIndex        =   98
            Top             =   5115
            Width           =   1125
         End
         Begin VB.Label Lbl 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "&Monto :"
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
            Left            =   8490
            TabIndex        =   14
            Top             =   4245
            Width           =   645
         End
         Begin VB.Label Lbl 
            AutoSize        =   -1  'True
            Caption         =   "&Forma Pago :"
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
            Left            =   7965
            TabIndex        =   8
            Top             =   1215
            Width           =   1080
         End
         Begin VB.Label Lbl 
            AutoSize        =   -1  'True
            Caption         =   "&Tipo Movimiento :"
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
            Index           =   10
            Left            =   7590
            TabIndex        =   2
            Top             =   810
            Width           =   1470
         End
         Begin VB.Label Lbl 
            AutoSize        =   -1  'True
            Caption         =   "De&scripción :"
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
            Index           =   16
            Left            =   435
            TabIndex        =   95
            Top             =   5505
            Width           =   1035
         End
      End
      Begin MSWinsockLib.Winsock Winsock1 
         Left            =   -71550
         Top             =   5685
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
         RemotePort      =   888
      End
      Begin MSDataListLib.DataCombo dtc 
         Bindings        =   "FrmMovCajaBs.frx":1441
         Height          =   315
         Left            =   -71520
         TabIndex        =   148
         Top             =   5790
         Width           =   2520
         _ExtentX        =   4445
         _ExtentY        =   556
         _Version        =   393216
         ListField       =   "NombreUsuario"
         Text            =   "--Selecciona un usuario--"
      End
      Begin MSAdodcLib.Adodc ADOcontrol 
         Height          =   330
         Index           =   4
         Left            =   -74520
         Top             =   6465
         Visible         =   0   'False
         Width           =   3405
         _ExtentX        =   6006
         _ExtentY        =   582
         ConnectMode     =   1
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
         Caption         =   "Supervisores Activos"
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
      Begin VB.Label Label16 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0,00"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   0
         Left            =   -65895
         TabIndex        =   61
         Top             =   5790
         Width           =   1635
      End
      Begin VB.Label Label16 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total Deducciones:"
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
         Left            =   -67530
         TabIndex        =   60
         Top             =   5850
         Width           =   1575
      End
   End
   Begin VB.TextBox Txt 
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
      Left            =   1725
      TabIndex        =   96
      Text            =   " "
      Top             =   7050
      Width           =   9360
   End
   Begin MSAdodcLib.Adodc ADOcontrol 
      Height          =   345
      Index           =   1
      Left            =   2250
      Top             =   6660
      Visible         =   0   'False
      Width           =   2700
      _ExtentX        =   4763
      _ExtentY        =   609
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
      Caption         =   "AdoCuenta"
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
   Begin MSAdodcLib.Adodc ADOcontrol 
      Height          =   330
      Index           =   3
      Left            =   7605
      Top             =   6585
      Visible         =   0   'False
      Width           =   2700
      _ExtentX        =   4763
      _ExtentY        =   582
      ConnectMode     =   3
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
      Caption         =   "AdoGrid"
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
   Begin MSAdodcLib.Adodc ADOcontrol 
      Height          =   330
      Index           =   2
      Left            =   2865
      Top             =   8070
      Visible         =   0   'False
      Width           =   2325
      _ExtentX        =   4101
      _ExtentY        =   582
      ConnectMode     =   1
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
      Caption         =   "AdoPropietario"
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
   Begin VB.Label Lbl 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Observ.Especificas:"
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
      TabIndex        =   19
      Top             =   7125
      Width           =   1605
   End
   Begin VB.Label Lbl 
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000002&
      Height          =   315
      Index           =   20
      Left            =   1725
      TabIndex        =   18
      Top             =   7485
      Width           =   9360
   End
   Begin VB.Label Lbl 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Observ.Generales:"
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
      Index           =   19
      Left            =   105
      TabIndex        =   17
      Top             =   7500
      Width           =   1605
   End
   Begin ComctlLib.ImageList ImageList1 
      Left            =   11340
      Top             =   675
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
            Picture         =   "FrmMovCajaBs.frx":145D
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmMovCajaBs.frx":15DF
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmMovCajaBs.frx":1761
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmMovCajaBs.frx":18E3
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmMovCajaBs.frx":1A65
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmMovCajaBs.frx":1BE7
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmMovCajaBs.frx":1D69
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmMovCajaBs.frx":1EEB
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmMovCajaBs.frx":206D
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmMovCajaBs.frx":21EF
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmMovCajaBs.frx":2371
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmMovCajaBs.frx":24F3
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "FrmMovCajaBs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
    '--------------------------------------------------------------------------------------------'
    '   SINATI TECH-SAC sistema de administracion de condominio
    '   Modulo de Caja, Recepción de pagos de Condominio, Ingresos y Egresos de
    '--------------------------------------------------------------------------------------------'
    'Variables públicas a nivel de módulo
    Public StrRutaInmueble As String
    Public strCodPC As String
    Public NCta As String
    Dim Endoso As New STEndoso
    Dim strUbica As String, strCodHA As String
    Dim strCodRebHA As String, strCodIV As String
    Dim strPeriodo As String, strQ As String
    Dim strRecibo As String, strCodCCHeq As String
    Dim strCodRCheq As String, strCodAbonoFut As String
    Dim strNdeposito As String, strCodAbonoCta As String
    Dim cnnPropietario As New ADODB.Connection  'Conexion Global a Nivel de Módulo
    Dim objRst As New ADODB.Recordset   'Recordset Global anivel de módulo
    Dim vecFPS(120, 3) As Currency, curAbono@
    Dim booAFuturo As Boolean, booACuenta As Boolean, booDed As Boolean
    Dim booIV As Boolean, booEV As Boolean, booHA As Boolean
    Dim curHono@, curSaldo@, curMD@, CurTotalCheque@, IntMonto@, curPagoHono@
    Dim mlEdit As Boolean, booPC As Boolean, booCon As Boolean
    Dim intRow%, B%, IntMesesMora%, I%, K%, IntHonoMorosidad%
    Dim rstlocal(1) As New ADODB.Recordset
    '---------------------------------------------------------------------------------------------
    Private Enum Indice
        FP = 0
        NDoc
        Banco
        FechaDoc
        Monto
        IDRecibo
        Inmueble
    End Enum
    '---------------------------------------------------------------------------------------------
    
    Private Sub cmb_Click(Index%): If Index = 0 Then Cmb(0).SetFocus
    End Sub

    Private Sub Cmb_KeyDown(Index%, KeyCode%, Shift%)
    If Index = 0 And KeyCode = 46 Then KeyCode = 0
    End Sub

    '---------------------------------------------------------------------------------------------
    Private Sub cmb_KeyPress(Index%, KeyAscii%)
    '---------------------------------------------------------------------------------------------
    'variables locales
    Dim Mensaje$
    '
    KeyAscii = Asc(UCase(Chr(KeyAscii)))    'convierte todo en mayúsculas
    If KeyAscii = 8 Then    'Tecla {backspace}
    '
        Select Case Index
    '
            Case 2, 3, 4
                If Cmb(Index) = "" Then Txt(2 + 4).SetFocus
            End Select
    '
    End If
     '
    If KeyAscii = 13 Then 'Tecla {enter}
        Frame5.Enabled = True
        Select Case Index
    '
            Case 0
            
                If Not inLista(Cmb(0)) Then Exit Sub
                If Txt(5) = "" Then MsgBox "No Existe Monto a Cancelar " & Chr(13) _
                    & "Verifique e intente nuevamente", vbInformation, "Error": Exit Sub
    '
                If Cmb(Index) = "" Then MsgBox "Debe Introducir un Valor, por favor....", _
                    vbCritical: Exit Sub
    '
    '---------------------------------------------------------------------------------------------
                If LblHono = "" Then LblHono = 0    'Si tiene honorarios
    '
                    If Cmb(1) = "INGRESO" And LblHono > 0 Then
                        curHono = CCur(LblHono)
                        'booHA = True
                        Mensaje = "Propietario tiene honorarios por " & Format(LblHono, "#,##0") _
                            & (Chr(13)) & "Monto total: Bs " _
                            & Format(CCur(Txt(5)) + CCur(LblHono), "#,##0.00") _
                            & (Chr(13)) & "¿Desea Aplicar Honorarios? "
                            
                       If Respuesta(Mensaje) = False Then 'MUESTRA LA FICHA DEDUCCIONES
                            FmeCuentas.Enabled = True
                            Call rtnTab(2)
                            Label16(12) = ""
                            'Label16(13) = " " & Format(Date, "mm-yy")
                            LblCodInm = " " & Dat(0)
                            Label16(3) = " " & Dat(1)
                            Label16(5) = " " & Dat(2)
                            Label16(6) = " " & DatProp
                            Label16(7) = " " & Dat(3)
                            Label16(8) = " " & Dat(4)
                            Label16(0) = "0,00"
                            Label16(13) = ""
                            With FlexDeducciones
                                .Rows = 2
                                .TextMatrix(1, 0) = strCodRebHA
                                .TextMatrix(1, 1) = "REBAJA DE HONORARIOS DE ABOGADO"
                                .TextMatrix(1, 2) = Format(CCur(LblHono), "#,##0.00")
                                .TextMatrix(1, 3) = CCur(LblHono)
                                .Col = 2
                            End With
                            Exit Sub
                       Else
                            Txt(5) = Format(CCur(Txt(5)) + CCur(LblHono), "#,##0.00")
                            LblHono = "0,00"
                            curPagoHono = curHono
                       End If
                       '
                    End If
                    '
                    Call Forma_Pago
    '               ----------------------------------
                Case 1
                    If Dat(0) = sysCodInm Or FlexFacturas.TextMatrix(1, 2) = "" Then
                        Txt(5).SetFocus
                    Else
                        Command1(0).Enabled = True: Command1(0).SetFocus
                    End If
                
                Case 2, 3, 4: Command(Index - 1).SetFocus
                                                        
                Case 5, 6, 7
                    Call Validacion(KeyAscii, "ABCDEFGHIJKLMNOPQRSTUVWXYZ.,")
                    Txt(Index + 1).SetFocus
            
            End Select
        '
    End If
    '
    End Sub
    
    
    Private Sub Cmb_GotFocus(Index As Integer)
        If Index <> 1 Then Exit Sub
        Dat(5) = ""
        Dat(6) = ""
    End Sub
    
    Private Sub Cmb_LostFocus(Index As Integer)
    '
    If Index >= 5 And Index <= 7 And (Cmb(Index) = "DEPOSITO" Or Cmb(Index) = "TRANSFERENCIA") Then
    
            frmCajaCta.strTitulo = "Cuentas Inm:" & Dat(0)
            Matriz_A = varCuentas
            If UBound(Matriz_A, 2) = 0 Then
                NCta = Matriz_A(0, 0)
                Cmb(Index - 3) = Matriz_A(2, 0)
            Else
                frmCajaCta.Show
                Unload frmCajaCta
                Cmb(Index - 3) = frmCajaCta.strTitulo
            End If
            Cmb(Index).Tag = NCta
            
            
    ElseIf Index = 0 And Not inLista(Cmb(0)) Then
        Exit Sub
    End If
    '
    End Sub

    Private Sub Command_KeyPress(Index%, KeyAscii%)
    If Index > 0 And KeyAscii = 8 Then Cmb(Index + 1).SetFocus
    End Sub


    'Rev.26/08/2002-------------------------------------------------------------------------------
    Private Sub Command1_Click(Index As Integer)
    '---------------------------------------------------------------------------------------------
    Cmb(0).Enabled = True
    Select Case Index
    '
        Case 0  'permite al objeto responder a los eventos generados por el usuario
    '   ------------------------------------------------------------------------------------------
        Frame1.Enabled = True
        Command1(1).Enabled = False
        FlexFacturas.Row = 1
        FlexFacturas.Col = 3
        FlexFacturas.SetFocus
        
        Case 1  'distribuye automaticamente el monto a cancelar a partir de la deuda mas antigua
    '   ------------------------------------------------------------------------------------------
        If Command1(1).Caption = "&Deshacer" Then
            If Not Respuesta("¿Está seguro que desea Deshacer? ") Then Exit Sub
            mnrow = 1
            mnRows = FlexFacturas.Rows - 1
            Do While FlexFacturas.CellPicture = ImgAceptar(0).Picture And mnrow <= mnRows
                FlexFacturas.TextMatrix(mnrow, 3) = Format(vecFPS(mnrow, 1), "#,##0.00")
                FlexFacturas.TextMatrix(mnrow, 4) = Format(vecFPS(mnrow, 2), "#,##0.00")
                FlexFacturas.TextMatrix(mnrow, 6) = "NO"
                FlexFacturas.Row = mnrow: FlexFacturas.Col = 6
                'Txt(10) = Format(CCur(Txt(10)) + CCur(vecFPS(mnrow, 2)), "#,##0.00")
                Set FlexFacturas.CellPicture = Nothing
                
                mnrow = IIf(mnrow + 1 <= FlexFacturas.Rows - 1, mnrow + 1, nmrow)
                FlexFacturas.Row = mnrow: FlexFacturas.Col = 6
                
            Loop
            Txt(10) = Format(CCur(Txt(10)) + CCur(Txt(5)), "#,##0.00")
            Txt(11) = Format(CCur(Txt(10)) + CCur(LblHono), "#,##0.00")
            Command1(1).Caption = "&Distribuir": Frame1.Enabled = False
            Command1(0).Enabled = True: Txt(5) = 0: Txt(9) = ""
        Else
            
            If (CCur(Txt(5)) > CCur(Txt(10))) And CCur(Txt(10)) > 0 Then  'Si el monto a pagar > Deuda
                Dim strMensaje$
                strMensaje = "Monto a distribuir es mayor a la deuda del propietario," _
                & vbCrLf & "Desea aplicar la diferencia a la próxima facturación?"
                If Not Respuesta(strMensaje) Then
                    Txt(5).SetFocus
                    Txt(5).SelStart = 0
                    Txt(5).SelLength = Len(Txt(5).Text)
                    Exit Sub
                End If
            End If
            Frame1.Enabled = True
            If Txt(5) = "" Or IsNull(Txt(5)) Or Txt(5) <= 0 Then
                MsgBox "Debe introducir una cantidad válida en el campo Monto", vbExclamation, _
                App.ProductName
                Txt(5).SetFocus
                Exit Sub
            End If
            IntMonto = CCur(Txt(5))
            I = 1
            Command1(0).Enabled = False
        '
            Do While IntMonto > 0 And I <= FlexFacturas.Rows - 1
                For j = 0 To 2
                    vecFPS(I, j) = FlexFacturas.TextMatrix(I, j + 2)
                Next
        '
                If IntMonto - FlexFacturas.TextMatrix(I, 4) >= 0 Then
        '
                        IntMonto = IntMonto - CCur(FlexFacturas.TextMatrix(I, 4))
                        FlexFacturas.TextMatrix(I, 3) = Format(CDbl(FlexFacturas.TextMatrix(I, 4)) + _
                        CDbl(FlexFacturas.TextMatrix(I, 3)), "#,##0.00")
                        FlexFacturas.TextMatrix(I, 4) = 0
                        FlexFacturas.TextMatrix(I, 6) = "SI"
                        Txt(9) = IIf(Txt(9) = "", "", Txt(9) + " / ") + FlexFacturas.TextMatrix(I, 1)
                    Else
            '
                        FlexFacturas.TextMatrix(I, 3) = Format(IntMonto + vecFPS(I, 1), "#,##0.00")
                        FlexFacturas.TextMatrix(I, 4) = Format(FlexFacturas.TextMatrix(I, 2) - _
                        FlexFacturas.TextMatrix(I, 3), "#,##0.00")
                        FlexFacturas.TextMatrix(I, 6) = "SI"
                        Txt(9) = IIf(Txt(9) = "", FlexFacturas.TextMatrix(I, 1), Txt(9) + _
                        " / " + "Abono a Cuenta " + FlexFacturas.TextMatrix(I, 1))
                        IntMonto = 0
                    
                    End If
                
                FlexFacturas.Row = I
                FlexFacturas.Col = 6
                FlexFacturas.CellPictureAlignment = flexAlignRightTop
                Set FlexFacturas.CellPicture = ImgAceptar(0).Picture
                I = I + 1
            Loop
            If IntMonto > 0 Then
                Txt(9) = IIf(Txt(9) = "", "Abono a Prox.Facturación", Txt(9) + _
                " / Abono a Prox.Facturación")
            Else
                IntMonto = 0
            End If
            Txt(10) = Format(CCur(Txt(10)) - CCur(Txt(5)), "#,##0.00")
            Txt(11) = Format(CCur(LblHono) + CCur(Txt(10)), "#,##0.00")
            Command1(1).Caption = "&Deshacer": Frame1.Enabled = True: Cmb(0).Enabled = True
            Cmb(0).SetFocus
        End If
        '
    End Select
    '
    End Sub
    
    '---------------------------------------------------------------------------------------------
    Private Sub Command_Click(Index As Integer)     '
    '---------------------------------------------------------------------------------------------
    'VARIABLES LOCALES
    Dim strMsg As String
    '
    If Index >= 0 And Index <= 3 Then
    
        Call RtnMuestraCalendar(Index)
        B = Index
        
    Else
        '
        If Txt(Index - 3) = "" Then Txt(Index - 3) = 0
        
        If Cmb(Index + 1) <> "CHEQUE" Then
            strMsg = "Debe seleccionar cheque para continuar."
        ElseIf Txt(Index + 2) = "" Then
            strMsg = "Falta Número del documento."
        ElseIf Cmb(Index - 2) = "" Then
            strMsg = "Falta Banco a/c del documento."
        ElseIf CCur(Txt(Index - 3)) = 0 Then
            strMsg = "Falta la cantidad del cheque."
        ElseIf Dat(4) = "" Then
            strMsg = "Falta el beneficiario del cheque."
        End If
        '
        If strMsg <> "" Then
            MsgBox strMsg, vbCritical, App.ProductName
        Else
            Call imprimir_cheque(Txt(Index - 3), Dat(1), Dat(3) = sysCodCaja)
        End If
        '
    End If
    '
    End Sub
    
    Private Sub Command2_Click()
    'VARIABLES LOCALES
    FlexDeducciones.AddItem ("")
    FlexDeducciones.RowHeight(FlexDeducciones.Rows - 1) = 315
    End Sub
    
    '---------------------------------------------------------------------------------------------
    Private Sub Command3_Click(Index As Integer)    '
    '---------------------------------------------------------------------------------------------
    'variables locales                      '*
    Dim ObjRstDep As ADODB.Recordset        '*
    Dim ObjRstCheq As ADODB.Recordset       '*
    Dim strMsg As String                    '*
    Dim m As Long                           '*
    Dim Impre$, booE As Boolean             '*
    Dim X As Printer                        '*
    Dim NDep$, Cta$, Bco$, fecha$, Cja$     '*
    '---------------------------------------
    '
    Select Case Index
        
        Case 0  'CANCELAR
        
            FlexDeducciones.Rows = 2
            With FlexFacturas   'ELIMINA LA MARCA GRAFICA DE LA FACTURA
                .Row = intRow
                .Col = 7
                Set .CellPicture = Nothing
            End With
            '
            'ELIMINA CUALQUIER INFORMACION DEL TEMDEDUCCIONES
            If StrRutaInmueble <> "" Then
            
                Set cnnPropietario = New ADODB.Connection
                cnnPropietario.Open cnnOLEDB + StrRutaInmueble
                cnnPropietario.Execute "DELETE * FROM TemDeducciones WHERE IDPeriodos='" & _
                strRecibo & "'"
                cnnPropietario.Close
                Set cnnPropietario = Nothing
                
            End If
            '
            Call rtnTab(0)
        
        Case 1  'CONFIRMA LAS DEDUCCIONES PROCESADAS
            '
            'EN CAS0 DE HABER DEDUCCIONES
            If Dat(0) = "" Or Dat(2) = "" Then
                
                MsgBox "Faltan Parametros Requeridos" & Chr(13) & "para Procesar esta Deducción" _
                & Chr(13) & "Presione el Boton 'Salir' y Vuelva a Intertarlo", vbExclamation, _
                App.ProductName
                Command3(0).SetFocus
                Exit Sub
                '
            
            End If
            If dtc.Tag = "" Or IsNull(dtc.Tag) Then
                MsgBox "Seleccione un supervisor de la lista, para enviarle la solicitud de aut" _
                & "orización", vbInformation, App.ProductName
                dtc.SetFocus
                Exit Sub
            End If
            '
            Command3(1).Enabled = False
            Command3(0).Enabled = False
            '
            'GUARDA LA INFORMACION EN EL TEMPORAL DEDUCCIONES/"TemDeducciones"
            Call RtnTemDeducciones("\" & Trim(LblCodInm) & "\")
            'si dispone de RealpoPup lo utiliza como medio para solicitar la autorización
            'a un supervisor
            
            If gcNivel > nuAdministrador Then
            
                strMsg = "AUTORIZACION DEDUCCION " & Dat(0) & "/" & Dat(2) & vbNewLine
                'utiliza el programa real popup para enviar mensajer a otro usuario
                
'                If Dir("C:\Archivos de programa\RealPopup\RealPopup.exe") <> "" Then
'
'
'                    m = Shell("C:\Archivos de programa\RealPopup\RealPopup -send ADMINISTRACION " & _
'                    Chr(34) & strMsg & Chr(34) & " -NOACTIVATE")
'
'                End If
                'utiliza la mensajería de sac para enviar los mensajes
                'Winsock1.RemoteHost = dtc.Tag,888
                Dim remoto As String
                Dim datIni As Date, datFin As Date
                
                remoto = dtc.Tag
                Winsock1.Close
                Winsock1.Connect remoto, 888
                datIni = Time
                Do
                    DoEvents
                Loop Until Time >= DateAdd("s", 7, datIni)
                If Winsock1.State = sckClosed Then
                    Command3(1).Enabled = True
                    Command3(3).Enabled = True
                Else
                    Winsock1.SendData strMsg
                End If
            Else
                MsgBox "Usted debe autorizar esta operación en modo local", vbInformation, _
                App.ProductName
            End If
            '
        Case 2  'PORCESAR DEPOSITO
    '   -------------------------------
            '
            With GridCheques(1)
                
                For I = 1 To .Rows - 1
                
                    If .TextMatrix(I, 0) <> "" And Trim(.TextMatrix(I, 0)) <> "TOTAL:" Then
                    
                        If .TextMatrix(I, 3) = "" Then
                        
                            If .TextMatrix(I, 1) <> "PROVINCIAL" Then
                                MsgBox "Ingrese el código de cuenta correspondiente a cada cheq" _
                                & "ue", vbCritical, App.ProductName
                                Exit Sub
                            Else
                                If Dat(10) <> "PROVINCIAL" Then
                                    MsgBox "Ingrese el código de cuenta correspondiente a cada " _
                                    & "cheque", vbCritical, App.ProductName
                                    Exit Sub
                                End If
                                '
                            End If
                            '
                        End If
                        '
                    End If
                '
                Next
                '
            End With
            'verifica datos mínimos requeridos para procesar el depósito
            If Trim(Txt(4)) = "" Then   'nº de depósito
                MsgBox "Introduzaca el Nº de Depósito", vbCritical, App.ProductName
                Exit Sub
            ElseIf Dat(10) = "" Then    'nombre bancpo
                MsgBox "Falta el nombre del banco", vbCritical, App.ProductName
                Exit Sub
            ElseIf Dat(9) = "" Then 'nº cuenta
                MsgBox "Fala el nº de cuenta del depósito", vbCritical, App.ProductName
                Exit Sub
            ElseIf Dat(7) = "" Then
                MsgBox "Seleccione un código de caja", vbCritical, App.ProductName
                Exit Sub
            End If
            '
            Command3(5).Enabled = False
            'asigna valores a lacas variables
            NDep = Trim(Txt(4))
            Bco = Dat(10)
            Cta = Dat(9)
            Cja = Dat(7)
            fecha = DTPicker1
            '
            'comienza el proceso en transacción
            'cnnConexion.BeginTrans
            'Call rtnBitacora("Iniciando Transacción...")
            '
            cnnConexion.Execute "INSERT INTO TDFDepositos(IDdeposito,Banco,Cuenta,Caja,Fecha,Us" _
            & "uario) VALUES ('" & NDep & "','" & Bco & "','" & Cta & "','" & Cja & "','" & _
            fecha & " ','" & gcUsuario & "')"
            Call rtnBitacora("Guarda depósito Nº " & NDep)
            j = 1
            
            On Error GoTo finalizar
            
            With GridCheques(1)
                '
                
                Do Until .TextMatrix(j, 0) = "" Or .TextMatrix(j, 0) = "TOTAL: "
                    
                    
                    cnnConexion.Execute "UPDATE Caja INNER JOIN (TDFCheques INNER JOIN Inmueble" _
                    & " ON TDFCheques.CodInmueble = Inmueble.CodInm) ON Caja.CodigoCaja = Inmue" _
                    & "ble.Caja SET TDFCheques.IDDeposito = '" & NDep & "' WHERE (((TDF" _
                    & "Cheques.Ndoc)='" & .TextMatrix(j, 0) & "') AND ((TDFCheques.Banco)= '" _
                    & .TextMatrix(j, 1) & "') AND ((TDFCheques.Monto)= CCur('" & _
                    .TextMatrix(j, 2) & "')) AND ((TDFCheques.FechaMov)=Date()) AND ((Caja.Codi" _
                    & "goCaja)='" & Cja & "') AND ((TDFCheques.IDTaquilla)=" & IntTaquilla & "));"
                    
                    Call rtnBitacora("Actualiza Doc. " & .TextMatrix(j, 0) & " Banco " & _
                    .TextMatrix(j, 1) & " Monto " & .TextMatrix(j, 2) & " Caja: " & Cja)
                    '
                    j = j + 1
                    '
                Loop
                j = j - 1
                If CCur(Label16(30)) > 0 Then
                '
                    cnnConexion.Execute "UPDATE TDFCheques INNER JOIN (Caja INNER JOIN Inmueble" _
                    & " ON Caja.CodigoCaja = Inmueble.Caja) ON TDFCheques.CodInmueble = Inmuebl" _
                    & "e.CodInm SET TDFCheques.IDDeposito = '" & NDep & "' WHERE (((TDFCheques." _
                    & "Fpago)= 'EFECTIVO') AND ((Caja.CodigoCaja)='" & Cja & "') AND ((isnull(T" _
                    & "DFCheques.IDDeposito))=true) AND ((TDFCheques.FechaMov)=Date()) AND ((TD" _
                    & "FCheques.IDTaquilla)=" & IntTaquilla & ") )"
                    '
                    Call rtnBitacora("Ingresado Efectivo Dep.: " & NDep & " Caja " & Cja)
                End If
    '
            End With
    '
finalizar:
            If Err.Number = -2147217900 Then
                MsgBox "Está introduciendo un número de depósito ya utilizado," & vbCrLf & "Int" _
                & "roduzca un nuevo numero e intente nuevamente", vbExclamation, App.ProductName
                'cnnConexion.RollbackTrans
                Call rtnBitacora(Err.Description)
            ElseIf Err.Number <> 0 Then
                MsgBox Err.Description & vbCrLf & "Comunica este mensaje al administrador del s" _
                & "istema", vbCritical, App.ProductName
                'cnnConexion.RollbackTrans
                Call rtnBitacora(Err.Description)
            Else
                'cnnConexion.CommitTrans
                Call rtnBitacora("Transacción terminada con éxito...")
                'si los cambios fueron realizados con éxito se procede a la impresión de los
                'endosos y del vaoucher de deposito
                If Dat(9) = "" Or Dat(7).Tag = "" Then
                    MsgBox "Parámetros insuficientes para endosar el cheque. Efectue este proce" _
                    & "so en forma manual", vbInformation, App.ProductName
                Else
                    If j > 0 Then
                        MsgBox "Introduzca Cheque para Imprimir el Endoso", vbInformation, "Endosar"
                        For Z = 1 To j
                            Endoso.Banco = Bco
                            Endoso.Cuenta = Cta
                            Endoso.Titular = Beneficiario(Dat(9), IIf(Cja = sysCodCaja, sysCodInm, Dat(7).Tag))
                            Endoso.Endosar
                        Next
                        Call rtnBitacora(j & " Cheques endosados..")
                    End If
                End If
                'imprime el voucher del depósito
                MsgBox "Introduzca el voucher del depósito. Presione Aceptar cuendo este listo " _
                & "para continuar", vbInformation, App.ProductName
                Impre = Printer.DeviceName
                '
                For Each X In Printers
                    '
                    If X.DeviceName = "Citizen 200GX" Or X.DriverName = "CIT9US" Then
                        Set Printer = X
                        booE = True
                    End If
                    '
                Next
                '
                If booE Then
                '
                    If Bco = "VENEZUELA" Then
                        Call dep_venezuela
                    ElseIf Bco = "PROVINCIAL" Then
                        Call dep_provincial
                    ElseIf Bco = "BANESCO" Then
                        Call dep_banesco
                    End If
                    Call rtnBitacora("Impreso Dep. " & dep & " Banco " & Bco)
                    'reasigna la impresora predeterminada
                    For Each X In Printers
                        If X.DeviceName = Impre Then Set Printer = X
                    Next
                Else
                    MsgBox "La impresora no está instalada, no se imprimirá el voucher", _
                    vbInformation, App.ProductName
                    Call rtnBitacora("Depósito no impreso, imp. no instalada..")
                End If
                Call rtnLimpiar_Grid(GridCheques(1))
                'envia un mensaje al usuario de finalizado el proceso
                MsgBox "Procesado Deposito '" & NDep & "' por Bs. " & Label16(31), _
                vbInformation, App.ProductName
                Txt(4) = "": Label16(30) = "0,00": Label16(31) = "0,00"
                '
                Set objRst = New ADODB.Recordset
                '
                objRst.Open "SELECT DISTINCT CodigoCaja as Codigo, DescripCaja  as Descripcion " _
                & "FROM Caja INNER JOIN Inmueble ON Caja.codigoCaja=Inmueble.Caja WHERE CodInm " _
                & "IN (SELECT CodInmueble From TDFCheques WHERE FechaMov=Date() AND (Fpago='EFE" _
                & "CTIVO' Or Fpago='CHEQUE') AND (IDDeposito='' or Isnull(IDDeposito))) ORDER B" _
                & "Y codigocaja;", cnnConexion, adOpenStatic, adLockReadOnly, adCmdText
                '
                For I = 7 To 10
                    Dat(I) = ""
                    If I = 7 Or I = 8 Then Set Dat(I).RowSource = objRst: Dat(I).Refresh
                Next
                '
            End If
            
        Case 3 'CONSULTAR DEPOSITO
        
            If Txt(4) = "" Or IsNull(Txt(4)) Then
                MsgBox "Introduzca el Número del Depósito a Consultar", vbInformation, _
                App.ProductName
                Exit Sub
            End If
            '
            Set ObjRstDep = New ADODB.Recordset
            strNdeposito = Trim(Txt(4))
            ObjRstDep.Open "SELECT TDFDepositos.*, Caja.DescripCaja FROM TDFDepositos INNER JOI" _
            & "N Caja ON TDFDepositos.Caja = Caja.CodigoCaja WHERE TDFDepositos.IDDeposito=" _
            & "'" & strNdeposito & "'", cnnConexion, adOpenStatic, adLockReadOnly, adCmdText
            '
            With ObjRstDep
                If .EOF Then
                    MsgBox "No se encuentra informacion sobre el depósito Nº '" & Txt(4) & "'" _
                    & vbCr & "Verifique este dato e intente nuevamente...", vbCritical, _
                    App.ProductName
                
                Else
                Command3(5).Enabled = True
                .MoveFirst
                Txt(4) = !IDDeposito
                Dat(10) = !Banco
                Dat(9) = !Cuenta
                'Frame3(7).Tag = !Caja
                DTPicker1.Value = !fecha
                ObjRstDep.Close
                Set ObjRstDep = Nothing
                Set ObjRstCheq = New ADODB.Recordset
                ObjRstCheq.Open "SELECT TDFCheques.Ndoc, TDFCheques.Banco, TDFCheques.Monto," _
                    & "TDFCheques.Fpago From TDFCheques WHERE (((TDFCheques.IDDeposito)='" _
                    & strNdeposito & "')) ORDER BY (TDFCheques.Fpago);", _
                    cnnConexion, adOpenStatic, adLockReadOnly, adCmdText
                    
                If Not ObjRstCheq.EOF Then
                    '
                    Dim curCheque@, CurEfectivo@
                    '
                    With ObjRstCheq
                        .MoveFirst
                        Call rtnLimpiar_Grid(GridCheques(0))
                        Call rtnLimpiar_Grid(GridCheques(1))
                        Label16(30) = "0,00"
                        GridCheques(1).Rows = .RecordCount + 3
                        For j = 1 To .RecordCount
                           If !FPago = "CHEQUE" Then
                                For I = 0 To 2
                                    GridCheques(1).TextMatrix(j, I) = _
                                    IIf(I = 2, Format(.Fields(I), "#,##0.00"), .Fields(I))
                                Next
                                curCheque = !Monto + curCheque
                            ElseIf !FPago = "EFECTIVO" Then
                                CurEfectivo = CurEfectivo + !Monto
                                Label16(30) = Format(CurEfectivo, "#,##0.00")
                            End If
                            .MoveNext
                        Next
                    End With
                    ObjRstCheq.Close
                    Set ObjRstCheq = Nothing
                    
                    With GridCheques(1)
                        .MergeRow(.Rows - 1) = True
                            For j = 0 To 1
                                .Col = j
                                .TextMatrix(.Rows - 1, j) = _
                                "TOTAL CHEQUES: ": .Row = .Rows - 1
                                .CellAlignment = flexAlignRightCenter
                            Next
                            .TextMatrix(.Rows - 1, 2) = _
                                Format(curCheque, "#,##0.00")
                            Label16(31) = Format(curCheque + CurEfectivo, "#,##0.00")
                    End With
                    Else
                        Call rtnLimpiar_Grid(GridCheques(1))
                    End If
                    
            End If
            End With
        
        Case 4  'EDITAR DEPOSITO
    '   -------------------------------
            If Txt(4) = "" Or Dat(10) = "" Or Dat(9) = "" Then
                MsgBox "Faltan datos para actualizar la información  de este depósito", _
                vbInformation, App.ProductName
                Exit Sub
            Else
                On Error Resume Next
                'actualiza la tabla depositos
                cnnConexion.BeginTrans
                cnnConexion.Execute "UPDATE TDFDepositos SET IDDeposito = '" & Txt(4) & "'," _
                & "Banco = '" & Dat(10) & "', Cuenta = '" & Dat(9) & "', Fecha = '" & DTPicker1 _
                & "', Usuario = '" & gcUsuario & "' WHERE IDDeposito= '" & strNdeposito & "'"
                'actualiza la tabla cheques
                cnnConexion.Execute "UPDATE TDFCheques SET IDDeposito='" & Txt(4) & "' WHERE ID" _
                & "Deposito='" & strNdeposito & "'"
                If Err.Number = 0 Then
                    cnnConexion.CommitTrans
                    Call rtnBitacora("Depósito Actulizado de " & strNdeposito & " a " & Txt(4))
                    MsgBox "Depósito Actualizado...", vbInformation, App.ProductName
                Else
                    cnnConexion.RollbackTrans
                    MsgBox Err.Description, vbCritical, "Error " & Err.Number
                End If
            
           End If
        Case 5  'reimpresion
                'valida que el depósito este efectuado
                Set ObjRstDep = New ADODB.Recordset
'                ObjRstDep.Open "SELECT * FROM TDFDepositos WHERE IDDeposito='" & Txt(4) & "'", _
'                cnnConexion, adOpenKeyset, adLockOptimistic, adCmdText
'
                ObjRstDep.Open "SELECT TDFDepositos.*, Inmueble.CodInm FROM TDFDepositos INNER " _
                & "JOIN Inmueble ON TDFDepositos.Caja = Inmueble.Caja WHERE TDFDepositos.IDDepo" _
                & "sito ='" & Txt(4) & "'", cnnConexion, adOpenKeyset, adLockOptimistic, adCmdText

                If ObjRstDep.EOF And ObjRstDep.BOF Then
                
                    MsgBox "Este depósito no ha sido procesado, presione el botón 'DEPOSITAR'"
                    ObjRstDep.Close
                    Set rstdep = Nothing
                    Exit Sub
                
                End If
                
                '
                'aqui imprime el voucher del depósito
                Impre = Printer.DeviceName
                For Each X In Printers
                    If X.DeviceName = "Citizen 200GX" Or X.DriverName = "CIT9US" Then
                        Set Printer = X
                        booE = True
                    End If
                Next
                '
                If booE Then
                    
                    If Dat(10) = "VENEZUELA" Then
                        Call dep_venezuela(Beneficiario(Dat(9), IIf(ObjRstDep!Caja = sysCodCaja, _
                        sysCodInm, ObjRstDep!CodInm)))
                    ElseIf Dat(10) = "PROVINCIAL" Then
                        Call dep_provincial(Beneficiario(Dat(9), IIf(ObjRstDep!Caja = sysCodCaja, _
                        sysCodInm, ObjRstDep!CodInm)))
                    ElseIf Dat(10) = "BANESCO" Then
                        Call dep_banesco(Beneficiario(Dat(9), IIf(ObjRstDep!Caja = sysCodCaja, _
                        sysCodInm, ObjRstDep!CodInm)))
                    End If
                    'reasigna la impresora predeterminada
                    For Each X In Printers
                        If X.DeviceName = Impre Then Set Printer = X
                    Next
                Else
                    MsgBox "La impresora no está instalada, no se imprimirá el voucher", _
                    vbInformation, App.ProductName
                End If
                ObjRstDep.Close
                Set ObjRstDep = Nothing
    End Select
    '
    End Sub

    Private Sub Dat_Change(Index As Integer)
    'VARIABLES LOCALES
    If Index = 11 Then
        With ADOcontrol(4).Recordset
            .MoveFirst
            .Find "NombreUsuario='" & Dat(11) & "'"
            If Not .EOF Then Dat(11).Tag = !IP
        End With
    '
    End If
    '
    End Sub

    'Rev.15/08/2002-------------------------------------------------------------------------------
    Private Sub Dat_Click(Index%, Area%)
    '---------------------------------------------------------------------------------------------------
    '
    If Area = 2 Then    'se ejecuta al seleccionar un elemento de la lista
    '
        Select Case Index
    '
            Case 0
    '
                IntMonto = 0
                curHono = 0   'Declara variables globales de módulo en cero
                Call RtnBuscaInmueble("Inmueble.CodInm", Dat(0))
                    
    '
            Case 1  'realiza la rutina anterior a traves del nombre del inmueble
            
                IntMonto = 0
                curHono = 0   'Declara variables globales de módulo en cero
                Call RtnBuscaInmueble("Inmueble.Nombre", Dat(1))
            
            Case 2  'llama la rutina buscapropietario para encontrar el nombre
                
                IntMonto = 0: curHono = 0   'Declara variables globales de módulo en cero
                Frame1.Enabled = False
                Call Desmarca
                Txt(5) = "0,00"
                If Dat(2) = "" Then Exit Sub
                Call BuscaPropietario("Codigo", Dat(2), "Nombre", DatProp)
                With Cmb(1)
                    .Enabled = True
                    .SetFocus
                    .ListIndex = 1
                End With
                Cmb(0) = ""
                For I = 1 To 3: Txt(I) = ""
                Next
                Call RtnFlex(Dat(2), FlexFacturas, IntMesesMora, IntHonoMorosidad, _
                    FlexFacturas.Cols, Txt(11), cnnPropietario, , True)
                LblHono = Txt(11)
                
                'campo deuda y deuda + honorarios
                If IsNumeric(Txt(11)) And IsNumeric(Txt(10)) Then
                    'reconvertimos los montos a bolivares 2007
                    Txt(10) = Format(CCur(Txt(10)) * 1000, "#,##0.00")
                    Txt(11) = Format(CCur(Txt(11)) + CCur(Txt(10)), "#,##0.00")
                    
                Else
                    Txt(10) = "0,00"
                    Txt(11) = "0,00"
                End If
                curAbono = IIf(FlexFacturas.TextArray(12) <> "", FlexFacturas.TextArray(12), 0)
                
            Case 5  'Busca la descripcion del codigo de cuenta
                If Dat(5) = "" Then Exit Sub
                Call RtnBuscaIngreso("CodGasto", Dat(5), "Titulo", Dat(6))
            
            Case 6  'Busca el codigo de cuenta correspondiente a la descripcion seleccionada
                Call RtnBuscaIngreso("Titulo", Dat(6), "CodGasto", Dat(5))
                
            Case 7  'Busca la Caja Seleccionada
                
                Call RtnBuscaCaja(8, 7)
                
            Case 8  'Busca La caja seleccionada
                Call RtnBuscaCaja(7, 8)
                
            Case 9, 10  'Busca coincidencia en inf. de cuenta bancaria
                If Index = 9 Then Dat(10) = Dat(9).BoundText
                If Index = 10 Then Dat(9) = Dat(10).BoundText
                
        End Select
                '----------''
    End If
    End Sub
    
    Private Sub Dat_GotFocus(Index As Integer)
    '
        Select Case Index
    '
            Case 5  'Codigo del Gastos o Transacción
    '       --------------------------------------------------
                '
                If Dat(0) = sysCodInm Then
                
                    Dat(5) = strCodIV   'COD.INM='9999' ASIGNA COD.DE INGRESOS VARIOS
                    
                Else
                    For I = 1 To FlexFacturas.Rows - 1
                        If FlexFacturas.TextMatrix(I, 6) = "SI" Then
                            If CCur(FlexFacturas.TextMatrix(I, 4)) = 0 Then
                                Dat(5) = strCodPC
                                If UCase(Txt(9)) Like "ABONO*" Then Dat(5) = strCodAbonoCta
                                If UCase(Txt(9)) Like "CH*" Then Dat(5) = strCodRCheq
                                Exit For
                            Else
                                Dat(5) = strCodAbonoCta
                            End If
                        End If
                    Next
                    If Dat(5) = "" Then Dat(5) = IIf(I = 2 And FlexFacturas.TextArray(10) = "", _
                    IIf(Left(Dat(2), 1) = "U", strCodIV, strCodAbonoFut), strCodIV)
                End If
                Dat(6).Enabled = True
                If Dat(5) = "" Then MsgBox "Debe Introducir algun valor..", _
                vbInformation: Dat(Index).SetFocus: Exit Sub
                Call RtnBuscaIngreso("CodGasto", Dat(5), "Titulo", Dat(6))
        End Select
    End Sub
    
    Private Sub Dat_KeyPress(Index As Integer, KeyAscii As Integer)
    If Index = 0 Then Call Validacion(KeyAscii, "1234567890")
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then   'Si Presiona la Tecla Entrar
        
        Select Case Index
            
            Case 0  'codigo inmueble busca el nombre del inmueble
            
                IntMonto = 0: curHono = 0   'Declara variables globales de módulo en cero
                If Dat(Index) = "" Then Dat(1).SetFocus: Exit Sub
                Call RtnBuscaInmueble("Inmueble.CodInm", Dat(0))
                
            Case 1  'nombre del inmueble busca el codigo correspondiente
            
                IntMonto = 0: curHono = 0   'Declara variables globales de módulo en cero
                If Dat(Index) = "" Then Dat(0).SetFocus: Exit Sub
                Call RtnBuscaInmueble("Inmueble.Nombre", Dat(1))
                
                
            Case 2  'codigo de apartamento busca nombre del propietario
                    'llama la rutina buscapropietario para encontrar el nombre
                    'del propietario de ese apartamento
                    Frame1.Enabled = False
                    Call Desmarca
                    Txt(5) = "0,00"
                    If Dat(2) = "" Then DatProp.SetFocus: Exit Sub
                    ADOcontrol(2).Refresh
                    IntMonto = 0: curHono = 0   'Declara variables globales de módulo en cero
                    Call BuscaPropietario("Codigo", Dat(2).Text, "Nombre", DatProp)
                    Txt(10) = Format(CCur(Txt(10)) * 1000, "#,##0.00")
                    If Dat(0) <> "" And DatProp <> "" Then
                        Cmb(1).Enabled = True: Cmb(1).SetFocus: Cmb(1).ListIndex = 1
                        Call RtnFlex(Dat(2), FlexFacturas, IntMesesMora, IntHonoMorosidad, _
                            FlexFacturas.Cols, Txt(11), cnnPropietario, Dat(0), True)
                            LblHono = IIf(Txt(11) > 0, Txt(11), "0,00")
                            If Txt(10) = "" Then Txt(10) = 0
                            If Txt(11) = "" Then Txt(11) = 0
                            If Txt(11) > 0 Then
                                Txt(11) = Format(CCur(Txt(11)) + CCur(Txt(10)), "#,##0.00")
                            Else
                                Txt(11) = Txt(10)
                            End If
                            curAbono = IIf(FlexFacturas.TextArray(12) <> "", FlexFacturas.TextArray(12), 0)
                    End If
                    
            Case 3  'nombre del propietario busca el n.de apartamento
                    
                    IntMonto = 0: curHono = 0   'Declara variables globales de módulo en cero
            
            Case 5  'codigo de cuenta de ingreso busca la descripcion
                    If Dat(Index) = "" Then MsgBox "Debe Introducir algun valor..", vbInformation: Dat(Index).SetFocus: Exit Sub
                    Call RtnBuscaIngreso("CodGasto", Dat(5), "Titulo", Dat(6))
            
            Case 6  'con la descripcion busca la cuenta de ingreso_egreso
                    If Dat(Index) = "" Then MsgBox "Debe Introducir algun valor..", vbInformation: Dat(Index).SetFocus: Exit Sub
                    Call RtnBuscaIngreso("Titulo", Dat(6), "CodGasto", Dat(5))
            
           
        End Select
        
    End If
    
    If KeyAscii = 8 Then 'Si presiona la tecla BackSpace
        
        Select Case Index
            
            Case 1: If Dat(Index) = "" Then Dat(Index - 1).SetFocus
            
            Case 2: If Dat(Index) = "" Then Dat(Index - 1).SetFocus
            
            Case 5: If Dat(Index) = "" Then Cmb(0).SetFocus
            
            Case 6: If Dat(Index) = "" Then Dat(Index - 1).SetFocus
                
        End Select
        
    End If
    
    End Sub
    
    Private Sub DataGrid1_DblClick()
    'variables locales
    Dim IntClave&
    '
    IntClave = ADOcontrol(3).Recordset.Fields("IndiceMovimientoCaja")
    '
    MousePointer = vbHourglass
    DoEvents
    ADOcontrol(0).Refresh
    With ADOcontrol(0).Recordset
        
        .MoveFirst
        .Find "IndiceMovimientoCaja = '" & IntClave & "'"
        Call RtnAvanza
        Call mostrar_area(IIf(!FormaPagoMovimientoCaja = "EFECTIVO", 3, 4))
    SSTab1.tab = 0
    End With
    MousePointer = vbDefault
    '
    End Sub
    
    'Rev.26/08/2002------Rutina que busca inf. del propietario seleccionado-----------------------
    Private Sub DatProp_Click(Area As Integer)
    '---------------------------------------------------------------------------------------------
    'al hacer click en algun elemento de la lista
    If Area = 2 Then    'se ejecuta la rutina que busca el codigo del
                        'propietario seleccionado
        Frame1.Enabled = False
        Call Desmarca
        Txt(5) = "0,00"
        Call BuscaPropietario("Nombre", DatProp.Text, "Codigo", Dat(2))
        Cmb(1).Enabled = True
        Cmb(1).SetFocus
        Cmb(1).ListIndex = 1
        Call RtnFlex(Dat(2), FlexFacturas, IntMesesMora, IntHonoMorosidad, _
            FlexFacturas.Cols, Txt(11), cnnPropietario)
        LblHono = Txt(11)
        Txt(11) = Format(CCur(Txt(11)) + CCur(Txt(10)), "#,##0.00")
        curAbono = IIf(FlexFacturas.TextArray(12) <> "", FlexFacturas.TextArray(12), 0)
    End If
    '
    End Sub
    

    
    'Rev.26/08/2002-------------------------------------------------------------------------------
    Private Sub DatProp_KeyPress(KeyAscii As Integer)
    '---------------------------------------------------------------------------------------------
    '
    KeyAscii = Asc(UCase(Chr(KeyAscii)))    'Convierte todo en mayúsculas
    If KeyAscii = 13 Then   'Presiona {enter}
        If DatProp = "" Then
            MsgBox "Debe Introducir un numero de apartamento", vbInformation, App.ProductName
            DatProp.SetFocus
            Exit Sub
        End If
        
        Call Desmarca
        Txt(5) = "0,00"
        Call BuscaPropietario("Nombre", DatProp.Text, "Codigo", Dat(2))
        If Dat(2) <> "" And DatProp <> "" Then
        Cmb(1).Enabled = True: Cmb(1).SetFocus: Cmb(1).ListIndex = 1
        Call RtnFlex(Dat(2), FlexFacturas, IntMesesMora, IntHonoMorosidad, _
            FlexFacturas.Cols, Txt(11), cnnPropietario)
        LblHono = Txt(11)
        Txt(11) = Format(CCur(Txt(11)) + CCur(Txt(10)), "#,##0.00")
        curAbono = IIf(FlexFacturas.TextArray(12) <> "", FlexFacturas.TextArray(12), 0)
        End If
    End If
    '
    End Sub
    

Private Sub dtc_Change()
With ADOcontrol(4).Recordset
    .MoveFirst
    .Find "NombreUsuario ='" & dtc & "'"
    If Not .EOF Then dtc.Tag = !IP
End With
End Sub

    Private Sub FlexDeducciones_EnterCell()
    With FlexDeducciones
        If .Col = 2 And .Row >= 1 And .Text <> "" Then .Text = CCur(.Text)
    End With
    End Sub

Private Sub FlexDeducciones_KeyPress(KeyAscii As Integer)
'Permite escribir dentro del Grid
With FlexDeducciones
    '
    Select Case .Col
        '
        Case 0  'Codigo del gasto
        '--------------------
            Call Validacion(KeyAscii, "01234567890")
            If KeyAscii > 26 And Len(.Text) <= 6 Then .Text = FlexDeducciones.Text & Chr(KeyAscii)
            '
        Case 1  'descripcion
        '--------------------
            If KeyAscii = 39 Then KeyAscii = 0  'SI PRESIONA EL CARACTER " ' " apostofre
            If KeyAscii > 26 And Len(.Text) <= 99 Then .Text = .Text & UCase(Chr(KeyAscii))
            '
        Case 2  'Monto
        '--------------------
            If KeyAscii = 46 Then KeyAscii = 44
            Call Validacion(KeyAscii, "0123456789,")
            If KeyAscii > 26 Then .Text = .Text & Chr(KeyAscii)
        '
    End Select
    '
    If KeyAscii = 8 Then    'backspace
        If Len(.Text) > 0 Then .Text = Left(.Text, Len(.Text) - 1)
    ElseIf KeyAscii = 13 Then   'enter
        If .Col = 0 Then
            If Trim(Label16(12)) <> "" Then Call Busca_Reintegro(Trim(.Text), .RowSel)
            If Trim(Label16(12)) = "" Then
                ADOcontrol(1).Recordset.MoveFirst
                ADOcontrol(1).Recordset.Find "CodGasto='" & .Text & "'"
                If Not ADOcontrol(1).Recordset.EOF And Not ADOcontrol(1).Recordset.BOF Then
                    .TextMatrix(.RowSel, 1) = IIf(IsNull(ADOcontrol(1).Recordset("Titulo")), _
                    "---Edite la descripción de este gasto---", ADOcontrol(1).Recordset("Titulo"))
                    .Col = 1
                Else
                    MsgBox "El Código del Gasto correspondiente al reintegro no está registrado" _
                    & " el catálogo de caja", vbInformation, App.ProductName
                    .SetFocus
                End If
            End If
        
        ElseIf .Col = 2 Then    'columna del monto
            '
            If .Row = 0 Then .Row = 1
            If .Text = "" Then .Text = 0
            For I = 1 To .Rows - 1
            
                If IsNumeric(.TextMatrix(I, 2)) Then
                
                    If Label16(0).Tag = "" Then Label16(0).Tag = 0
                    Label16(0).Tag = CCur(.TextMatrix(I, 2)) + CCur(Label16(0).Tag)
                    
                End If
                
            Next I
            '
            If CCur(.Text) > CCur(.TextMatrix(1, 3)) Then
                MsgBox "No puede deducir un monto superior al monto abonado a la factura", vbCritical, _
                App.ProductName
                .Text = Format(.TextMatrix(.RowSel, 3), "#,##0.00")
                Label16(0) = .Text
                .SetFocus
            Else
                
                If CCur(Label16(0).Tag) > CCur(.TextMatrix(1, 3)) Then
                    MsgBox "La sumatoria de las deducciones sobrepasan el total de la factura", _
                    vbExclamation, App.ProductName
                    .Text = Format(CCur(.TextMatrix(1, 3)) - CCur(Label16(0)), "#,##0.00")
                    .SetFocus
                    Else
                        If Not IsNumeric(Label16(0)) Then Label16(0) = 0
                    Label16(0) = Format(CCur(Label16(0)) + CCur(.Text), "#,##0.00")
                    .AddItem ("")
                    .RowHeight(.Rows - 1) = 315
                    .Row = .Row + 1
                    .Col = 0
                End If
            End If
            Label16(0).Tag = ""
            '
        ElseIf .Col + 1 < .Cols - 1 Then .Col = .Col + 1
        '
        End If
        '
    End If
    '
End With
'
End Sub

    Private Sub FlexDeducciones_KeyUp(KeyCode%, Shift%)
    'permite editar el contenido del flexdeducciones
    If KeyCode = 46 Then FlexDeducciones.Text = ""
    End Sub

    Private Sub FlexDeducciones_LeaveCell()
    With FlexDeducciones
        If .Col = 2 And .Row >= 1 And .Text <> "" Then
            .Text = Format(.Text, "#,##0.00")
        End If
    End With
    End Sub

    Private Sub FlexDeducciones_RowColChange()
    'MsgBox "cambio"
    Dim I%, Total@
    With FlexDeducciones
        For I = 1 To .Rows - 1
            If IsNumeric(.TextMatrix(I, 2)) Then
                Total = Total + CCur(.TextMatrix(I, 2))
                Label16(0) = Format(Total, "#,##0.00 ")
            End If
        Next I
        
    End With
    End Sub

    'Rev.26/08/2002----Procedimientos que responden al evento click sobre el grid que muestra-----
    Private Sub FlexFacturas_Click()    '----------------------------------las fasturas pendientes
    '----------------------------------------------------------------------------de un propietario
    'variables locales
    Dim TxtBuscar$, byDonde%, byPos1%, byPos2%
    '
    If FlexFacturas.ColSel <> 7 Then    'se ejecuta si hizo clik sobre cualquier
    '                                   columna menos sobre las de deducciones
        FlexFacturas.Col = 6
        If FlexFacturas.CellPicture = ImgAceptar(0).Picture Then    'Si la factura ya está marcada
    '                                                           para cancelar, desmarca la factura
            
            With FlexFacturas
    '
                .Row = FlexFacturas.RowSel
                .Col = 6
                Set .CellPicture = Nothing
                If Txt(5) = "" Then Txt(5) = 0
                    Txt(5) = Format(CCur(Txt(5)) - (CCur(.TextMatrix(.Row, 3) - vecFPS(.RowSel, 1))), "##,##0.00")
                    Txt(10) = Format(CCur(Txt(10) + CCur(.TextMatrix(.Row, 3)) - vecFPS(.RowSel, 1)), "#,##0.00")
                    Txt(11) = Format(CCur(Txt(10) + CCur(LblHono)), "#,##0.00")
                    .TextMatrix(.RowSel, 3) = Format(vecFPS(.RowSel, 1), "#,##0.00")
                    .TextMatrix(.RowSel, 4) = Format(vecFPS(.RowSel, 2), "#,##0.00")
                    .TextMatrix(.RowSel, 4) = Format(.TextMatrix(.RowSel, 4), "#,##0.00")
                    .TextMatrix(.Row, 6) = "NO"
                    TxtBuscar = .TextMatrix(.RowSel, 1)  'Busca el mes a borrar
                    Txt(9) = Replace(Txt(9), TxtBuscar & " /", "")
                    Txt(9) = Replace(Txt(9), TxtBuscar, "")
                    
                    'byDonde = InStr(Txt(9).Text, TxtBuscar)
'                    If byDonde > 0 Then  'Marca la posicion donde se encuenta
'                        byPos1 = InStr(byDonde, Txt(9), "/")
'                        If byPos1 Then
'                            byPos2 = InStr(byDonde, Txt(9), "/")
'                            If Not byPos2 Then byPos2 = Len(Txt(9))
'                        Else
'                            byPos1 = 1
'                            byPos2 = Len(Txt(9))
'                        End If
'                        Txt(9).SetFocus
'                        Txt(9).SelStart = byPos1 - 1
'                        Txt(9).SelLength = byPos2: SendKeys (Chr(8))
'                    End If
                    
    '
            End With
    '
        Else    'La factura será marcada pará cancelar (total/parcial)
    '   -----------------------------------------------------------------------
            Call RtnFactura(True)
        End If
    Else
        'se ejecuta al hacer clik sobre la col. deducciones
        'verifica que la factura este marcada para cancelar
        With FlexFacturas
            .Row = .RowSel
            intRow = .Row
            .Col = 6
            If .CellPicture = 0 Then
                MsgBox "Debe marcar primero la factura para cancelar..", vbInformation, _
                App.ProductName
                Exit Sub
            End If
            .Col = 7
            .CellPictureAlignment = flexAlignCenterTop
            Set .CellPicture = ImgAceptar(0).Picture
            Label16(12) = " " & .TextMatrix(.Row, 0)
            Label16(13) = " " & .TextMatrix(.Row, 1)
            Call rtnLimpiar_Grid(FlexDeducciones)
            FlexDeducciones.TextMatrix(1, 3) = Format(vecFPS(.Row, 2) - _
            .TextMatrix(.Row, 4), "#,##0.00")
            '
        End With
        Call rtnTab(2)
        LblCodInm = " " & Dat(0)
        Label16(3) = " " & Dat(1)
        Label16(5) = " " & Dat(2)
        Label16(6) = " " & DatProp
        Label16(7) = " " & Dat(3)
        Label16(8) = " " & Dat(4)
        Label16(0) = "0,00"
        'Call rtnDistribuye
    End If
    '
    End Sub

    '26/08/2002-------Rutina que contiene el procedimiento cuando una celda de la columna 3-------
    Private Sub FlexFacturas_EnterCell()    'toma el foco
    '---------------------------------------------------------------------------------------------
    '
    With FlexFacturas
    '
        If .Col = 3 And .Row > 0 Then
            If .Text = "" Then .Text = 0
            .Text = CCur(.Text)
        End If
    '
    End With
    '
    End Sub

    '26/08/2002---Rutina que permite introducir datos directamente al Grid de las factuas
    Private Sub FlexFacturas_KeyPress(KeyAscii As Integer)  'Permite aplicar abonos directamente
    '--------------------------------------------------------a una factura en cualquier orden
    '
    With FlexFacturas
    '
        If .Col = 3 And .Row > 0 Then
    '
            If KeyAscii = 46 Then KeyAscii = 44 'convierte punto en coma
            Call Validacion(KeyAscii, "1234567890,")
            If KeyAscii = 0 Then Exit Sub
            
            If KeyAscii = 8 Then        'presiona {backspace}
                If .Text = "" Then Exit Sub
                .Text = Left(.Text, Len(.Text) - 1)
            ElseIf KeyAscii = 13 Then   'Preiona {enter}
                .Col = 6
                If .Text = "" Then .Text = 0
                If .TextMatrix(.RowSel, 3) = "" Then .TextMatrix(.RowSel, 3) = 0
                If CCur(.TextMatrix(.RowSel, 3)) > CCur(.TextMatrix(.RowSel, 4)) Then
                    MsgBox "Monto del Abono mayor al total de la factura", vbCritical, _
                    App.ProductName
                    .Text = 0
                    Exit Sub
                ElseIf FlexFacturas.CellPicture = ImgAceptar(0).Picture Then
                    MsgBox "Factura ya está marcada para cancelar, si desea desmarcarla" _
                        & vbCrLf & "haga click sobre ella", vbInformation, App.ProductName
                    Exit Sub
                End If
                .Col = 3
                If .Text = 0 Or (curAbono > 0 And .Row = 1) Then
                    Call RtnFactura(True)
                Else
                    Call RtnFactura(False)
                End If
                If .RowSel + 1 <= (.Rows - 1) Then
                    .Row = .RowSel + 1
                End If
                
            Else    'presiona un caracter válido
                If .Row > 1 Then .Text = .Text & Chr(KeyAscii)
            End If
        End If
    '
    End With
    '
    End Sub

    '23/08/2002---Rutina que borra el contenido (total/parcial) de la celda seleccionada----------
    Private Sub FlexFacturas_KeyUp(KeyCode As Integer, Shift As Integer)
    '---------------------------------------------------------------------------------------------
    '
    With FlexFacturas
    '
        If .Col = 3 And .Row > 0 Then
            If KeyCode = 46 Then    'Delete
            .Text = ""
            End If
        End If
    '
    End With
    '
    End Sub
    '26/08/2002------Rutina que maneja los sucesos que ocurren cuando un celda de la columna 3----
    Private Sub FlexFacturas_LeaveCell()    'pierde el foco
    '---------------------------------------------------------------------------------------------
    '
    With FlexFacturas
    '
        If .Col = 3 And .Row > 0 Then
            .Text = Format(.Text, "#,##0.00")
        End If
    '
    End With
    '
    End Sub


    Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    Dim bGuardar As ComctlLib.Button
    Dim sPregunta As String
    If KeyCode = 71 And Shift = 2 Then  'ctrl + g (guardar registro)
        If Toolbar1.Buttons("Save").Enabled Then
            sPregunta = "¿Está seguro de guardar esta transacción?"
            If Respuesta(sPregunta) = True Then
                Set bGuardar = Toolbar1.Buttons("Save")
                Call Toolbar1_ButtonClick(bGuardar)
            End If
        End If
    ElseIf KeyCode = 78 And Shift = 2 Then  'ctrl + n (nuevo registro)
        If Toolbar1.Buttons("New").Enabled Then
            Set bGuardar = Toolbar1.Buttons("New")
            Call Toolbar1_ButtonClick(bGuardar)
        End If
    
    End If
    End Sub

    'Rev.15/08/2002-------------------------------------------------------------------------------
    Private Sub Form_Load() '
    '---------------------------------------------------------------------------------------------
    '
    'Crea un espacio de trabajo para las op. de caja
    'Dim WrkCaja As Workspace    'Crea un espacion de trabajo
    Dim F%, strSQL$ 'Lista de Transacciones para actualizar el Grid
    '
    'Set WrkCaja = CreateWorkspace("", "Admin", "")
    'Set Endoso = New STEndoso
    Text3 = Date
    If gcNivel = nuADSYS Then Text2.Visible = True: Text3.Visible = True
    
    Lbl(0) = LoadResString(101)
    Lbl(3) = LoadResString(102)
    Lbl(2) = LoadResString(103)
    Lbl(17) = LoadResString(119)
    Lbl(14) = LoadResString(120)
    Command3(2).Picture = LoadResPicture("Deposito1", vbResIcon)
    Command3(3).Picture = LoadResPicture("Deposito", vbResIcon)
    Command3(4).Picture = LoadResPicture("Deposito2", vbResIcon)
    SSTab1.tab = 0
    MntCalendar.Value = Date
    'Carga la lista de bancos------------------------------------------------------------------
    Call rtnFrmActivo
    Set objRst = New ADODB.Recordset
    objRst.Open "Bancos", cnnConexion, adOpenKeyset, adLockReadOnly, adCmdTable
    objRst.MoveFirst
    '
    F = 15
    Do Until objRst.EOF
        F = F + 15
        DoEvents
        Call RtnProUtility("Cargando Lista de Bancos....", F)
        For I = 2 To 4
            Cmb(I).AddItem (objRst.Fields("NombreBanco"))
        Next
        objRst.MoveNext
    Loop
    objRst.Close
    Set objRst = Nothing
    '   ------------------------------------------------------------------------------------------
    '   Llena Recorset con las transacciones del cajero para actualizar los controles
    Call RtnProUtility("Seleccionando Transacciones efectuadas hoy....", 2000)
    ADOcontrol(0).ConnectionString = cnnOLEDB + gcPath + "\sac.mdb"
    ADOcontrol(0).CommandType = adCmdTable
    ADOcontrol(0).RecordSource = "MovimientoCaja"
    DoEvents
    ADOcontrol(0).Refresh
    
    ADOcontrol(0).Recordset.Filter = "FechaMovimientoCaja=#" & Format(Date, "d/m/yy") & "# AND " _
    & "IDTaquilla=" & IntTaquilla
    Call RtnEstado(6, Toolbar1, ADOcontrol(0).Recordset.EOF Or ADOcontrol(0).Recordset.BOF)
    Call RtnProUtility("Configurando Presentación en pantalla....", 3000)
    '
    strSQL = "SELECT MC.IDTaquilla, MC.InmuebleMovimientoCaja, I.Nombre, Mc.AptoMovimiento" _
    & "Caja, I.Caja, MC.NumeroMovimientoCaja, MC.MontoMovimientoCaja, MC.FormaPagoMovim" _
    & "ientoCaja, MC.FechaMovimientoCaja, MC.IndiceMovimientoCaja FROM inmueble as I IN" _
    & "NER JOIN MovimientoCaja as MC ON I.CodInm = Mc.InmuebleMovimientoCaja Where (((" _
    & "MC.IDTaquilla) = " & IntTaquilla & ") And ((MC.FechaMovimientoCaja) = DATE()))" _
    & "ORDER BY MC.IndiceMovimientoCaja DESC"
    '
    ADOcontrol(3).ConnectionString = cnnOLEDB + gcPath + "\sac.mdb"
    ADOcontrol(3).CommandType = adCmdText
    ADOcontrol(3).RecordSource = strSQL
    ADOcontrol(3).Refresh
    DoEvents
    ADOcontrol(2).RecordSource = "SELECT * FROM Propietarios ORDER BY Codigo"
    ADOcontrol(2).CommandType = adCmdText
    
   '
    If ADOcontrol(0).Recordset.RecordCount > 0 Then 'si tiene transacciones registradas
    
        ADOcontrol(0).Recordset.MoveFirst
        Call RtnProUtility("Buscando Información Necesaria....", 4500)
        Call BuscaInmueble("Inmueble.CodInm", Dat(0))   'Busca Inf. del Inmueble
        Set objRst = ObjCmd.Execute
        Call RtnProUtility("Generando Espacios de Trabajo....", 5800)
    '
        With objRst 'Inicializa las variables temporales
    '
            If objRst.EOF Then Exit Sub
            Dat(1) = .Fields("Nombre")
            Dat(3) = .Fields("Caja")
            Dat(4) = .Fields("Descripcaja")
            StrRutaInmueble = gcPath & .Fields("Ubica") & "inm.mdb"
            strCodIV = .Fields("CodIngresosVarios")
            strCodHA = .Fields("CodHA")
            strCodAbonoCta = .Fields("CodAbonoCta")
            strCodAbonoFut = .Fields("CodAbonoFut")
            strCodPC = .Fields("CodPagoCondominio")
            strCodCCHeq = .Fields("CodCCheq")
            strCodRCheq = .Fields("CodRCheq")
            ADOcontrol(1).ConnectionString = cnnOLEDB + StrRutaInmueble
            ADOcontrol(1).CommandType = adCmdText
            ADOcontrol(1).RecordSource = "SELECT CodGasto,Titulo  FROM TGastos WHERE CodGasto L" _
            & "ike '90%' OR CodGasto Like '8%' OR Fondo=True;"
            ADOcontrol(1).Refresh
            DoEvents
    '
        End With
        
        objRst.Close
        Set objRst = Nothing
        cnnPropietario.Open cnnOLEDB + StrRutaInmueble
        ADOcontrol(2).ConnectionString = cnnPropietario
        ADOcontrol(2).Refresh
    '
        With ADOcontrol(2).Recordset
            .Find "Codigo='" & Dat(2) & "'"
            If .EOF Then
                .MoveFirst
                .Find "Codigo='" & Dat(2) & "'"
            End If
            DatProp = .Fields("Nombre")
        End With
        cnnPropietario.Close
        Set cnnPropietario = Nothing
    '   -----------------------------------------
'        If Dat(6) Like "INGRESO*" Then
'            Dat(5) = strCodIV
'        Else
'            If Dat(6) = "" Then Dat(6) = "PAGO CONDOMINIO"
'            Call RtnBuscaIngreso("Titulo", Trim(Dat(6)), "CodGasto", Dat(5))
'        End If
    '   -----------------------------------------
    '
    End If
    Call RtnProUtility("Finalizando...", 6015)
    '
    With FlexFacturas
        .FontWidth = 3.5
    '   Configura la presentación del Grid(Titulos, Ancho de columna, N° de Filas)
    '   debe dejarse al ancho en esta parte porque la propiedad tag esta utilizada
        Call centra_titulo(FlexFacturas)
        .ColWidth(0) = 1000
        .ColWidth(1) = 700
        .ColWidth(2) = 1100
        .ColWidth(3) = 1100
        .ColWidth(4) = 1100
        .ColWidth(5) = 1100
        .ColWidth(6) = 700
        .ColWidth(7) = 500
        .ColWidth(8) = 0
    End With
    '
    With FlexDeducciones
    
        .FontWidth = 3.5
        .RowHeightMin = 315
        Call centra_titulo(FlexDeducciones, True)
        
    End With
    With ADOcontrol(4)
        .CursorLocation = adUseClient
        .CursorType = adOpenKeyset
        .ConnectionString = cnnOLEDB & gcPath & "\tablas.mdb"
        .RecordSource = "SELECT * FROM Usuarios WHERE Nivel <= " & nuAdministrador & _
        " AND LogIn = True"
        .Refresh
        DoEvents
    End With
    rstlocal(0).Open "SELECT * FROM Inmueble WHERE Inactivo=False ORDER BY CodInm ", cnnConexion, _
    adOpenKeyset, adLockOptimistic, adCmdText
    rstlocal(1).Open "SELECT * FROM Inmueble WHERE Inactivo=False ORDER BY Nombre", cnnConexion, _
    adOpenKeyset, adLockOptimistic, adCmdText
    Set Dat(0).RowSource = rstlocal(0)
    Set Dat(1).RowSource = rstlocal(1)
    '
    End Sub
    
    '------------------------------------------------------
    Private Sub Form_Resize()   '
    '------------------------------------------------------
    If Me.WindowState = vbMaximized Then
        DoEvents
        SSTab1.Left = (Screen.Width - (Screen.TwipsPerPixelY * 4) - SSTab1.Width) / 2
        Lbl(7).Left = SSTab1.Left
        Lbl(19).Left = SSTab1.Left
        Txt(13).Left = Lbl(7).Left + Lbl(7).Width
        Lbl(20).Left = Lbl(19).Left + Lbl(19).Width
    End If
    End Sub

    '14/08/2002---------------------------------------Se descarga el formulario MovCaja
    Private Sub Form_Unload(Cancel As Integer)
    '-------------------------------------------------------------------------------------------------
    'Set WrkCaja = Nothing   'Destruye el espacio de trabajo
    Set Endoso = Nothing
    Set Caja = Nothing
    Set rstlocal(0) = Nothing
    Set rstlocal(1) = Nothing
    End Sub

Private Sub GridCheques_Click(Index As Integer)

Select Case Index
    
    Case 0  'LISTA DE CHEQUES EN TRANSITO
        
        If GridCheques(0).Text = "" Then Exit Sub
        If Txt(4) = "" And Dat(10) <> "PROVINCIAL" Then
            MsgBox "Primero Indique el Nº del Deposito", vbExclamation, App.ProductName
            Txt(4).SetFocus
            Exit Sub
        End If
        If GridCheques(1).Rows = 12 Then
            MsgBox "Imposible agregar más Cheques a este depósito", vbCritical, App.ProductName
            Exit Sub
        End If
        I = 1
        Label16(27) = CLng(Label16(27) - 1)
        With GridCheques(1)
        
1000        If .TextMatrix(I, 0) = "TOTAL: " Or .TextMatrix(I, 0) = "" Then
                .AddItem ("")
                For j = 0 To 2
                    .TextMatrix(I, j) = GridCheques(0).TextMatrix(GridCheques(0).RowSel, j)
                Next
                .Row = I
                .Col = 1
                .CellAlignment = flexAlignLeftCenter
                .MergeRow(I + 1) = True
                For j = 0 To 1
                    .Col = j
                    .TextMatrix(I + 1, j) = "TOTAL: "
                    .Row = I + 1: .CellAlignment = flexAlignRightCenter
                Next
                    CurTotalCheque = CurTotalCheque + CCur(.TextMatrix(I, 2))
                With GridCheques(0)
                    If .Rows > 3 Then
                        .RemoveItem (.RowSel)
                    Else
                        Call rtnLimpiar_Grid(GridCheques(0))
                    End If
                End With
                .TextMatrix(I + 1, 2) = _
                Format(CurTotalCheque, "#,##0.00 ")
                Label16(31) = Format(CurTotalCheque + Label16(30), "#,##0.00 ")
                .Row = I + 1: .Col = 2
            Else
                I = I + 1: GoTo 1000
            End If
        End With
        
    
    Case 1  'DETALLE DE DEPOSITO
    If GridCheques(1).ColSel = 3 Then
                
    Else
    If GridCheques(1).Text = "" Then Exit Sub
    I = 1
    Label16(27) = CLng(Label16(27) + 1)
        With GridCheques(0)

2000        If .TextMatrix(I, 0) = "" Then
                
                For j = 0 To 2
                    .TextMatrix(I, j) = GridCheques(1).TextMatrix(GridCheques(1).RowSel, j)
                Next
                CurTotalCheque = CurTotalCheque - CCur(.TextMatrix(I, 2))
                .AddItem ("")
                With GridCheques(1)
                    If .Rows > 4 Then
                        .RemoveItem (.RowSel)
                        .TextMatrix(.Rows - 2, 2) = Format(CurTotalCheque, "#,##0.00")
                        Label16(31) = Format(CurTotalCheque + Label16(30), "#,##0.00 ")
                    Else
                        Call rtnLimpiar_Grid(GridCheques(1))
                        CurTotalCheque = 0
                        GridCheques(1).RemoveItem (.Rows - 1)
                        Label16(31) = Label16(30)
                    End If
                End With
            Else
                I = I + 1: GoTo 2000
            End If
        End With
        End If
End Select

End Sub


Private Sub GridCheques_KeyPress(Index As Integer, KeyAscii As Integer)
'
If Index = 1 And GridCheques(1).Col = 3 Then
    '
    Call Validacion(KeyAscii, "0123456789")
    If KeyAscii > 26 Then GridCheques(1).Text = GridCheques(1).Text & Chr(KeyAscii)
    If KeyAscii = 8 Then    'backspace
        If Len(GridCheques(1).Text) > 0 Then _
        GridCheques(1).Text = Left(GridCheques(1).Text, Len(GridCheques(1).Text) - 1)
    End If
    '
End If
'
End Sub

Private Sub GridCheques_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
If Index = 1 And KeyCode = 46 And GridCheques(1).Col = 3 Then
    GridCheques(1).Text = ""
End If
End Sub

    Private Sub Label16_Change(Index As Integer)
    If Index = 0 Then   'total deducciones
        Command3(1).Enabled = False
        If IsNumeric(Label16(0)) Then
            If CCur(Label16(0)) > 0 Then Command3(1).Enabled = True
        End If
    End If
    End Sub

    '---------------------------------------------------------------------------------------------
    Private Sub Label16_Click(Index As Integer) '
    '---------------------------------------------------------------------------------------------
    '
    Select Case Index
    '
        Case 15 'ELIMINA CUALQUIER INFORMACION DEL TEMDEDUCCIONES
                
                Set cnnPropietario = New ADODB.Connection
                cnnPropietario.Open cnnOLEDB + StrRutaInmueble
                cnnPropietario.Execute "DELETE * FROM TemDeducciones WHERE IDPeriodos='" _
                & strPeriodo & "'"
                cnnPropietario.Close
                Set cnnPropietario = Nothing
                TimDemonio.Interval = 0
                frame3(3).Visible = False
                Command3(0).Enabled = True
                Call rtnTab(0)
                'Elimina la marca de la factura
                '-------------------------------------
                With FlexFacturas
                    .Row = intRow
                    .Col = 7
                    Set .CellPicture = Nothing
                End With
                  
        Case 29
            If Label16(29) = "" Then Label16(29) = 0
            If Label16(29) > 0 Then 'SI EL EFECTIVO ES SUPERIOR A CERO
                Label16(30) = Label16(29)
                Label16(31) = Format(CCur(Label16(29) + CCur(Label16(31))), "#,##0.00")
                Label16(29) = "0,00"
            End If
        
        Case 30
            If Label16(30) > 0 Then 'SI EL EFECTIVO ES SUPERIOR A CERO
                Label16(29) = Label16(30)
                Label16(31) = Format(CCur(Label16(31) - CCur(Label16(30))), "#,##0.00")
                Label16(30) = "0,00"
            End If
            
    End Select
    '
    End Sub
        
    Private Sub MntCalendar_DateClick(ByVal DateClicked As Date)
        MntCalendar.Visible = False
        MskFecha(B) = DateClicked: Txt(B).SetFocus
    End Sub
    
    Private Sub MntCalendar_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then MntCalendar.Visible = False: Txt(B).SetFocus
    
    If KeyAscii = 13 Then MskFecha(B) = MntCalendar.Value: MntCalendar.Visible = False: Txt(B).Enabled = True: Txt(B).SetFocus
    
    End Sub
    


Private Sub mnuGuardar_Click()
'guardar una operacion
Dim btnGuardar As ComctlLib.Button
Set btnGuardar = Toolbar1.Buttons("Save")
Call Toolbar1_ButtonClick(btnGuardar)
End Sub

Private Sub MskFecha_KeyPress(Index As Integer, KeyAscii As Integer)
If Index = 0 Then KeyAscii = 0: Exit Sub
Call Validacion(KeyAscii, "1234567890") 'SOLO PERMITE LA ENTRADA DE VALORES NUMERIOS
Select Case KeyAscii
    
    Case 8 'SI PRESIONA BACKSPACE ENTONCES
    
        Select Case Index
            
            Case 1, 2, 3
                
                If MskFecha(Index) = "" Then Cmb(Index + 1).SetFocus
        
        End Select


            
    Case 13 'SI PRESIONA ENTER ENTONCES

    Select Case Index
    
        Case 1, 2, 3
        
            If MskFecha(Index) <> "" Then Txt(Index).SetFocus
            
    End Select

End Select
End Sub



    'REV.27/08/2002---Opciones de Busqueda--------------------------------------------------------
    Private Sub OptBusca_Click(Index As Integer)
    '---------------------------------------------------------------------------------------------
    '
    Dim strSQL$
    ADOcontrol(3).CommandType = adCmdText
    Select Case Index   'Selecciona una opción
    '
        Case 0, 1, 2, 6 'Ordena la lista por Inmueble
    '   -----------------------------------------
            ADOcontrol(3).Recordset.Sort = OptBusca(Index).Tag
            Exit Sub
        
        Case 3, 4, 5, 7 'Filtra la caja por Inmueble
    '   -----------------------------------------
            If Index = 7 Then ADOcontrol(3).Recordset.Filter = "": TxtTot = ADOcontrol(3).Recordset.RecordCount: Exit Sub
            If TxtBus = "" Then
                MsgBox "Debe especificar en el campo 'Buscar' el codigo del condominio" _
                    & vbCrLf & "que vamos a filtrar la busqueda", vbExclamation, App.ProductName
                TxtBus.SetFocus: Exit Sub
            Else
                ADOcontrol(3).Recordset.Filter = OptBusca(Index).Tag & "='" & TxtBus & "'"
            End If
        
'        Case 5  'Filtra por un número de caja específico
'    '   -----------------------------------------
'            If TxtBus = "" Then
'                    MsgBox "Debe especificar en el campo 'Buscar' el codigo de caja" _
'                        & vbCrLf & "que vamos a filtrar la busqueda"
'                    TxtBus.SetFocus: Exit Sub
'            Else
'                    strSQl = " AND Inmueble.Caja='" & TxtBus _
'                        & "' ORDER BY Inmueble.Caja"
'            End If
        
'        Case 4  'filtra por una forma de pago específica
'    '   -----------------------------------------
'            If TxtBus = "" Then
'                    MsgBox "Debe especificar en el campo 'Buscar' la forma de pago" _
'                        & vbCrLf & "que vamos a filtrar la busqueda"
'                    TxtBus.SetFocus: Exit Sub
'            Else
'                    strSQl = " AND MovimientoCaja.FormaPagoMovimientoCaja='" & TxtBus _
'                        & "' ORDER BY MovimientoCaja.FormaPagoMovimientoCaja"
'            End If
                
    End Select
    '---------------------------------------------------------------------------------------------
'    ADOcontrol(3).RecordSource = "SELECT MovimientoCaja.*, inmueble.Caja,inmueble.Nombre " _
'                & "FROM inmueble INNER JOIN MovimientoCaja ON inmueble.CodInm = " _
'                & "MovimientoCaja.InmuebleMovimientoCaja WHERE " _
'                & "MovimientoCaja.FechaMovimientoCaja=Date() AND IDTaquilla=" _
'                & IntTaquilla & strSQl
'    ADOcontrol(3).Refresh
    TxtTot = ADOcontrol(3).Recordset.RecordCount
    End Sub


    '22/08/2002--------Rutina que se ejecuta al usuario hacer click sobre una ficha de la carpeta-
    Private Sub SSTab1_Click(PreviousTab As Integer)
    '---------------------------------------------------------------------------------------------
    'variables locales
    Dim strSQL As String
    '
    MntCalendar.Visible = False
    
    Select Case SSTab1.tab
        'Ficha general
        Case 0
            SSTab1.Height = 6480
            'If MntCalendar.Tag = 1 Then MntCalendar.Visible = True
    '   --------------------------
    '
        Case 1  'Lista de Transacciones
    '   -----------------------------------
            SSTab1.Height = 6480
            ADOcontrol(3).Recordset.Requery
    
        Case 2  'Ficha Deducciones
    '   -------------------------------
            SSTab1.Height = 6480
            FlexDeducciones.Rows = 2
            Command3(1).Enabled = True
            Command3(0).Enabled = True
            ADOcontrol(4).Recordset.Requery
            
        Case 3  'Lista de cheques recibidos
    '   -------------------------------------------
'            For I = 1 To 10: Toolbar1.Buttons(I).Enabled = False
'            Next
            For K = 0 To 1: Call centra_titulo(GridCheques(K), True)
            Next
    '       ---------------------------
            Dim RstCheque As New ADODB.Recordset
            Set objRst = New ADODB.Recordset
            'selecciona solo las cajas que tengan movimiento en la fecha
    '       ---------------------------
            strSQL = "SELECT DISTINCT CodigoCaja as Codigo, DescripCaja  as Descripcion FROM Ca" _
            & "ja INNER JOIN Inmueble ON Caja.codigoCaja=Inmueble.Caja WHERE CodInm IN (SELECT " _
            & "CodInmueble From TDFCheques WHERE FechaMov=Date() AND (Fpago='EFECTIVO' Or Fpago" _
            & "='CHEQUE') AND (IDDeposito='' or Isnull(IDDeposito)) and IDTaquilla=" & IntTaquilla & ") ORDER BY codigocaja;"
            
            objRst.Open strSQL, cnnConexion, adOpenKeyset, adLockOptimistic, adCmdText
            
  '         ---------------------------
            If Not objRst.EOF Then
                Set RstCheque = New ADODB.Recordset
                
                strSQL = "SELECT COUNT(IDTaquilla) From TDFCheques WHERE IDTaquilla=" & _
                IntTaquilla & " AND Fpago='CHEQUE' AND (IsNull(IDDeposito) or IDDeposito='') " _
                & "AND FechaMov=#" & Format(Date, "mm/dd/yyyy") & "#;"
                '
                RstCheque.Open strSQL, cnnConexion, adOpenKeyset, adLockOptimistic, adCmdText
                Label16(27) = RstCheque.Fields(0)
                
                Set Dat(7).RowSource = objRst
                Set Dat(8).RowSource = objRst
                
            End If
    '    ---------------------------
            
            SSTab1.Height = 7320
            DTPicker1.Value = Date
            
    '
    End Select
    '
    End Sub


Private Sub Text2_KeyPress(KeyAscii As Integer)
'Ver todas las cajas o una caja determinada
Dim strSQL As String
On Error GoTo Salir:
'
If KeyAscii = 13 Then

    strSQL = "SELECT MC.IDTaquilla, MC.InmuebleMovimientoCaja, I.Nombre, Mc.AptoMovimientoCaja," _
    & "I.Caja, MC.NumeroMovimientoCaja, MC.MontoMovimientoCaja, MC.FormaPagoMovimientoCaja, MC." _
    & "FechaMovimientoCaja, MC.IndiceMovimientoCaja FROM inmueble as I INNER JOIN MovimientoCaj" _
    & "a as MC ON I.CodInm = Mc.InmuebleMovimientoCaja Where (((MC.IDTaquilla) " & Text2 & ") A" _
    & "nd ((MC.FechaMovimientoCaja) =#" & Format(Text3, "mm/dd/yy") & "#)) ORDER BY MC.IndiceMo" _
    & "vimientoCaja DESC"
    ADOcontrol(3).RecordSource = strSQL
    '
    ADOcontrol(0).Recordset.Filter = ""
    If Text2 = "" Then
        ADOcontrol(0).Recordset.Filter = "FechaMovimientoCaja=#" & Text3 & "#"
    Else
        ADOcontrol(0).Recordset.Filter = "FechaMovimientoCaja=#" & Text3 & "# AND IDTaquilla" & Text2
    End If
    ADOcontrol(3).Refresh
    '
End If
Salir:
    If Err.Number <> 0 Then
        MsgBox Err.Description, vbCritical, "Error " & Err
    End If
'
End Sub

    Private Sub TimDemonio_Timer()
    '
    With AdoDeducciones
    
        With .Recordset
            .Requery
            '.MoveFirst
            .Find "Autoriza = -1"
            If .EOF Then Exit Sub
            MsgBox "Deducciones Autorizadas....", vbInformation, App.ProductName
            
            If .Fields("CodGasto") = strCodRebHA Then
                'If curPagoHono = 0 Then booHA = False
                LblHono = Format(0, "#,##0.00")
                Txt(5) = Format(CCur(Txt(5) + CCur(curPagoHono)), "#,##.000")
                Txt(11) = Format(CCur(Txt(10) - CCur(LblHono)), "#,##0.00")
            Else
                Txt(5) = Format(CCur(Txt(5)) - CCur(Label16(0)), "#,##0.00")
            End If
            '
            With FlexFacturas
            
                For I = 1 To .Rows - 1
                   If .TextMatrix(I, 0) = Trim(Label16(12)) Then _
                   .TextMatrix(I, 8) = Trim(Label16(0))
                Next I
                
            End With
            '
            Label16(17) = Format(CCur(Label16(17)) + CCur(Label16(0)), "#,##0.00")
            TimDemonio.Interval = 0
            frame3(3).Visible = False
            Call rtnTab(0)
            On Error Resume Next
            If Cmb(0).Enabled Then Cmb(0).SetFocus
            '
        End With
    '
    End With
    '
    End Sub
    
    '---------------------------------------------------------------------------------------------
    Private Sub Toolbar1_ButtonClick(ByVal Button As ComctlLib.Button)  '
    '---------------------------------------------------------------------------------------------
    '
    With ADOcontrol(0).Recordset
    '
        Select Case Button.Index
    '
            Case 1      'Primer Registro
                    If .RecordCount <= 0 Then Exit Sub
                    .MoveFirst
                    If Txt(12) <> "" And Txt(12) <> 0 Then
                    Else
                        Call RtnVisible("TRUE")
                        Call RtnVisible("FALSE")
                    End If
                    RtnAvanza
     '
            Case 2      'Registro Previo
     '
                    If .RecordCount <= 0 Then Exit Sub
                            .MovePrevious
                            'On Error Resume Next
                            If Not .BOF Then
                            If Txt(12) <> "" And Txt(12) <> 0 Then
                                Call RtnVisible("TRUE")
                            Else
                                Call RtnVisible("FALSE")
                            End If
                            RtnAvanza
                        Else
                            .MoveFirst
                            If Txt(12) <> "" And Txt(12) <> 0 Then
                                Call RtnVisible("TRUE")
                            Else
                                Call RtnVisible("FALSE")
                            End If
                        End If
                     
            
            Case 3      'Siguiente Registro
                
                If .RecordCount <= 0 Then Exit Sub
                    .MoveNext
                    If Not .EOF Then
                        On Error Resume Next
                        If Txt(12).Text <> "" And Txt(12) <> 0 Then
                            Call RtnVisible("TRUE")
                        Else
                            Call RtnVisible("FALSE")
                        End If
                        RtnAvanza
                    Else
                        .MoveLast
                        If Txt(12) <> "" And Txt(12) <> 0 Then
                            Call RtnVisible("TRUE")
                        Else
                            Call RtnVisible("FALSE")
                        End If
                    End If
    
            Case 4  'Último Registro
                
                If .RecordCount <= 0 Then Exit Sub
                .MoveLast
                If Txt(12) <> "" And Txt(12) <> 0 Then
                    Call RtnVisible("TRUE")
                Else
                    Call RtnVisible("FALSE")
                End If
                RtnAvanza
            
            Case 5 'agregar registro
            
                '-------------------
                'codigo temporal, para hacerle seguimiento a un error de
                'este modulo
                '-------------------------------------------------------
                If IntTaquilla = 0 Then
                    Dim strMsg As String
                    
                    strMsg = "El Sistema ha detectado un error en este módulo." & vbCrLf & _
                    "Cierre SAC e intenete nuevamente procesar esta transacción." & vbCrLf & _
                    "Si el problema persiste contacte vía telefónica al proveedor del software." & vbCrLf & _
                    "Se enviará una notificación vía mail al administrador del sistema."
                    
                    Call enviar_email
                    MsgBox strMsg, vbCritical, App.ProductName
                    Exit Sub
                    
                End If
                .AddNew
                MntCalendar.Tag = 0
                Call RtnEstado(5, Toolbar1, True)
                CHK.Value = vbUnchecked
                If Dir(App.Path & Archivo_Temp) <> "" Then Call Imprimir_Recibos
                Frame6.Enabled = True
                cnnConexion.BeginTrans  'comienza el proceso por lotes
                Call RtnVisible("FALSE")
                For I = 5 To 6: Dat(I).Enabled = False
                Next
                Dat(2).Enabled = True
                FmeCuentas.Enabled = True
                Dat(0).SetFocus
                DatProp.Enabled = False
                Cmb(0).Enabled = False
                Cmb(1).Enabled = False
                MskFecha(0) = Date
                Command1(0).Enabled = False
                SSTab1 = 0
                Command1(1).Enabled = False
                
                'blanquea controles no enlazados con el ADOcajadia
                Dat(1) = ""
                Dat(4) = ""
                Dat(3) = ""
                Dat(5) = ""
                LblHono = ""
                DatProp = ""
                Txt(10) = ""
                Txt(11) = ""
                Label16(17) = "0,00"
                Lbl(20) = ""
                Txt(13) = ""
                Call rtnLimpiar_Grid(FlexFacturas)
                '
                'actualiza las banderas
                mlEdit = False
                booAFuturo = False
                booACuenta = False
                booDed = False
                booIV = False
                booEV = False
                booHA = False
                booPC = False
                curPagoHono = 0
                
            Case 6  'Rev.12/08/2002 Guarda un Registro en MovimientoCaja
        '   Procesamiento por lotes, Utilizando la conexión pública a la BD
        '   --------------------------------------------------------------------
        '   esta rutina revierte toda la transacción
        '
            Dim ObjCheques As ADODB.Recordset
            Dim Fact As String
            '
            Screen.MousePointer = vbHourglass
            
            Call RtnEstado(6, Toolbar1, .EOF Or .BOF)
            
            For I = 0 To 3: MskFecha(I).PromptInclude = True
            Next
            'valida los datos mínimos necesarios para procesar una transacción
            '
            On Error GoTo rtnReversa    'Si ocurre algún error durante todo el proceso
            '
        '   Si está editando-----------------------------------------
            If mlEdit Then
            '       si se modifica el Tipo, Número, Banco, Fecha
                    strRecibo = !IDRecibo
                    Call Actualiza_FormaPago(strRecibo)
                    Call Guardar_FPago
                    '-
                    For I = 1 To 3
                        j = I - 1
                        If Txt(I) = "" Or IsNull(Txt(I)) Then Txt(I) = 0
                        .Fields("FPago" & IIf(I = 1, "", j)) = Cmb(4 + I)
                        .Fields("NumDocumentoMovimientoCaja" & IIf(I = 1, "", j)) = _
                        Txt(5 + I)
                        .Fields("BancoDocumentoMovimientoCaja" & IIf(I = 1, "", j)) _
                        = Cmb(1 + I)
                        If IsDate(MskFecha(I)) Then
                            .Fields("FechaChequeMovimientoCaja" & IIf(I = 1, "", j)) = _
                            MskFecha(I)
                        Else
                            .Fields("FechaChequeMovimientoCaja" & IIf(I = 1, "", j)) = Null
                            MskFecha(I).PromptInclude = False
                        End If
                       .Fields("MontoCheque" & IIf(I = 1, "", j)) = CCur(Txt(I))
                    Next I
                    '
                    Call rtnBitacora("Actualizar Transaccion #" & strRecibo)
                    .Update
                    For I = 1 To 3: MskFecha(I).PromptInclude = True
                    Next
                    Call rtnEditar(False)
                    MsgBox "Cambios Actualizados", vbInformation, App.ProductName
                    ADOcontrol(3).Recordset.Requery
                    mlEdit = False
                    Screen.MousePointer = vbDefault
                    Exit Sub
                    '
            End If
        '   ---------------------------------------------------------
            Toolbar1.Enabled = False
            If ftnValidar Then
                Screen.MousePointer = vbDefault
                Call RtnEstado(5, Toolbar1, .EOF Or .BOF)
                Toolbar1.Enabled = True
                Exit Sub
            End If
            '
            Command1(1).Caption = "&Distribuir"
            'strRecibo = Right(Dat(0), 2) & Dat(2) & Format(Date, "ddmmyy") & Format(txt(0), "00")
        '   Resume la descripción es un pago de condominio y la longitud es mayor a 70 caracteres
            If Len(Txt(9)) > 70 And Dat(5) = strCodPC Then Call RtnDescripcion
        '   Si es un egreso de caja el monto lo convierte en negativo e imprime recibo de egreso
            If Cmb(1).ListIndex = 0 Then
                Txt(5) = Format(CCur(Txt(5) * -1), "#,##0.00")
                booEV = False
            End If
        '   -------------------------------------------------------------------------------------
        '   Guarda el registro en la Tabla MovimientoCaja
        '   -------------------------------------------------------------------------------------
            .Fields("Usuario") = gcUsuario
            .Fields("Freg") = Date
            .Fields("IDTaquilla") = IntTaquilla
            .Fields("IdRecibo") = strRecibo
            .Fields("Hora") = Format(Time, "hh:mm ampm")
            
            If Not IsNumeric(Txt(12)) Then Txt(12) = 0
            If Not IsNumeric(Txt(1)) Then Txt(1) = 0
            If Not IsNumeric(Txt(2)) Then Txt(2) = 0
            If Not IsNumeric(Txt(3)) Then Txt(3) = 0
            '
            '-->montos en bolivares 2007
            .Fields("MontoMovimientoCajaBs") = CCur(Txt(5))
            .Fields("EfectivoMovimientoCajaBs") = CCur(Txt(12))
            .Fields("MontoChequeBs") = CCur(Txt(1))
            .Fields("MontoChequeBs1") = CCur(Txt(2))
            .Fields("MontoChequeBs2") = CCur(Txt(3))
            '
            '-->
            '*********************************
            'If Not IsNumeric(Txt(12)) Then Txt(12) = 0
            .Update
            Call rtnBitacora("Guardar Transaccion Caja " & IntTaquilla & " Op." & .Fields("IDRecibo"))
            FmeCuentas.Enabled = False
            'si la transacción afecta una cuenta de fondo deja el registro
            '
            Call Guardar_FPago
            
            If booFondo(Dat(5), StrRutaInmueble) Then
                '
                cnnConexion.Execute "INSERT INTO MovFondo(CodGasto,Fecha,Tipo," & _
                "Periodo,Concepto,Debe,Haber) IN '" & StrRutaInmueble & _
                "' VALUES('" & Dat(5) & "',Date(),'" & _
                IIf(Cmb(1) = "EGRESO", "ND", "NC") & "','01/" & _
                Format(Date, "mm/yyyy") & "','" & Txt(9) & "','" & _
                IIf(Cmb(1) = "EGRESO", convertirBsF(CDbl(Txt(5))), 0) & _
                "','" & IIf(Cmb(1) = _
                "EGRESO", 0, convertirBsF(CDbl(Txt(5)))) & "');"
                '
            
            End If
        '----------------------------------------
            
        '   --------------------------------------------------------------------------------------
        '   Si la op. genera intereses moratorios se inserta el registo en TDFperiodos
        '   --------------------------------------------------------------------------------------
        '
            If curHono <> 0 And Not IsNull(curHono) Then
        '
                curHono = convertirBsF(curHono)
                
                cnnConexion.Execute "INSERT INTO Periodos (IDRecibo,IdPeriodos," & _
                "CodGasto,Descripcion,Monto) VALUES ('" & strRecibo & "','" & _
                strRecibo & "','" & strCodHA & "','HONORARIOS DE ABOGADO','" _
                & curHono & "')"
                curHono = 0
                strPeriodo = strRecibo
                Call RtnDeducciones(strRecibo)
                booHA = IIf(curPagoHono = 0, False, True)
                'el cliente pagó honorarios y además tiene convenio
'                If curPagoHono > 0 And booCon Then
'                    Call Regristra_Pago
'                End If
        '
            End If
        '   -------------------------------------------------------------------------------------
        '   Genera un registro en TDFperiodos correspondiente a c/factura cancelada
        '   -------------------------------------------------------------------------------------
            If Dat(5) = strCodPC Or Dat(5) = strCodAbonoCta Or Dat(5) = strCodRCheq Then
        '
                Dim CurFs@, CurRecibo@, CurFac@, strCodigo$, StrDetalle$, fecha$
                Dim apto$, Mes$, curDed@
                '
                CurFs = 0
                CurRecibo = 0
        '
                For I = 1 To FlexFacturas.Rows - 1
                    If FlexFacturas.TextMatrix(I, 6) = "SI" Then
        '               Asigna valores a variables------------------------------------------------
                        strCodigo = IIf(FlexFacturas.TextMatrix(I, 4) = 0, _
                            IIf(FlexFacturas.TextMatrix(I, 0) Like "CH*", _
                            strCodRCheq, strCodPC), strCodAbonoCta)
                        
                        
                        StrDetalle = IIf(strCodigo = strCodPC, "PAGO CONDOMINIO", IIf(strCodigo _
                        = strCodAbonoCta, "ABONO A CUENTA", "REP. CHEQUE DEV."))
                        
                        CurFac = IIf(FlexFacturas.TextMatrix(I, 4) <> 0, _
                            CCur(FlexFacturas.TextMatrix(I, 3)) _
                            - CCur(vecFPS(I, 1)), vecFPS(I, 2))
                        
                        CurFac = convertirBsF(CurFac)
                        
                        strPeriodo = strRecibo & Left(FlexFacturas.TextMatrix(I, 1), 2) & Right(FlexFacturas.TextMatrix(I, 1), 2)
                        
                        If strCodigo = strCodPC And IsNumeric(FlexFacturas.TextMatrix(I, 0)) Then
                            '
                            booPC = True
                            If IsNumeric(FlexFacturas.TextMatrix(I, 8)) Then
                                curDed = convertirBsF( _
                                    CDbl(FlexFacturas.TextMatrix(I, 8)) _
                                )
                            Else
                                curDed = 0
                            End If
                            
                            Call Guardar_NumFact(FlexFacturas.TextMatrix(I, 0), _
                                CurFac - curDed)
                            '-------
                            
                        ElseIf strCodigo = strCodAbonoCta And IsNumeric(FlexFacturas.TextMatrix(I, 0)) Then
                            
                            Fact = FlexFacturas.TextMatrix(I, 0)
                            apto = Dat(2).Text
                            Mes = "01-" & FlexFacturas.TextMatrix(I, 1)
                            
                            cnnConexion.Execute "INSERT INTO DetFact(Fact,Det" & _
                            "alle,Codigo,CodGasto,Periodo,Monto,Fecha,Hora," & _
                            "Usuario) IN '" & StrRutaInmueble & _
                            "' VALUES('" & Fact & "','" & StrDetalle & "','" & _
                            apto & "','" & strCodigo & "','" & Mes & "','" & _
                            CurFac * -1 & "',Date(),Format(Ti" _
                            & "me(),'hh:mm:ss'),'" & gcUsuario & "');"
                            
                        End If
                        
        '               Si es Rep. Cheq. Dev. Actualiza la tabla Cheque Devuelto
                        If strCodigo = strCodRCheq Then
                            Fact = FlexFacturas.TextMatrix(I, 0)
                            
                            cnnConexion.Execute "UPDATE ChequeDevuelto IN '" & _
                            gcPath & "\" & Dat(0) & "\" & "inm.mdb' SET " & _
                            "Recuperado=True WHERE Codigo='" & _
                            Dat(2) & "' AND NumCheque = '" & _
                            Right(FlexFacturas.TextMatrix(FlexFacturas.RowSel, 0), Len(FlexFacturas.TextMatrix(FlexFacturas.RowSel, 0)) - 4) & "'"
                            Call rtnBitacora("Cheque #" & Fact & " Recuperado")
                            Rem------------------------------------+
                            'Elimina el registro en sac.mdb
                            'cnnConexion.Execute "DELETE * ChequeDevuelto WHERE CodInm='" & _
                            Dat(0) & "' AND Apto='" & Dat(2) & "' AND Numero ='" & _
                            Right(FlexFacturas.TextMatrix(FlexFacturas.RowSel, 0), 6) & "';"
                            Rem-----------------------------------+
                            'Actualiza la tabla 'factura'
                            cnnConexion.Execute "UPDATE Factura IN '" & _
                            StrRutaInmueble & "' SET " & "Pagado = Pagado + '" & _
                            CurFac & "', Saldo=Saldo - '" & _
                            CurFac & "', freg= Date(), usuario='" & gcUsuario & "', fecha=Forma" _
                            & "t(Time(),'hh:mm:ss') WHERE CodProp = '" & Dat(2) & "' AND FACT= '" & Fact & "'"
                           
                        End If
        '              Inserta registros Tabla "periodos"----------------------------------------
                        Fact = strRecibo & Left(FlexFacturas.TextMatrix(I, 1), 2) & _
                                Right(FlexFacturas.TextMatrix(I, 1), 2)
                                
                        cnnConexion.Execute "INSERT INTO Periodos (IDRecibo," & _
                        "IDPeriodos,Periodo,CodGasto,Descripcion,Monto," & _
                        "Facturado) VALUES ('" & strRecibo & "','" _
                        & Fact & "','" & FlexFacturas.TextMatrix(I, 1) & _
                        "','" & strCodigo & "','" & StrDetalle & "','" & _
                        CurFac & "','" & convertirBsF(CDbl(vecFPS(I, 0))) & "')"
                        
                        If FlexFacturas.TextMatrix(I, 0) <> "" Then Fact = FlexFacturas.TextMatrix(I, 0)
                        
                        '
        '            Busca las deducciones por periodo
                    Call RtnDeducciones(Fact)
        '            Actualiza la deuda del propietario en Tabla "Facturas"---------------------
                        If FlexFacturas.TextMatrix(I, 6) = "SI" And strCodigo <> strCodRCheq Then
        '
                            fecha = Format(1 & "-" & FlexFacturas.TextMatrix(I, 1), "mm/dd/yy")
                            cnnConexion.Execute "UPDATE Factura IN '" & StrRutaInmueble & "' SE" _
                            & "T " & "Pagado = Pagado + '" & CurFac & "',Saldo = Saldo - '" & _
                            CurFac & "', freg=Date(), usuario='" & gcUsuario & "', fecha=Format(Time(),'hh:mm:ss')" _
                            & " WHERE CodProp = '" & Dat(2) & "' AND Periodo = #" & fecha & "#"
                            
                        End If
                        CurFs = CurFs + CurFac                              'Acumulador
                        CurRecibo = IIf(strCodigo = strCodPC, (CurRecibo + 1), CurRecibo)
        '              ------------------------------------'Marca la impresión de abono a cuenta
                        If strCodigo = strCodAbonoCta Then booACuenta = True
                        
                    End If
        '
                Next    'fin del bulce
                
                If IntMonto > 0 Then
                    Dim strParam$
                    strParam = Format(Left(FlexFacturas.TextMatrix((I - 1), 1), 2) + 1, "00") _
                    & "-" & Right(FlexFacturas.TextMatrix((I - 1), 1), 2)
                    If Left(strParam, 2) = 13 Then
                        strParam = "01-" & Format(Right(strParam, 2) + 1, "00")
                        End If
                    Call RtnAbonoFut(IntMonto, strParam)
                    CurFs = CurFs + IntMonto
                End If
                
        '      Actualiza la deuda general del inmueble
                cnnConexion.Execute "UPDATE Inmueble SET Deuda = Deuda - '" & CurFs & "' WHERE " _
                & "CodInm = '" & Dat(0) & "'"
        '      Actualiza la deuda general del propietario
                cnnConexion.Execute "UPDATE Propietarios IN '" & StrRutaInmueble & "' SET Deuda" _
                & "=Deuda - '" & CurFs & "', Recibos = Recibos - " & CurRecibo & ", UltPago = '" _
                & Txt(5) & "', FecUltPag ='" & Date & "', FecReg='" & Date & "', Usuario ='" & _
                gcUsuario & "', Notas = '" & Txt(13) & "' WHERE codigo = '" & Dat(2) & "'"
        '       Registra el pago en la tabla convenio (si el propietario tiene uno)
'                If booCon Then Call Regristra_Pago
                
            ElseIf Dat(5) = strCodAbonoFut Then 'Es un Abono A Fututo
            
                strParam = Format(Date + 1, "mm-YY")
                CurFs = convertirBsF(CDbl(Txt(5)))
                Call RtnAbonoFut(CurFs, strParam)
        '      Actualiza la deuda general del inmueble
                cnnConexion.Execute "UPDATE Inmueble SET Deuda = Deuda - '" & CurFs _
                & "' WHERE CodInm = '" & Dat(0) & "'"
        '      Actualiza la deuda general del propietario
                cnnConexion.Execute "UPDATE Propietarios IN '" & StrRutaInmueble & "' SET Deuda" _
                & "= Deuda - '" & CurFs & "', Recibos = Recibos - " & CurRecibo & ", UltPago='" _
                & Txt(5) & "', FecUltPag ='" & Date & "', FecReg='" & Date & "', Usuario ='" & _
                gcUsuario & "', Notas = '" & Txt(13) & "' WHERE codigo = '" & Dat(2) & "'"
            
            ElseIf Dat(5) = strCodIV Then   'ingresos varios
                booIV = True
            '
            ElseIf Dat(5) = strCodCCHeq Then 'Cambio de cheques
                CurEfectivo = CCur(Txt(5)) * -1
                CurEfectivo = convertirBsF(CurEfectivo)
                
                cnnConexion.Execute "INSERT INTO tdfCHEQUES (IDRecibo,IdTaquilla," & _
                "CodInmueble, FechaMov,Fpago, Monto, Ndoc, Banco, MontoBs ) " & _
                "Values ('" & strRecibo & "'," & IntTaquilla & ",'" & Dat(0) & _
                "', DATE(),'EFECTIVO','" & CurEfectivo & "','" & _
                strRecibo & "','Ccheq')"
            '
            ElseIf Dat(5) = strCodHA Then   'Cobro Honorarios de Abogado
                Call RtnDeducciones(strRecibo)
            End If
        '   -----------------------------------------------------------------------------------------
            'Call RtnDeducciones(strRecibo)
rtnReversa:

            If Err.Number = 0 Then  'La Transacción fué procesada con éxito
                
                cnnConexion.CommitTrans
                '-->campos convertidos a Bs. fuertes
                .Update "MontoMovimientoCaja", convertirBsF(.Fields("MontoMovimientoCajaBs"))
                .Update "EfectivoMovimientoCaja", convertirBsF(.Fields("EfectivoMovimientoCajaBs"))
                .Update "MontoCheque", convertirBsF(.Fields("MontoChequeBs"))
                .Update "MontoCheque1", convertirBsF(.Fields("MontoChequeBs1"))
                .Update "MontoCheque2", convertirBsF(.Fields("MontoChequeBs2"))
                '
                
                strQ = "SELECT MC.IDTaquilla, MC.IDRecibo, MC.InmuebleMovimientoCaja, I.Nombre, I.Caja," _
                & "C.DescripCaja , MC.AptoMovimientoCaja, MC.CuentaMovimientoCaja, MC.DescripcionMovimiento" _
                & "Caja, MC.MontoMovimientoCaja, P.CodGasto, P.Descripcion, P.Periodo, P.Monto, D.CodGasto," _
                & "D.Titulo, D.Monto, MC.FechaMovimientoCaja, MC.NumDocumentoMovimientoCaja, MC.NumDocument" _
                & "oMovimientoCaja1, MC.NumDocumentoMovimientoCaja2, MC.BancoDocumentoMovimientoCaja, MC.Ba" _
                & "ncoDocumentoMovimientoCaja1, MC.BancoDocumentoMovimientoCaja2, MC.FechaChequeMovimientoC" _
                & "aja, MC.FechaChequeMovimientoCaja1, MC.FechaChequeMovimientoCaja2, MC.FormaPagoMovimient" _
                & "oCaja, MC.MontoCheque, MC.MontoCheque1, MC.MontoCheque2, MC.EfectivoMovimientoCaja , MC." _
                & "FPago, MC.FPago1, MC.FPago2, 0, I.FondoAct, P.Facturado, MC.CodGasto FROM (((Caja as C INNER " _
                & "JOIN inmueble as I ON C.CodigoCaja = I.Caja) INNER JOIN MovimientoCaja as MC ON I.CodInm" _
                & "= MC.InmuebleMovimientoCaja) LEFT JOIN Periodos as P ON MC.IDRecibo = P.IDRecibo) LEFT J" _
                & "OIN Deducciones as D ON P.IDPeriodos = D.IDPeriodos WHERE MC.FechaMovimientoCaja=DATE() " _
                & "AND MC.IdTaquilla=" & IntTaquilla & " AND MC.IDRecibo='" & strRecibo & "' ORDER BY I.Ca" _
                & "ja, MC.FormaPagoMovimientoCaja, MC.InmuebleMovimientoCaja , MC.AptoMovimientoCaja;"
    
                Dim qdfCreada As Boolean
            
        '       -----------------------------'Si se procesó un abono a futuro imprime el reporte
                If booAFuturo = True Then
                
                    Call rtnGenerator(gcPath & "\sac.mdb", strQ, "QdfCaja")
                    qdfCreada = True
                    
                    Call RtnImpresion(strRecibo, " AND {QDFCaja.P.CodGasto}='" _
                        & strCodAbonoFut & "'", "ABONO PROX. FACTURACION", StrRutaInmueble)
                        
                End If
        '        -----------------------------'Imprime Abono a cuenta si los hay
                If booACuenta = True Then
                
                    If Not qdfCreada Then
                        Call rtnGenerator(gcPath & "\sac.mdb", strQ, "QdfCaja")
                        qdfCreada = True
                    End If
                    
                    Call RtnImpresion(strRecibo, " AND {QDFCaja.P.CodGasto}='" _
                        & strCodAbonoCta & "'", "ABONO A CUENTA", StrRutaInmueble)
                End If
                        
                If booDed = True Then
                
                    If Not qdfCreada Then
                        Call rtnGenerator(gcPath & "\sac.mdb", strQ, "QdfCaja")
                        qdfCreada = True
                    End If
                    If Respuesta("¿Desea el recibo de deducciones?") Then   'Imprime rec.
                        Call RtnImpresion(strRecibo, " AND {QDFcaja.D.Monto}<>0", _
                        "DEDUCCIONES", StrRutaInmueble)
                    End If
                    
                End If
        '        -----------------------------'Imprime recibo de ingresos varios
                If booIV = True Then
                    If Not qdfCreada Then
                        Call rtnGenerator(gcPath & "\sac.mdb", strQ, "QdfCaja")
                        qdfCreada = True
                    End If
                    Call RtnImpresion(strRecibo, "", "INGRESOS VARIOS", StrRutaInmueble)
                End If
        '        -----------------------------'Imprime recibo de egresos varios
                If booEV = True Then
                    If Not qdfCreada Then
                        Call rtnGenerator(gcPath & "\sac.mdb", strQ, "QdfCaja")
                        qdfCreada = True
                    End If
                    Call RtnImpresion(strRecibo, "", "EGRESOS VARIOS", StrRutaInmueble)
                End If
        '      -----------------------------'Imprime recibo de cobro de Honorarios
                If booHA Then
                    'si se cobran honorarios genera un registro en la tabla honorarios
                    cnnConexion.Execute "INSERT INTO Honorarios (IDRecibo,IDAbogado,IDEstatus) V" _
                    & "ALUES ('" & strRecibo & "',0,0)"
                    '
                    If Not qdfCreada Then Call rtnGenerator(gcPath & "\sac.mdb", strQ, "QdfCaja")
                    
                    If Dat(3) <> sysCodCaja Then Call Cpp_Honorarios
                    Call RtnImpresion(strRecibo, " AND {QDFCaja.P.CodGasto}='" & _
                    strCodHA & "'", "HONORARIOS DE ABOGADO", StrRutaInmueble)
                    
                End If
                '
                'On Error Resume Next
                If booPC = True Then Call Imprimir_Recibos
                Call rtnLimpiar_Grid(FlexFacturas)
                Call rtnBitacora("Transacción procesada con éxito")
                Fact = "Transacción Procesada con Éxito..."
                '      -----------------------------'Imprime recibos de pago
            
        
            Else                                'La trans. originó errores en uno de sus pasos
            
                cnnConexion.RollbackTrans
                cnnConexion.Execute "DELETE * FROM MovimientoCaja WHERE IDRecibo='" _
                & strRecibo & "'"
                Call rtnLimpiar_Grid(FlexFacturas)
                Fact = "La Operación no fué llevada a cabo éxito, inténtelo nuevamente," _
                & vbCrLf & "si el problema persiste consulte al administrador del sistema" _
                & vbCrLf & Err.Description & vbCrLf & Err.Source
                Err.Clear
                Call rtnBitacora("Error: " & Err.Description)
        '
            End If
        '   --------------------------------------------------------------------------------------
            On Error Resume Next
            cnnConexion.Execute "DELETE * FROM TemDeducciones In '" & StrRutaInmueble & "' WHER" _
            & "E Taquilla =" & IntTaquilla
        '   --------------------------------------------------------------------------------------
            
            For I = 0 To 3: MskFecha(I).PromptInclude = False
            Next
            '
            'If Not ADOcontrol(0).Recordset.EOF And Not ADOcontrol(0).Recordset.BOF Then _
            ADOcontrol(0).Recordset.MoveLast
            Frame1.Enabled = False
            Frame5.Enabled = False
            Command1(0).Enabled = True
            Command1(1).Enabled = False
        '   Actualiza banderas
            mlEdit = False
            Call Desmarca
            Call mostrar_area(3)
            Screen.MousePointer = vbDefault
            Toolbar1.Enabled = True
            Frame6.Enabled = True
            
            If Err <> 0 Then
                Call rtnBitacora("Error " & Err & " " & Err.Description)
                MsgBox Fact, vbCritical, App.ProductName
            Else
                
                SetTimer hWnd, NV_CLOSEMSGBOX, 2000&, AddressOf TimerProc
                Call MessageBox(hWnd, Fact, App.ProductName, vbInformation)
                
            End If
            
            
    '           ---------------------------------------------------------------Fin Guarda Registro
            Case 7   'Buscar Registro
    '       -------------------------------
                If SSTab1.TabEnabled(1) Then
                    SSTab1.tab = 1
                    FrmBusca.Enabled = True
                    FrmBusca1.Enabled = True
                    TxtBus.SetFocus
                Else
                    MsgBox "Opción disponible luego de ser autorizado el cierre de caja", _
                    vbInformation, App.ProductName
                End If
    
            Case 8      'CANCELAR REGISTRO
    '       ------------------------------------------
        '       ELIMINA CUALQUIER REGISTRO DEL TEMDEDUCCIONES
                If StrRutaInmueble <> "" Then
                    Set cnnPropietario = New ADODB.Connection
                    cnnPropietario.Open cnnOLEDB + StrRutaInmueble
                    cnnPropietario.Execute "DELETE * FROM TemDeducciones WHERE Taquilla =" & _
                    IntTaquilla
                    cnnPropietario.Close: Set cnnPropietario = Nothing
                End If
        '
                If Command1(1).Caption = "&Deshacer" Then MsgBox "Debe Deshacer Distribuir Para" _
                & " Canelar", vbInformation, App.ProductName: Command1(1).SetFocus: Exit Sub
                For I = 0 To 3: MskFecha(I).PromptInclude = True
                Next
                .CancelUpdate
                Call mostrar_area(3)
                Frame6.Enabled = True
                Frame1.Enabled = False
                
                If Not mlEdit Then cnnConexion.RollbackTrans
                Call rtnEditar(False)
                Call RtnEstado(8, Toolbar1, .EOF Or .BOF)

                If Not .EOF Then .MoveFirst
                Call RtnVisible("FALSE")
                For I = 0 To 3: MskFecha(I).PromptInclude = False
                Next
                Call rtnLimpiar_Grid(FlexFacturas)
                Call Desmarca
                mlEdit = False
                Command1(0).Enabled = True: Command1(1).Enabled = True
                Label16(17) = "0,00"
                SSTab1.tab = 0
                SSTab1.TabEnabled(0) = True
                SSTab1.TabEnabled(2) = False
                
                If Dat(0) = "" Or IsNull(Dat(0)) Then
                
                        Txt(11) = ""
                        Txt(10) = ""
                        Dat(1) = ""
                        DatProp = ""
                        Dat(3) = ""
                        Dat(4) = ""
                        MskFecha(Index) = "__/__/____"
                        LblHono = "0,00"
                        Label16(17) = "0,00"
                        SSTab1.TabEnabled(3) = Toolbar1.Buttons("New").Enabled
                        Exit Sub
                        
                End If
                '
                Call BuscaInmueble("Inmueble.CodInm", Dat(0))
                '
                Set objRst = ObjCmd.Execute             'referente al inmueble seleccionado
                '
                With objRst
                
                If objRst.EOF Then Exit Sub
                    '
                    'Dat(0) = .Fields("CodInm")            'Codigo del Inmueble
                    Dat(1) = .Fields("Nombre")            'nombre del inmueble
                    Dat(3) = .Fields("Caja")                'caja del inmueble
                    StrRutaInmueble = gcPath & .Fields("Ubica") & "inm.mdb"
                    Dat(4) = .Fields("Descripcaja")         'descripcion de la caja
                    '
                End With
                objRst.Close
                Set objRst = Nothing
                Call RtnQdfPropietario(Dat(2))
                'Call RtnBuscaIngreso("Titulo", Dat(6), "CodGasto", Dat(5))
                '
            Case 9      'ELIMINAR REGISTRO
            '-----------------------------------------
            '
                If blnCaja Or gcNivel <= nuSUPERVISOR Then
                    
                    Call RtnEstado(9, Toolbar1, .EOF Or .BOF)
                    If Respuesta("Esta a punto de eliminar la trans. " & Dat(6) & vbCrLf & "Apto: " & .Fields("AptoMovimientoCaja") & " del Inmueble: " & .Fields("InmuebleMovimientoCaja") & ". ¿Desea Continuar?") Then
                        MousePointer = vbHourglass
                        Call rtnBitacora("Caja:Eliminando Registro " & !IDRecibo)
                        Call RtnConfigUtility(True, "Eliminando Registro '", _
                            "Iniciando Proceso", "Inmueble=" & Dat(0) & vbCrLf & "Apto:" _
                            & Dat(2) & vbCrLf & "Descrip.:" & Dat(5) & " " & Dat(6))
                        Call RtnEliminarMovCaja(!IDRecibo, Dat(5))
                        MousePointer = vbDefault
                        If Err.Number = 0 Then
                            ADOcontrol(0).Recordset.Requery
                        
                            If ADOcontrol(0).Recordset.RecordCount > 0 Then _
                            ADOcontrol(0).Recordset.MoveFirst
                            ADOcontrol(3).Refresh
                        End If
                        
                    End If
                    '
                    'actualiza las banderas
                    mlEdit = False
                Else
                    MsgBox "Diríjase a su supervisor" & _
                    " para eliminar esta transacción...", vbInformation, App.ProductName
                    Call rtnBitacora("Solicitud Eliminar Trans.#" & !IDRecibo)
                    'si dispone de RealpoPup lo utiliza como medio para solicitar la autorización
                    'a un supervisor
                    If Dir("C:\Archivos de programa\RealPopup\RealPopup.exe") <> "" Then
                        strMsg = "Eliminar transacción #" & !IDRecibo
                        m = Shell("C:\Archivos de programa\RealPopup\RealPopup -send archivo " & _
                        Chr(34) & strMsg & Chr(34) & " -NOACTIVATE")
                    End If
                    '
                End If
            '
            Case 10     'EDITAR REGISTRO
                Frame6.Enabled = False
                Call RtnEstado(10, Toolbar1, True)
                Call rtnEditar(True)
                mlEdit = True
            '
            Case 12     'CERRAR FORMULARIO
                On Error Resume Next
                FrmMovCajaBs.Hide
                '
            Case 11    'IMPRIMIR REGISTRO
            '       -------------------------------------
                Dim rstCuadre As ADODB.Recordset
                Dim pReport As ctlReport
                '
                
                If SSTab1.tab = 3 Then
                    Set pReport = New ctlReport
                    pReport.Reporte = gcReport & "cajadepositos.rpt"
                    pReport.FormuladeSeleccion = "{TDFCheques.IDTaquilla}=" & IntTaquilla & " and {TDFCheques.FechaMov}=CDate (" & Format(Date, "yyyy,mm,dd") & ")"
                    pReport.TituloVentana = "Depósitos Efectuados Hoy"
                    pReport.Salida = crPantalla
                    pReport.Imprimir
                    Set pReport = Nothing
                Else
                    Set rstCuadre = New ADODB.Recordset
                    '
                    rstCuadre.Open "SELECT * FROM Taquillas WHERE IDTaquilla=" & IntTaquilla, _
                    cnnConexion, adOpenKeyset, adLockOptimistic
                    '
                    If rstCuadre!Cuadre Or gcNivel <= nuADSYS Then
                    '
                        FrmReport.FraCaja.Visible = True
                        FrmReport.Frame1.Visible = True
                        FrmReport.Hora = Format(rstCuadre("Hora"), "hh:mm")
                        mcTitulo = "Cuacre de Caja"
                        mcReport = "CajaReport.rpt"
                        mcDatos = gcPath + "\sac.mdb"
                        mcOrdCod = ""
                        mcOrdAlfa = ""
                        mcCrit = ""
                        FrmReport.Show 1
                        mcDatos = gcPath + gcUbica + "Inm.mdb"
                    Else
                        MsgBox "No está autorizado, consulte con el Supervisor", vbInformation, _
                        App.ProductName
                    End If
                    rstCuadre.Close
                    Set rstCuadre = Nothing
                End If
                '
            End Select
    '
    End With
    '
    SSTab1.TabEnabled(3) = Toolbar1.Buttons("New").Enabled
    '
    End Sub
    
    
    '---------------------------------------------------------------------------------------------
    Sub BuscaPropietario(StrCampo$, StrControl$, StrRecord$, StrControl1 As DataCombo)
    '---------------------------------------------------------------------------------------------
    'rutina que busca el nombre o el codigo de un propietario
    'segun parametros enviados desde el control que llama la rutina
    
    If StrControl = "" Then Exit Sub
    FlexFacturas.Rows = 2
    Call rtnLimpiar_Grid(FlexFacturas)
    Txt(13) = ""
    With ADOcontrol(2).Recordset
        '
            If .EOF And .BOF Then Exit Sub
            .MoveFirst
            .Find StrCampo & " LIKE '*" & StrControl & "*'"
            
            If .EOF Then
                MsgBox "No Tengo Registrado ese Propietario " & "'" & StrControl & "'", _
                vbInformation, App.ProductName
                Txt(10) = "0,00"
                Txt(11) = "0,00"
                If StrCampo = "Codigo" Then
                    Dat(2).SetFocus
                Else
                    DatProp.SetFocus
                End If
                Exit Sub
            End If
            StrControl1.Text = .Fields(StrRecord)
            Dat(2) = .Fields("Codigo")
            DatProp = .Fields("Nombre")
            Txt(13) = IIf(IsNull(.Fields("Notas")) Or .Fields("Notas") = "", "", .Fields("Notas"))
            Txt(10) = Format(.Fields("deuda"), "#,##0.00")
            LblHono = Format(0, "#,##0.00")
            Txt(11) = Txt(10)
            'numero de recibo ó operación
            strRecibo = Right(Dat(0), 2) & Dat(2) & Format(Date, "ddmmyy") & Format(Txt(0), "00")
            If !Convenio Then
                frmConsultaCon.Show vbModal, FrmAdmin
                booCon = True
            Else
                booCon = False
            End If
    '
    End With
    '
    End Sub
    
    '---------------------------------------------------------------------------------------------
    '   Rutina: RtnListaPro
    '
    '   Agrega a la lista del control combo los codigos y nombres de los
    '   propietarios de derminado condominio
    '---------------------------------------------------------------------------------------------
    Sub RtnListapro()
        
    If Dir(StrRutaInmueble) = "" Then MsgBox "Informacion de Propietarios" _
    & " no existe" & Chr(13) & "Consulte al Soporte Tecnico", vbCritical, App.ProductName: Exit Sub
    If cnnPropietario.State = 1 Then
        cnnPropietario.Close
    End If
    cnnPropietario.ConnectionString = cnnOLEDB + StrRutaInmueble
    cnnPropietario.Open
    ADOcontrol(2).ConnectionString = cnnPropietario
    ADOcontrol(2).Refresh
    '
    End Sub
    
    Sub RtnVisible(valor As String)
        Lbl(8).Visible = valor
        Txt(12).Visible = valor
    End Sub
    
    
    '---------------------------------------------------------------------------------------------
    '   Rutina:     rtnBuscaIngreso
    '
    '   Busca el concepto/código de una cuenta de gasto determinada
    '---------------------------------------------------------------------------------------------
    Private Sub RtnBuscaIngreso(StrCampo$, StrControl$, StrRecord$, StrControl1 As DataCombo)
    '
    With ADOcontrol(1).Recordset
        '.MoveFirst
        .Find StrCampo & " = '" & Trim(StrControl) & "'"
        If .EOF Then
             .MoveFirst
            .Find StrCampo & " = '" & Trim(StrControl) & "'"
            If .EOF Then
                MsgBox "No tengo esta cuenta registrada..", vbInformation, App.ProductName
                Exit Sub
            Else
                StrControl1.Text = IIf(IsNull(.Fields(StrRecord)), "----VACIO----", _
                .Fields(StrRecord))
            End If
        Else
            StrControl1.Text = IIf(IsNull(.Fields(StrRecord)), "----VACIO----", _
            .Fields(StrRecord))
        End If
    End With
    '
    End Sub
    
    
    Private Sub Txt_DblClick(Index As Integer)
    '
    If Index = 5 Or Index = 12 Then
        '
        If Txt(Index) = "" Then Txt(Index) = 0
        If CLng(Txt(Index)) > 0 Then
            frmResto.curBs = Txt(Index)
            frmResto.Show
        End If
        '
    End If
    '
    End Sub

    Private Sub txt_GotFocus(Index As Integer)
    
    If Index = 5 Then Command1(1).Enabled = True
    Txt(5).SelStart = 0
    Txt(5).SelLength = 20
    
    Select Case Index
    
        Case 1
            If Txt(5) = "" Then Cmb(0).SetFocus: Exit Sub
            Txt(12) = IIf(IsNull(Txt(12)) Or Txt(12) = "", 0, Txt(12))
            Txt(Index) = Format(Txt(5) - CCur(Txt(12)), "#,##0.00")
            With Txt(Index)
                .SelStart = 0
                .SelLength = Len(.Text)
            End With
            
        Case 2
        
            If (Txt(1)) = "" Then Cmb(0).SetFocus: Exit Sub
            Txt(12) = IIf(IsNull(Txt(12)) Or Txt(12) = "", 0, Txt(12))
            Txt(Index) = Format(CCur(Txt(5)) - CCur(Txt(1)) - CCur(Txt(12)), "#,##0.00")
            With Txt(Index)
                .SelStart = 0
                .SelLength = Len(.Text)
            End With
            
        Case 3
        
            If Txt(1) = "" Or (Txt(2)) = "" Then Cmb(0).SetFocus: Exit Sub
            Txt(12) = IIf(IsNull(Txt(12)) Or Txt(12) = "", 0, Txt(12))
            Txt(Index) = Format(CCur(Txt(5)) - CCur(Txt(1)) - CCur(Txt(2)) - CCur(Txt(12)), "#,##0.00")
            With Txt(Index)
                .SelStart = 0
                .SelLength = Len(.Text)
            End With
            '
    End Select
    '
    End Sub
    
    'REV.12/08/2002-------------------------------------------------------------------------
    Private Sub txt_KeyPress(Index As Integer, KeyAscii As Integer)
    '---------------------------------------------------------------------------------------------
    '
    KeyAscii = Asc(UCase(Chr(KeyAscii)))    'CONVIERTE TODO MAYUSCULAS
    '
    Select Case Index
    '
        Case 1, 2, 3, 5, 12
    '   ---------------------
            If KeyAscii = 46 Then KeyAscii = 44 'CONVIERTE EL PUNTO EN COMA
            Call Validacion(KeyAscii, "0123456789,")
    '
        Case 4, 6, 7, 8
    '   ------------------
            Call Validacion(KeyAscii, "1234567890")
    End Select
    '
    '--------------------------------------------------------------------------------------------
    'RUTINA QUE MANEJA EL RECORRIDO DEL CURSOR DENTRO DE LOS CUADROS DE TEXTO*
    '--------------------------------------------------------------------------------------------
    '
    If KeyAscii = 8 Then 'CUANDO PRESIONA BAKSPACE ENTONCES
    '
        Select Case Index
    '
            Case 6, 7, 8
    '       --------------
                If Txt(Index) = "" Then Cmb(Index - 1).SetFocus
    '
        End Select
    '
    End If
    '
    If KeyAscii = 13 Then   'CUANDO PRESIONE ENTER ENTONCES
    '
    Select Case Index
    '
        Case 1, 2, 3    'CUADRO DE TEXTO MONTO DE CHEQUES
    '   -----------------
            Dim Resp As Boolean
            If Txt(Index) <> "" Then
                Resp = FtnValidaCheque(Index)
                If Resp = True Then Exit Sub
            End If
            
    '
            For I = 1 To 3  'SUMA EL CONTENIDO DE LOS TRES EN LA VARIABLE SUMA
                Txt(I) = IIf(Txt(I) = "", 0, Txt(I))
                suma = suma + CCur(Txt(I))
                Txt(I) = Format(CCur(Txt(I)), "#,##0.00")
            Next
            '
            suma = suma + CCur(Txt(12))
            'SI EL CAMPO MONTO ESTA VACIO SE LLENA CON LA VARIABLE SUMA
            If Txt(5) = "" Then Txt(5) = Format(CCur(suma), "#,##0.00")
            'SI LA VARIABLE SUMA ES MAYOR QUE EL MONTO A CANCELAR
            If CCur(suma) > Txt(5) Then  'SE ENVIA UN MENSAJE DE PANTALLA AL USUARIO, SE SALE
                MsgBox "Monto del Cheque mayor a Total a Pagar", vbExclamation, App.ProductName
                Txt(Index).SelStart = 0
                Txt(Index).SelLength = Len(Txt(Index))
                Exit Sub
            End If
            ' CUANDO LA SUMA DE LOS CHEQUES COMPLETE EN TOTAL DEL
            If CCur(suma) = Txt(5) Then 'MONTO A CANCELAR PASA EL FOCO A COD.CTA.INGRESO/EGRESO
                Dat(5).Enabled = True: Dat(6).Enabled = True: Dat(5).SetFocus
            Else
                '
                If Index = 3 Then     'SI ESTA EN EL MONTO(3) PASA EL FOCO COD.CTA.INGRESO/EGRESO
                    Dat(5).Enabled = True: Dat(6).Enabled = True: Dat(5).SetFocus
                Else
                Cmb(Index + 5).Enabled = True
                Cmb(Index + 5).SetFocus     'SI NUNGUNA DE LAS ANTERIORES PASA EL FOCO A LA
                    '
                    Select Case Cmb(0).Text 'SIGUIENTE LINEA
                        '
                        Case "CHEQUE", "AMBOS"  'SI FORMA DE PAGO CHEQUE O AMBOS SELECCIONA CHEQUE
                            Cmb(Index + 5).ListIndex = 0
                        '
                        Case "DEPOSITO" 'SI FORMA DE PAGO DEPOSITO SELECCIONA DEPOSITO
                            Cmb(Index + 5).ListIndex = 1: Txt(Index + 6).SetFocus
                        '
                        Case "TJTA.CREDITO" 'SI FORMA DE PAGO T.CREDITO SELECCIONA T.CREDITO
                            Cmb(Index + 5).ListIndex = 2:: Txt(Index + 6).SetFocus
                        '
                        Case "TJTA.DEBITO"  'SI FORMA DE PAGO T.DEBITO SELECCIONA T.DEBITO
                            Cmb(Index + 5).ListIndex = 3:: Txt(Index + 6).SetFocus
                    End Select
                    '
                End If
            '
            End If
        '
        Case 5      'Monto a cancelar
        '---------------------------------
            If Dat(0) = sysCodInm Or FlexFacturas = "" Then
                Cmb(0).Enabled = True: Cmb(0).SetFocus
            Else
                Command1(1).SetFocus
            End If
            Txt(5) = Format(Txt(5), "#,##0.00")
            
        Case 12     'Bs. en efectivo
        '-------------------------------
            Txt(Index) = Format(Txt(Index), "#,##0.00")
            Cmb(5).ListIndex = 0: Cmb(5).SetFocus
         '
    End Select
    '
    If Index > 5 And Index < 9 Then Cmb(Index - 4).Enabled = True: Cmb(Index - 4).SetFocus
    '
    End If
    '
    End Sub
    
    '---------------------------------------------------------------------------------------------
    Sub RtnMuestraCalendar(Indice As Integer)   '
    '---------------------------------------------------------------------------------------------
    '
    MntCalendar.Visible = Not MntCalendar.Visible
    If MntCalendar.Visible = True Then
        MntCalendar.Top = 4680 + Command(Indice).Top
        MntCalendar.Left = 4080 + (SSTab1.Left - 75)
        MntCalendar.SetFocus
        MntCalendar.Value = Date
        MntCalendar.Tag = 1
    Else
        Txt(Indice).SetFocus
        MntCalendar.Tag = 0
        'Exit Sub
    End If
    End Sub
    
    
    'REV.-07/08/2002-------------------------------------------------
    Sub RtnQdfPropietario(StrParametro As String) '-
    '-----------------------------------------------------------------------
    'Variables a nivel de procedimiento
    Dim objCmd1 As New ADODB.Command
    '----------------------------------
    If cnnPropietario.State = 1 Then cnnPropietario.Close
    If StrRutaInmueble = "" Then StrRutaInmueble = gcPath + "\" + Dat(0) + "\inm.mdb"
    cnnPropietario.Open cnnOLEDB + StrRutaInmueble
    With objCmd1
        .ActiveConnection = cnnPropietario
        .CommandType = adCmdText
        .CommandText = "SELECT * FROM propietarios WHERE Codigo='" & Trim(StrParametro) & "'"
    End With
    'Ejecuta el comando y llena el ADODB.Recordset
    Set objRst = objCmd1.Execute
    '
    With objRst
    '   Asigna los valores encontrados a los controles señalados
        If objRst.EOF Then _
        MsgBox "No Tengo Registrado ese propietario '" & StrParametro & "'", vbInformation, _
        App.ProductName: Exit Sub
        DatProp = !Nombre:   'Dat(2) = !Codigo
        Txt(13) = IIf(IsNull(!Notas), "", !Notas)
        Txt(10) = IIf(IsNull(!Deuda), 0, Format(!Deuda, "#,##0.00"))
        Txt(11) = "0,00"
    End With
    'Destruye el Recorset y la conexion al inmueble del propietario
    objRst.Close: Set objRst = Nothing: cnnPropietario.Close: Set cnnPropietario = Nothing
    '
    End Sub
    
    
    Sub RtnClaveGrid()
    
    Dim IntClave As Long
        If IsNull(ADOcontrol(0).Recordset.Fields("IndiceMovimientoCaja")) Then Exit Sub
        IntClave = ADOcontrol(0).Recordset.Fields("IndiceMovimientoCaja")
            With ADOcontrol(3).Recordset
                If Not .EOF And Not .BOF Then
                    .MoveFirst
                    .Find "IndiceMovimientoCaja = '" & IntClave & "'"
                End If
            End With
    
    End Sub
    
    
    
    Sub rtnDistribuye()
    Command2.Top = (FlexDeducciones.Top + (315 * (FlexDeducciones.Rows - 1)))
    End Sub
    
    
    '------------------------------------------
    Sub RtnTemDeducciones(StrRuta As String) '-
    '------------------------------------------
    'variables locales
    Dim ObjCmd As New ADODB.Command
    Dim I As Integer
    '
    With AdoDeducciones
    '
        .ConnectionString = cnnOLEDB + StrRutaInmueble
        .CommandType = adCmdTable
        .RecordSource = "TemDeducciones"
        .Refresh
    '
            With .Recordset     'agrega los datos del grid al recorset
                'N1
'                If Trim(Label16(12)) <> "" Then
'                    strPeriodo = Trim(Label16(12))
'                Else
                    strPeriodo = strRecibo & Left(Trim(Label16(13)), 2) & _
                    Right(Trim(Label16(13)), 2)
'                End If
    '
                For I = 1 To FlexDeducciones.Rows - 1
    '
                    If FlexDeducciones.TextMatrix(I, 0) <> "" And _
                    FlexDeducciones.TextMatrix(I, 1) <> "" And _
                    FlexDeducciones.TextMatrix(I, 2) <> "" Then
                    '
                        .AddNew
                        .Fields("IDperiodos") = strPeriodo
                        .Fields("NumFact") = IIf(Trim(Label16(12)) <> "", Trim(Label16(12)), _
                        strRecibo)
                        .Fields("CodGasto") = FlexDeducciones.TextMatrix(I, 0)
                        .Fields("Titulo") = FlexDeducciones.TextMatrix(I, 1)
                        .Fields("Monto") = CCur(FlexDeducciones.TextMatrix(I, 2))
                        .Fields("Autoriza") = 0
                        .Fields("Usuario") = gcUsuario
                        .Fields("FecReg") = Date
                        .Fields("Taquilla") = IntTaquilla
                        .Update
                        If FlexDeducciones.TextMatrix(I, 0) = strCodRebHA Then
                            curPagoHono = CCur(LblHono) - CCur(FlexDeducciones.TextMatrix(I, 2))
                            'booHA = IIf(curPagoHono = 0, False, True)
                        End If
                        '
                    End If
    '
                Next
    '
            End With
            .Recordset.Close    'cierra el objeto ADODB.Recordset y la conexion
            frame3(3).Visible = True
    '
            'N3
            With AdoDeducciones
                .ConnectionString = cnnOLEDB & StrRutaInmueble
                .CommandType = adCmdText
                .RecordSource = "SELECT TemDeducciones.Autoriza,temdeducciones.codgasto " _
                & "From TemDeducciones WHERE TemDeducciones.IDPeriodos= '" & strPeriodo & "'"
                .Refresh
            End With
            TimDemonio.Interval = 30
    '
    End With
    '
    End Sub

    'Rev.27/08/2002-------------------------------------------------------------------------------
    '   Rutina:     rtnDeducciones
    '
    '   Se agregan los registros a las facturas del cliente, se agrega un único registro
    '   al movimiento de la caja por el total deducciones.
    '---------------------------------------------------------------------------------------------
    Private Sub RtnDeducciones(strMes As String) '-
    'variables locales
    Dim totalDed As Currency
    Dim I As Integer
    Dim StrP As String
    I = 0
    
    With AdoDeducciones
    '
        .ConnectionString = cnnOLEDB + StrRutaInmueble
        .CommandType = adCmdText
        .RecordSource = "SELECT DISTINCT * From TemDeducciones WHERE NumFact='" & strMes & _
        "' AND IDPeriodos='" & strPeriodo & "'"
        .Refresh
    '   NO TIENE DEDUCCIONES SALE DE LA RUTINA*
        With .Recordset
            If Not .EOF Or Not .BOF Then
                .MoveFirst
                StrP = !IDPeriodos
                
                Do
                    I = I + 1
        '           Se agregan los registros de las deducciones autorizadas
                    'cnnConexion.Execute "INSERT INTO Deducciones (IDperiodos, NumFact, " _
                    & "CodGasto, Titulo, Autoriza,Monto,Usuario,FecReg) VALUES ('" & _
                    !IDPeriodos & "','" & !NumFact & "','" & !CodGasto & "','" & !Titulo _
                    & "',TRUE,'" & !Monto & "','" & !Usuario & "',Date())"
                    If Not !codGasto = strCodRebHA Then
        '           Se agrega la inf. al recibo del propietario
                        cnnConexion.Execute "INSERT INTO DetFact(Fact,Detalle,Codigo,CodGasto,P" _
                        & "eriodo,Monto,Fecha,Hora,Usuario) IN '" & StrRutaInmueble & "' SELECT" _
                        & " TOP 1 '" & !NumFact & "','" & !Titulo & "','" & Dat(2) & "','" & _
                        !codGasto & "',Periodo,'" & !Monto * -1 & "',Date(),Format(Time(),'hh:m" _
                        & "m:ss'),'" & !Usuario & "' FROM DetFact IN '" & StrRutaInmueble & _
                        "' WHERE Fact='" & strMes & "';"
                    End If
                    '
                    '
                    If .RecordCount = 1 Then    'Si es una sola deducción
                        cnnConexion.Execute "INSERT INTO Deducciones (IDperiodos, CodGa" _
                        & "sto, Titulo, Autoriza,Monto,Usuario,FecReg) VALUES ('" & !IDPeriodos _
                        & "','" & !codGasto & "','" & !Titulo & "',TRUE,'" & !Monto & "','" _
                        & !Usuario & "',Date())"
                    End If
                    '
                    '
                    totalDed = totalDed + !Monto
                    .MoveNext
                Loop Until .EOF
                booDed = True
                'Se agrega un único registro a las deducciones autorizadas
                If I > 1 Then
                    cnnConexion.Execute "INSERT INTO Deducciones (IDperiodos, CodGasto,Titulo," _
                    & "Autoriza,Monto,Usuario,FecReg) VALUES ('" & StrP & "','999999','DEDUCCIO" _
                    & "NES VARIAS',TRUE,'" & totalDed & "','" & gcUsuario & "',Date())"
                End If
                '
            End If
            '
        End With
        '
    End With
    '
    End Sub

    Sub RtnBuscaInmueble(Qdf$, DC As DataCombo)
    '
    Txt(13) = ""
    Call BuscaInmueble(Qdf, DC)
    Set objRst = ObjCmd.Execute
    Call rtnLimpiar_Grid(FlexFacturas)
    With objRst
        If objRst.EOF Then
            For I = 1 To 4: Dat(I) = ""
            Next
            MsgBox "Inmueble No Registrado..", vbInformation, App.ProductName
            Dat(0).SetFocus
            Exit Sub
        End If
        Dat(0) = .Fields("CodInm")
        Dat(1) = .Fields("Nombre")
        Dat(3) = .Fields("Caja")
        Dat(4) = .Fields("DescripCaja")
        Lbl(20) = IIf(IsNull(.Fields("Notas")) Or .Fields("notas") = "", "", .Fields("Notas"))
        StrRutaInmueble = gcPath & .Fields("Ubica") & "inm.mdb"
        strCodIV = .Fields("CodIngresosVarios")
        strCodHA = .Fields("CodHA")
        strCodRebHA = .Fields("CodRebHA")
        strCodAbonoCta = .Fields("CodAbonoCta")
        strCodAbonoFut = .Fields("CodAbonoFut")
        strCodPC = .Fields("CodPagocondominio")
        strCodCCHeq = .Fields("CodCCheq")
        strCodRCheq = .Fields("CodRCheq")
        IntHonoMorosidad = .Fields("HonoMorosidad")
        IntMesesMora = .Fields("MesesMora")
        ADOcontrol(1).ConnectionString = cnnOLEDB + StrRutaInmueble
        ADOcontrol(1).CommandType = adCmdText
        ADOcontrol(1).RecordSource = "SELECT CodGasto,Titulo  FROM TGastos WHERE CodGasto Like '9" _
        & "0%' OR CodGasto Like '8%' OR Fondo=True;"
        ADOcontrol(1).Refresh
        '*********************************************
        If FmeCuentas.Enabled = False Then
            Set objRst = Nothing
            Exit Sub
        End If
        Dat(2) = ""
        DatProp = ""
        FlexFacturas.Rows = 2
        Dat(2).Enabled = True
        DatProp.Enabled = True
        Dat(2).SetFocus
        '
    End With
    
    Set objRst = Nothing
    Call RtnListapro '("SELECT * FROM Propietarios ORDER BY Codigo")
    
    objRst.Open "SELECT MAX(NumeroMovimientoCaja) AS MAXIMO FROM Inmueble INNER JOIN MOVIMIENTO" _
    & "CAJA ON inmueble.CodInm = MOVIMIENTOCAJA.InmuebleMovimientoCaja WHERE (((MOVIMIENTOCAJA." _
    & "FechaMovimientoCaja)=Date()) AND ((inmueble.Caja)='" & Dat(3) & "'));", cnnConexion, _
    adOpenKeyset, adLockOptimistic

    If VarType(objRst.Fields(0)) = vbNull Then
        
        Txt(0) = Format(1, "00")
    
    Else
        
        Txt(0) = Format(objRst.Fields(0) + 1, "00")
    
    End If
    '
    objRst.Close
    Set objRst = Nothing
    '
    End Sub
    '-------------------------------------------------
    '               '   RUTINA QUE SE EJECUTA CADA VEZ
    Sub RtnAvanza() '   QUE SE MUEVE EL CURSOR DENTRO
    '               '   DEL ADODB.Recordset
    '-------------------------------------------------
    Call RtnBuscaInmueble("Inmueble.CodInm", Dat(0))
    Call RtnQdfPropietario(Dat(2).Text)
    Call RtnClaveGrid
    Call rtnLimpiar_Grid(FlexFacturas)
    '
    End Sub

'Rev 12/08/2002-----------------------------------------------------------------------------------
'   función:    rtnvalidar
'
'   verifica que se tengan todos los datos mínimos requeridos para procesar una trans.
'   Devuelve True si no pasa el proceso satisfactoriamente, de lo contrario retorna false
'-------------------------------------------------------------------------------------------------
Private Function ftnValidar() As Boolean
'variables locales
Dim I%, Titulo$, Marca As Boolean, CurCheques@
'
Titulo = App.ProductName

For I = 0 To 2  'VERIFICA LOS DATOS DEL INMUEBLE/PROPIETARIO
'
    
    If Dat(I) = "" Then
        ftnValidar = MsgBox("Falta '" & Dat(I).ToolTipText & "'", vbExclamation, Titulo)
        Dat(I).SetFocus
        Exit Function
    End If
'
Next
'imposible efectuar abonos a futuro si el prop. es la unidad del facturacion
If Dat(2) = "U" & Dat(0) And Dat(5) = strCodAbonoFut Then
    ftnValidar = MsgBox("Imposible ingresar un 'abono a futuro' cargado a " & DatProp, _
    vbExclamation, App.ProductName)
End If
'
If Txt(9) = "" Then         'Valida la descripción de la operación
    ftnValidar = MsgBox("Falta la Descripción de la operación..", vbExclamation, Titulo)
    Txt(9).SetFocus
    Exit Function
End If
'
For I = 0 To 1  'VERIFICA FORMA DE PAGO / TIPO DE MOVIMIENTO

    If Cmb(I) = "" Then
'
        ftnValidar = MsgBox("Falta '" & Cmb(I).ToolTipText & "'", vbExclamation, Titulo)
        Cmb(I).SetFocus
        Exit Function
'
    End If
    '
Next
'
'DE ACUERDO A LA FORMA DE PAGO VERIFICA EL DETALLE DEL MISMO
Select Case Cmb(0).ListIndex
'
    Case 0  'EN CASO DE 'AMBOS'
'               -----------------------------

        If Txt(12) = "" Then
            ftnValidar = MsgBox("Falta '" & Txt(12).ToolTipText & "'", vbExclamation, Titulo)
            Exit Function
        End If
'
    Case 3  'EFECTIVO
'   -----------------------
        If Txt(5) < 0 Or Txt(5) = "" Then
            ftnValidar = MsgBox("Error en el Monto a Cancelar", vbCritical, Titulo)
            Exit Function
        End If
    Case 1, 2, 4, 5  'CASO 'CHEQUE','DEPOSITO' O 'TARJETAS'
'   -----------------------------------------------------------------------------
        For I = 1 To 3
            If IsNull(Txt(I)) Then Txt(I) = "0,00"
            If Txt(I) = "" Then Txt(I) = "0,00"
        Next
        CurCheques = CCur(Txt(1) + CCur(Txt(2)) + CCur(Txt(3)))
         For I = 5 To 7
            If (Cmb(I) = "DEPOSITO" Or Cmb(I) = "TRANSFERENCIA") And Cmb(I).Tag = "" Then
                ftnValidar = MsgBox("No selecciono ninguna de las cuentas del condominio...", _
                vbInformation, Titulo)
                Exit Function
            End If
            If Cmb(I) <> "DEPOSITO" And Cmb(I) <> "TRANSFERENCIA" Then Cmb(I).Tag = ""
         Next I
End Select
'
For I = 5 To 6  'VERIFICA LOS DATOS CODIGO DE OP.
'
    If Dat(I) = "" Then
'
        ftnValidar = MsgBox("Falta '" & Dat(I).ToolTipText & "'", vbExclamation, Titulo)
        If Not Dat(I).Enabled Then Dat(I).Enabled = True: Dat(I).SetFocus
        Exit Function
'
    End If
'
Next
'
If Dat(5) = strCodPC Or Dat(5) = strCodAbonoCta Then
        
20      With FlexFacturas  'ESTE MARCARDA EN EL FLEXFACTURA
            Dim Total As Currency
            .Col = 6
            
            For I = 1 To (.Rows - 1)
                .Row = I
                If .CellPicture <> 0 Then   'VERIFICA EL TOTAL DE LO PAGADO
                    Marca = True
                    Total = Total + IIf(FlexFacturas.TextMatrix(I, 4) <> 0, _
                    CCur(FlexFacturas.TextMatrix(I, 3)) - CCur(vecFPS(I, 1)), vecFPS(I, 2))
                End If
            Next I
            If Not Marca Then   'debe haber por lo menos una factura marcada para cancelar
                ftnValidar = MsgBox("Debe Marcar por lo Menos Una Factuta Para Cancelar", _
                vbExclamation, Titulo)
                Exit Function
            End If
            Total = Total - CCur(Label16(17)) + curHono + IntMonto
                
            If Total <> CCur(Txt(5)) Then  'SI CAMPO MONTO DIFERENTE A:
                                           'FACTURAS MARCADAS - HON.ABOG. - DEDUCCIONES
                ftnValidar = MsgBox("Ha Transgredido Niveles de Seguridad," & Chr(13) & "Monto " _
                & "Marcado para Cancelar: " & Format(Total, "#,##0.00") & Chr(13) & "Monto Regi" _
                & "strado a Cancelar: " & Format(Txt(5), "#,##0.00"), vbExclamation, Titulo)
                Exit Function
                    
            End If
            If Total <> CurCheques And Cmb(0).ListIndex <> 0 And Cmb(0).ListIndex <> 3 Then
                ftnValidar = MsgBox("Total Detalle del Pago discrepa del " & vbCr & "Total a Ca" _
                & "ncelar, Revise Monto del " & vbCr & "o los " & Cmb(0).Text & "S.....", _
                vbExclamation, Titulo)
                Exit Function
            End If
            '
        End With
End If
'
End Function

'-------------------------------------------------------------------------------
Sub RtnDescripcion()
'------------------------------------------------------------------------------
'RUTINA QUE LEE LAS FACTURAS MARCADAS PARA CANCELAR*
'       GENERA LA DESCRIPCION DEL PAGO             *
'
Dim X%, Z%, strDesde$, strHasta$
'
With FlexFacturas
'LLENE UNA MATRIZ CON LOS PERIODOS MARCADOS PARA CANCELAR
Dim vecFECHA(1 To 120) As String
.Col = 6
X = 1
    For I = 1 To .Rows - 1
        .Row = I
        If .CellPicture <> 0 Then
        
            vecFECHA(X) = .TextMatrix(I, 1)
            X = X + 1
            
        End If
    Next
    
End With
X = X - 1
Z = 1
Txt(9) = ""

Do While X >= Z
    strDesde = Left(vecFECHA(Z), 2)
    If Z <> X Then
        Do While CInt(Left(vecFECHA(Z), 2) + 1) = CInt(Left(vecFECHA(Z + 1), 2))
            Z = Z + 1
            strHasta = Left(vecFECHA(Z), 2)
            If Z = X Then Exit Do
        Loop
    End If

If strHasta = "" Then
    If Txt(9) = "" Then
        Txt(9) = vecFECHA(Z)
    Else
        Txt(9) = Txt(9) + " / " + vecFECHA(Z)
    End If
Else
    If Txt(9) = "" Then
        Txt(9) = strDesde + " A " + vecFECHA(Z)
    Else
        Txt(9) = Txt(9) + " / " + strDesde + " A " + vecFECHA(Z)
    End If
End If
Z = Z + 1
strHasta = ""
Loop
End Sub

    'Rev-22/08/2002-------------------------------------------------------------------------------
    Private Sub RtnBuscaCaja(A%, B%)
    'variables locales
    Dim strSQL As String
    Dim RstCuentas As ADODB.Recordset
    Dim RstCheques As New ADODB.Recordset
    Dim RstEfectivo As ADODB.Recordset
    Dim CnnCuentas As ADODB.Connection
    '
    Command3(5).Enabled = False
    Dat(A) = Dat(B).BoundText
    frame3(7).Caption = "Detalle de Deposito"
    Label16(30) = "0,00"
    Label16(31) = "0,00"
    Txt(4) = ""
    '
    Set RstCuentas = New ADODB.Recordset
    Set RstEfectivo = New ADODB.Recordset
        '
    RstCuentas.Open "SELECT Inmueble.CodInm,Inmueble.Ubica FROM Caja INNER JOIN Inmueble ON" _
    & " Caja.CodigoCaja = Inmueble.Caja WHERE Caja.CodigoCaja='" & Dat(7) & "'", cnnConexion, _
    adOpenKeyset, adLockReadOnly, adCmdText
        '
    If RstCuentas.EOF Or RstCuentas.BOF Then Exit Sub
    '
    strUbica = IIf(Dat(7) = sysCodCaja, "\" & sysCodInm & "\", RstCuentas!Ubica)
    Dat(7).Tag = IIf(Dat(7) = sysCodCaja, sysCodInm, RstCuentas!CodInm)
    '
    Set CnnCuentas = New ADODB.Connection
    CnnCuentas.Open cnnOLEDB & gcPath & strUbica + "inm.mdb"
    Set RstCuentas = New ADODB.Recordset
    '
    RstCuentas.Open "SELECT Cuentas.NumCuenta, Bancos.NombreBanco FROM Bancos INNER JOIN Cuenta" _
    & "s ON Bancos.idBanco = Cuentas.idBanco WHERE Cuentas.Inactiva=False ORDER BY Pred", _
    CnnCuentas, adOpenKeyset, adLockReadOnly, adCmdText
    '
    If Not RstCuentas.EOF Then
    
        Set Dat(9).RowSource = RstCuentas
        Set Dat(10).RowSource = RstCuentas
        
        For I = 9 To 10: Dat(I).Text = RstCuentas.Fields(I - 9)
        Next
        If UCase(RstCuentas.Fields("NombreBanco")) Like "*PROVINCIAL*" Then
            Randomize
            Txt(4) = Int((999999 - 1 + 1) * Rnd + 1)
        End If
        Txt(4).SetFocus
        
    End If
    '
    strSQL = "SELECT Ndoc,Banco,Monto FROM TdfCheques WHERE CodInmueble IN (SELECT CodI" _
    & "nm FROM Inmueble WHERE Caja='" & Dat(7) & "') AND FechaMov=Date() and IDTaquilla" _
    & "=" & IntTaquilla & " and (IsNull(IdDeposito) Or IDDeposito='') and Fpago='CHEQUE';"
    '
    RstCheques.Open strSQL, cnnConexion, adOpenKeyset, adLockBatchOptimistic, adCmdText
    
    If Not RstCheques.EOF Then RstCheques.MoveFirst
        I = 0
        Call rtnLimpiar_Grid(GridCheques(0))
        GridCheques(0).Rows = RstCheques.RecordCount + 2
        Label16(26) = RstCheques.RecordCount
        '
        Do Until RstCheques.EOF
            I = I + 1
                With GridCheques(0)
                For j = 0 To 2
                    .TextMatrix(I, j) = IIf(j = 2, Format(RstCheques.Fields(j), "#,##0.00"), RstCheques.Fields(j))
                Next
                RstCheques.MoveNext
                End With
        Loop
        RstCheques.Close
        Set RstCheques = Nothing
        '
        RstEfectivo.Open "SELECT Sum(TDFCheques.Monto) AS EFECTIVO FROM Caja INNER JOIN (TD" _
        & "FCheques INNER JOIN Inmueble ON TDFCheques.CodInmueble = Inmueble.CodInm) ON Caj" _
        & "a.CodigoCaja = Inmueble.Caja WHERE (((Caja.CodigoCaja)='" & Dat(7) & "') AND ((T" _
        & "dfCheques.FechaMov)=Date()) And ((TdfCheques.IDTaquilla)=" & IntTaquilla & ") AND" _
        & " ((TdfCheques.IdDeposito)Is Null) AND ((TdfCheques.Fpago)='EFECTIVO')) GROUP BY " _
        & "Caja.CodigoCaja", cnnConexion, adOpenKeyset, adLockReadOnly
            
        If Not RstEfectivo.EOF Then
            Label16(29) = Format(RstEfectivo.Fields(0), "#,##0.00")
        Else
            Label16(29) = Format(0, "#,##0.00")
        End If
        RstEfectivo.Close
        GridCheques(0).Row = 1
        GridCheques(0).Col = 0
        Call rtnLimpiar_Grid(GridCheques(1))
        GridCheques(1).Rows = 3
        CurTotalCheque = 0
    '
    End Sub

    
    '------------------------------------------------------------------------------------
    Private Sub RtnAbonoFut(DblMonto@, strPer$)     'Rev.12/08/2002   '
    '------------------------------------------------------------------------------------
    'Rutina que registra el abono a futuro en SAC
    '*****INSERTA EL REGISTRO EN TDFAbonos*********************
    cnnConexion.Execute "INSERT INTO TDFAbonos(IDRecibo,Monto) " _
    & "VALUES ('" & strRecibo & "','" & DblMonto & "')"
    '
    '*****INTRODUCE LA INFORMACION EN TDFPeriodos*************
    cnnConexion.Execute "INSERT INTO Periodos(IDRecibo,IDPeriodos, Periodo, CodGasto, " _
    & "Descripcion,Monto,Facturado) VALUES ('" & strRecibo & "','" & strRecibo & strCodAbonoFut _
    & "','" & strPer & "','" & strCodAbonoFut & "','Abono Próx. Facturación','" _
    & DblMonto & "',0 )"
    '
    booAFuturo = True
    '
    End Sub

    

    '21/08/2002----Rutina que verifica que el documento que esta entrando en caja no fuese aplicado
    Private Function FtnValidaCheque(Indice%) As Boolean  'en dias anteriores
    '-------------------------------------------------------------------------------------------------
    '
    Dim ObjCheques As ADODB.Recordset
    Set ObjCheques = New ADODB.Recordset
    '-------------------------------------------------------------------------------------------------
    'Valida los campos requeridos por el registro {FPago + Ndoc + Banco + FechaDoc + Monto }
    If Txt(Indice) > 0 Then  'Si la cantidad es superior a cero
    '
            If Cmb(Indice + 4) = "" Then Cmb(Indice + 4).SetFocus: FtnValidaCheque = _
                MsgBox("Falta '" & Cmb(Indice).ToolTipText & "'")
            If Txt(Indice + 5) = "" Then Txt(Indice + 5).SetFocus: FtnValidaCheque = _
                MsgBox("Falta '" & Txt(Indice + 5).ToolTipText & "'")
            If Cmb(Indice + 1) = "" Then Cmb(Indice + 1).SetFocus: FtnValidaCheque = _
                MsgBox("Falta '" & Cmb(Indice + 1).ToolTipText & "'")
            If MskFecha(Indice) = "" Then MskFecha(Indice).SetFocus: FtnValidaCheque = _
                MsgBox("Falta '" & MskFecha(Indice).ToolTipText & "'")
            If FtnValidaCheque = True Then Exit Function
    '
    End If
    '
    '--------------------------------------------------Selecciona alguna coincidencia-----------------
    ObjCheques.Open "SELECT * FROM tdfCheques WHERE Ndoc ='" & Txt(Indice + 5) & "' AND BANCO ='" _
    & "" & Cmb(Indice + 1) & "' AND FPago='" & Cmb(Indice + 4) & "'", cnnConexion, adOpenKeyset, _
    adLockOptimistic
    '-------------------------------------------------------------------------------------------------
    With ObjCheques
    '
    If Not .EOF Then  'si existe alguna coincidencia, valida la fecha de procesamiento
    '
        If !FechaMov = Date Then    'fue procesado hoy valida que sea por la misma taquilla
            If !IDTaquilla <> IntTaquilla Then
    '           El cheque fué ingresado por otra taquilla exit sub
                FtnValidaCheque = MsgBox("Este " & !FPago & " fué ingresado por la taquilla N° '" _
                    & !IDTaquilla & "' El día de Hoy, por favor verifique el pago de este cliente", _
                    vbCritical, !FPago & " Ingresado")
            End If
        Else
    '       Imposible agregar este cheque, ya fué aplicado anteriormente
            FtnValidaCheque = MsgBox("Este " & !FPago & " ya fué ingresado a caja, el dia " _
                & !FechaMov & ". CodInm.: '" & !CodInmueble & "' - Apto.: '" & _
                Mid(!IDRecibo, 3, Len(!IDRecibo) - 10), vbCritical, !FPago & " INGRESADO...")
        End If
    '
    End If
    '
    End With
    '
    End Function

    '26/08/2002--Rutina que marca una factura para cancelar, actualiza la información-------------
    Private Sub RtnFactura(booEvento As Boolean)    'en pantalla
    '---------------------------------------------------------------------------------------------
    '
    With FlexFacturas   'SE LLENA EL VECTOR FACTURADO/PAGADO/SALDO
    '                   'NECESARIO EN CASO DE REVERTIR ALGUN REGISTRO
        For I = 0 To 2
            vecFPS(.RowSel, I) = .TextMatrix(.RowSel, I + 2)
            'Debug.Print vecFPS(.RowSel, I)
        Next
        .Row = .RowSel
        .Col = 6
        .CellPictureAlignment = flexAlignRightTop
        Set .CellPicture = ImgAceptar(0).Picture    'Marca la factura
    '           ----------------------------------------------------------------------------------
        If Txt(5) = "" Then Txt(5) = 0
        If booEvento = True Then    'La llamada la hace el evento click
            Txt(10) = Format(CCur(Txt(10) - .TextMatrix(.Row, 4)), "#,##0.00")
            Txt(5) = Format(CCur(Txt(5)) + CCur(.TextMatrix(.Row, 4)), "##,##0.00")
            .TextMatrix(.Row, 3) = Format(CCur(.TextMatrix(.Row, 4)) + _
                CCur(.TextMatrix(.Row, 3)), "#,##0.00")
    '
        Else    'La llamada la hace el evento KeyPress
    '
            Txt(10) = Format(CCur(Txt(10) - .TextMatrix(.RowSel, 3)), "#,##0.00")
            Txt(5) = Format(CCur(Txt(5)) + .TextMatrix(.RowSel, 3), "##,##0.00")
            vecFPS(.RowSel, 1) = 0
    '
        End If
        .TextMatrix(.Row, 4) = Format(CCur(.TextMatrix(.Row, 2) - _
                .TextMatrix(.RowSel, 3)), "#,##0.00")
        .TextMatrix(.Row, 6) = "SI"
        Txt(11) = Format(CCur(Txt(10) + CCur(LblHono)), "#,##0.00")
        Txt(9) = IIf(Trim(Txt(9)) = "", "", Txt(9) + " / ") + _
            IIf(.TextMatrix(.Row, 0) Like "CH*", .TextMatrix(.Row, 0), _
            IIf(.TextMatrix(.Row, 4) = 0, .TextMatrix(.Row, 1), "Abono a Cta. " & .TextMatrix(.Row, 1)))
    '           ----------------------------------------------------------------------------------
    End With
    If Cmb(0).Enabled Then Cmb(0).SetFocus
    '
    End Sub

    '---------------------------------------------------------------------------------------------
    Private Sub rtnEditar(blnSINO As Boolean)     '
    '---------------------------------------------------------------------------------------------
    '
    FmeCuentas.Enabled = blnSINO
    Frame5.Enabled = blnSINO
    For I = 0 To 2: Dat(I).Locked = blnSINO
    Next
    DatProp.Locked = blnSINO
    Txt(5).Locked = blnSINO
    For I = 1 To 3
        Txt(I).Locked = blnSINO
    Next
    '
    End Sub
    
    '---------------------------------------------------------------------------------------------
    Private Sub rtnFrmActivo()
    '---------------------------------------------------------------------------------------------
    '
    With SSTab1
    '
        If gcNivel < nuSUPERVISOR Then
            .TabEnabled(0) = True
            .TabEnabled(1) = True
            .TabEnabled(2) = True
            '.TabEnabled(3) = True
        Else
            .TabEnabled(0) = Not blnCaja
            .TabEnabled(1) = blnCaja
            .TabEnabled(2) = blnCaja
            '.TabEnabled(3) = blnCaja
        End If
        '
    End With
    '
    End Sub

    Private Sub rtnTab(intTab%)
    '
    With SSTab1
        .TabEnabled(2) = Not .TabEnabled(2)
        .TabEnabled(0) = Not .TabEnabled(0)
        .tab = intTab
        If intTab = 2 Then FlexDeducciones.SetFocus
    End With
    '
    End Sub
    
    '---------------------------------------------------------------------------------------------
    '   Función:    Beneficiario
    '
    '   Entrada:    Código de Caja
    '
    '   Retorna el nombre del inmueble
    '---------------------------------------------------------------------------------------------
    Private Function Beneficiario(Numero_Cuenta$, Codigo_Inm$)
    'variables locales
    Dim rstBeneficiario  As New ADODB.Recordset
    Dim rstData As String
    
    rstData = gcPath & "\" & Codigo_Inm & "\inm.mdb"
    '
    rstBeneficiario.Open "SELECT * From Cuentas WHERE NumCuenta='" & Numero_Cuenta & "';", _
    cnnOLEDB + rstData, adOpenStatic, adLockReadOnly
    
    Beneficiario = IIf(IsNull(rstBeneficiario!Titular), "", rstBeneficiario!Titular)
    '
    rstBeneficiario.Close
    Set rstBeneficiario = Nothing
    '
    End Function

    Private Sub Forma_Pago()
    '
    Call mostrar_area(Cmb(0).ListIndex)
    Select Case Cmb(0).ListIndex
        '
        Case 3  'Efectivo
            Call RtnVisible("False")
            Dat(5).Enabled = True
            Dat(6).Enabled = True
            Dat(5).SetFocus
            '
            For I = 1 To 3
                Cmb(I + 4) = ""
                Txt(I + 5) = ""
                Cmb(I + 1) = ""
                Txt(I) = ""
            Next I
            '
        Case 1  'cheque
            Call RtnVisible("False")
            Cmb(5).ListIndex = 0
            Txt(6).SetFocus
        
        Case 2  'Deposito
            Call RtnVisible("False")
            Cmb(5).ListIndex = 1
            Txt(6).SetFocus
            Cmb_LostFocus (5)
        
        Case 0  'ambos
            Call RtnVisible("True")
            Txt(12).SetFocus
            
        Case 4  'Tarjeta de Crédito
            Call RtnVisible("False")
            Cmb(5).ListIndex = 2
            Txt(6).SetFocus
            Cmb_LostFocus (5)
            
        Case 5  'Tarjeta de Débito
            Call RtnVisible("False")
            Cmb(5).ListIndex = 3
            Txt(6).SetFocus
            
    End Select
    
    '
    End Sub

    Private Sub mostrar_area(tipo_pago As Integer)
    If tipo_pago = 3 Then   'si el pago es efectivo
        Frame5.Visible = False
        Frame1.Height = 3450
        FlexFacturas.Height = 3270
    Else
        Frame5.Visible = True
        Frame1.Height = 2070
        FlexFacturas.Height = 1875
    End If
    End Sub
    Sub Desmarca()
    On Error Resume Next
    For j = 1 To FlexFacturas.Rows - 1
        vecFPS(j, 2) = 0
        FlexFacturas.Row = j
        FlexFacturas.Col = 6
        Set FlexFacturas.CellPicture = Nothing
        FlexFacturas.Col = 7
        Set FlexFacturas.CellPicture = Nothing
    Next
    End Sub


    '---------------------------------------------------------------------------------------------
    '   Rutina: Guardar_NumFact
    '
    '   Guarda el número del recibo cancelado y el monto real pagado
    '---------------------------------------------------------------------------------------------
    Private Sub Guardar_NumFact(strNumFact As String, Cancelado As Currency)
    Dim numFichero%   'variables locales
    Dim strArchivo$
    '
    numFichero = FreeFile
    strArchivo = App.Path & Archivo_Temp
    Open strArchivo For Append As numFichero
    Write #numFichero, strNumFact, Cancelado
    Close numFichero
    '
    End Sub

    '---------------------------------------------------------------------------------------------
    '   Rutina: Imprimir_Recibos
    '
    '   Recorre el archivo Archivo_Temp y emite un comprobante por cada
    '   factura registrada en él, al llegar al final lo elimina
    '---------------------------------------------------------------------------------------------
    Private Sub Imprimir_Recibos()
    'Variables locales
    Dim strArchivo$, Recibo$
    Dim Pago@, numFichero%
    Dim Carpeta$, CodInm$
    Dim NomInm$
    '
    'On Error GoTo RegError
    '
    numFichero = FreeFile
    strArchivo = App.Path & Archivo_Temp
    '
    Open strArchivo For Input As numFichero 'abre el archivo de recibos cancelados
        'Control = 5
        Do
            '
            Input #numFichero, Recibo, Pago
            
            If CHK.Value = vbUnchecked Then 'imprime el recibo
               Carpeta = "\" & Dat(0) & "\"
               CodInm = Dat(0)
               NomInm = Dat(1)
                If Not Recibo = "" Then Call Printer_Pago(Recibo, Pago, Carpeta _
                , CodInm, NomInm, Recibo, True, 0, crImpresora, , , , True)
                
            Else    'lo coloca en la lista   de recibos enviar a edif
                
                cnnConexion.Execute "INSERT INTO Recibos_Enviar(IDRecibo,Fact,Monto) VALUES ('" _
                & strRecibo & "','" & Recibo & "','" & Pago & "')"
                Call rtnBitacora("Recibo #" & Recibo & " enviado al Edif.")
                
            End If
            If Err <> 0 Then Call rtnBitacora(Err.Description)
            '
        Loop Until EOF(numFichero)
        
    Close numFichero
    '
    Kill strArchivo
    '
    End Sub
    
    
    '-------------------------------------------------------------------------------------
    'Guarda la información en TDFCheques (Detalle de la forma de pago)
    '-------------------------------------------------------------------------------------
    Private Sub Guardar_FPago()
    'Variables locales
    Dim CurEfectivo@, Ind%, CurEfectivoBsF@
    
    Select Case Cmb(0).Text
    '
        Case Is = "EFECTIVO" 'El pago es efectivo
    '   -----------------------------------------------------------------------------
            CurEfectivo = CCur(Txt(5))
            CurEfectivoBsF = convertirBsF(CurEfectivo)
            
            cnnConexion.Execute "INSERT INTO tdfCHEQUES (IDRecibo,IdTaquilla, CodInmueb" _
            & "le, FechaMov, Fpago, Monto, Ndoc, Banco,MontoBs ) Values ('" & strRecibo & "'," & _
            IntTaquilla & ",'" & Dat(0) & "',DATE() ,'EFECTIVO','" & CurEfectivo & "','" _
            & strRecibo & "','s/n','" & CurEfectivo & "')"
            Call rtnBitacora("Pago en efectivo....." & Format(CurEfectivo, "#,##0.00"))
            
        Case Is <> "EFECTIVO"
    '   -----------------------------------------------------------------------------
        '
            If Cmb(0) = "AMBOS" Then
                CurEfectivo = IIf(Cmb(1).Text = "INGRESO", CCur(Txt(12)), CCur(Txt(12) * -1))
                CurEfectivoBsF = convertirBsF(CurEfectivo)
                
                cnnConexion.Execute "INSERT INTO tdfCHEQUES (IDRecibo,IdTaquilla," & _
                "CodInmueble, FechaMov, Fpago, Monto, Ndoc, Banco, MontoBs ) " & _
                "Values ('" & strRecibo & "'," & IntTaquilla & ",'" & Dat(0) & _
                "',date(),'EFECTIVO','" & CurEfectivoBsF & "','" & strRecibo & _
                "','s/n','" & CurEfectivo & "')"
                Call rtnBitacora("Pago en efectivo....." & Format(CurEfectivo, "#,##0.00"))
            End If
            '
            For Ind = 1 To 3
                If Txt(Ind) > 0 Then
                    FormaPago_Add IntTaquilla, Cmb(4 + Ind), Txt(5 + Ind), Cmb(1 + Ind), _
                    MskFecha(Ind), Txt(Ind), strRecibo, Dat(0), _
                    IIf(Cmb(4 + Ind).Text = "CHEQUE", "", Cmb(4 + Ind).Tag)
                    Call rtnBitacora(Cmb(4 + Ind) & " " & Txt(5 + Ind) & " " & Cmb(1 + Ind) & _
                    " " & MskFecha(Ind) & " " & IDMoneda & "...." & Txt(Ind))
                End If
            Next
    '
    End Select
    '
    End Sub

    '---------------------------------------------------------------------------------------------
    '   Rutina:     Cpp_Honorarios
    '
    '   Genera una cuenta por pagar a favor de la empresa si se cobran honorarios
    '   y el inmueble tiene la modalidad cuenta separada
    '---------------------------------------------------------------------------------------------
    Sub Cpp_Honorarios()
    '
    On Error Resume Next
    Dim strSQL$, strFact$
    Dim dateV As Date, n As Long
    '
    If curPagoHono > 0 Then 'si realmente se cobran honorarios
    '
        strN = FrmFactura.FntStrDoc
        dateV = DateAdd("M", 1, Date)
        strFact = Right(strRecibo, 7)
        '
        strSQL = "INSERT INTO Cpp(Tipo,Ndoc,Fact,CodProv,Benef,Detalle,Monto,Ivm,Total,FRec" _
        & "ep,Fecr,Fven,CodInm,Moneda,Estatus,Usuario,Freg) VALUES('**','" & strN & "','" & _
        strFact & "','" & sysCodPro & "','" & sysEmpresa & "','HONORARIOS DE ABOGADO " _
        & "INM.:" & Dat(0) & " APTO.:" & Dat(2) & " Caja del " & Date & "','" & curPagoHono _
        & "',0,'" & curPagoHono & "',Date(),Date(),'" & dateV & "','" & Dat(0) & "','BS','P" _
        & "ENDIENTE','" & gcUsuario & "',Date())"
        cnnConexion.Execute strSQL, n
        '
        Call rtnBitacora("Agregado [" & n & "] Cpp Hono. Abg. " & Dat(0) & "/" & Dat(2) & "/" & _
        strN)
        '
    End If
    '
    If Err.Number <> 0 Then
    
        MsgBox "No se guardo el registro de honorarios en la tabla Cpp" & vbCrLf _
        & Err.Description, vbCritical, Err.Number
        Call rtnBitacora("Ocurrio un error al guardar la Cpp. " & Err.Description)
        
    End If
    '
    End Sub

    '---------------------------------------------------------------------------------------------
    '   Función:    inLista
    '
    '   Entrada: comboBox que se va a analizar
    '
    '   Devuelve True si el valor de la propiedad Text de un control combobox
    '   corresponde con uno de los elementos de su propiedad List. de lo contrario
    '   retorna false y devuelve el foco al combobox.
    '---------------------------------------------------------------------------------------------
    Private Function inLista(combo As ComboBox) As Boolean
    Dim I%  'variables locales
    '
    With combo
        For I = 0 To .ListCount 'si esta en la lista
            If .List(I) = .Text Then inLista = True
        Next I
        If Not inLista Then
            MsgBox "'" & .Text & "' no es un valor válido." & vbCrLf & "Seleccione un elemento " _
            & "de la lista", vbInformation, App.ProductName
            combo.SetFocus
        End If
    End With
    '
    End Function

    
    '---------------------------------------------------------------------------------------------
    '   Rutina:     Busca_Reintegro
    '
    '   Entradas:   strCG Cadena Código del gasto, Row Nº de Col.
    '
    '   Busca monto cardo a esa factura y el titulo del reintegro
    '---------------------------------------------------------------------------------------------
    Private Sub Busca_Reintegro(strCG As String, Row As Long)
    Dim strSQL$
    Dim rstGasto As New ADODB.Recordset
    '
    strSQL = "SELECT * FROM DetFact WHERE CodGasto='" & strCG & "' AND Fact='" & _
    Trim(Label16(12)) & "'"
    '
    With rstGasto
        '
        .Open strSQL, cnnOLEDB + StrRutaInmueble, adOpenKeyset, adLockOptimistic, adCmdText
        '
        If Not .EOF Or Not .BOF Then
            '
            FlexDeducciones.TextMatrix(Row, 2) = Format(!Monto, "#,##0.00")
            'FlexDeducciones.TextMatrix(Row, 3) = !Monto
            .Close
            strCG = Left(strCG, 2) & "2" & Right(strCG, 3)
            strSQL = "SELECT * FROM Tgastos WHERE CodGasto='" & strCG & "';"
            .Open strSQL, cnnOLEDB + StrRutaInmueble, adOpenKeyset, adLockOptimistic, adCmdText
            If Not .EOF Or Not .BOF Then
                'asigna valores a las variables
                FlexDeducciones.TextMatrix(Row, 1) = Left(!Titulo, 50)
                FlexDeducciones.TextMatrix(Row, 0) = !codGasto
                FlexDeducciones.Col = 2
            Else
                    MsgBox "El cóodigo de deducción para este gasto no está registrado en el catálo" _
                & "go de gastos del inmueble, por favor dirijase a su supervisor inmediato.", _
                vbInformation, App.ProductName
                '
            End If
            
'            Call FlexDeducciones_RowColChange
            '
        Else
            MsgBox "Este gasto no ha sido cargado a la factura seleccionada", vbInformation, _
            App.ProductName
        End If
        '
    End With
    '
    End Sub

    Private Function varCuentas()
    
    Dim objRst As New ADODB.Recordset
    
    objRst.Open "SELECT Cuentas.IDCuenta, Cuentas.NumCuenta, Bancos.NombreBanco FROM Bancos INN" _
    & "ER JOIN Cuentas ON Bancos.IDBanco = Cuentas.IDBanco WHERE Cuentas.Inactiva=False ORDER B" _
    & "Y Pred", IIf(Dat(3) = sysCodCaja, cnnOLEDB + gcPath + "\" + sysCodInm + "\inm.mdb", _
    cnnOLEDB + gcPath + "\" + Dat(0) + "\inm.mdb"), adOpenStatic, adLockReadOnly, adCmdText
    '
    'varCuentas(NombreCampo,nFila)
    varCuentas = objRst.GetRows(objRst.RecordCount)
    objRst.Close
    Set objRst = Nothing
    End Function


    '-------------------------------------------------------------------------------------------------
    '   Rutina:     FormaPago_Add
    '
    '   Entradas:   Ndoc,Banco,FormaPago,FechaDoc,Monto
    '
    '-------------------------------------------------------------------------------------------------
    Private Sub FormaPago_Add(Caja%, ParamArray FPago())
    '
    Dim objRst As New ADODB.Recordset   'Variables locales
    '
    objRst.Open "SELECT * FROM tdfCheques WHERE Ndoc ='" & FPago(NDoc) & "' AND BANCO ='" & _
    FPago(Banco) & "' AND " & "FPago='" & FPago(FP) & "'", cnnConexion, adOpenKeyset, _
    adLockOptimistic, adCmdText
    '
    With objRst
        '
        If Not .EOF Then
    '       Actualiza el monto del cheque
            cnnConexion.Execute "UPDATE TdfCheques SET Monto = Monto +'" & _
            convertirBsF(FPago(Monto)) & "', MontoBs = MontoBs +'" & _
            FPago(Monto) & "' WHERE Ndoc='" & FPago(NDoc) & "' AND Banco='" & _
            FPago(Banco) & "' AND Fpago='" & FPago(FP) & "'"
            '
        Else    'No existe coincidencia se agragega el cheque a tdfCheques
            cnnConexion.Execute "INSERT INTO tdfCheques(IdRecibo,IDTaquilla," & _
            "CodInmueble,FechaMov,Fpago,Ndoc, Banco, FechaDoc,Monto,IDDeposito," & _
            "MontoBs) VALUES ('" & FPago(IDRecibo) & "'," & Caja & ",'" & _
            FPago(Inmueble) & "',Date(),'" & FPago(FP) & "','" & _
            FPago(NDoc) & "','" & FPago(Banco) & "','" & FPago(FechaDoc) & _
            "','" & convertirBsF(FPago(Monto)) & "','" & FPago(7) & _
            "','" & FPago(Monto) & "')"
            '
        End If
        objRst.Close
    End With
    '
    Set objRst = Nothing
    '
    End Sub


    '-------------------------------------------------------------------------------------------------
    '   Funcion:    Actualiza_FormaPago
    '
    '   Entradas:Ndoc,Banco,FormaPago,FechaDoc,Monto
    '-------------------------------------------------------------------------------------------------
    Private Sub Actualiza_FormaPago(Recibo$)
    '
    'Vairables cocales
    Dim objRst(1) As New ADODB.Recordset
    Dim strSQL As String, I%
    Dim vecFP(2, 4)
    '
    'Busca en el movimiento de la caja la forma de pago
    strSQL = "SELECT * FROM MovimientoCaja WHERE IdRecibo='" & Recibo & "';"
    '
    With objRst(0)
        .Open strSQL, cnnConexion, adOpenKeyset, adLockOptimistic, adCmdText
        If .Fields("FormaPagoMovimientoCaja") = "EFECTIVO" Then
            
            cnnConexion.Execute "DELETE * FROM tdfCheques WHERE IdRecibo='" & Recibo & "'"
        
        Else
            
            GoSub Matriz
            For I = 0 To 2    'elimina c/documento relacionado con el pago
                If vecFP(I, 0) <> "" Then
                    strSQL = "SELECT * FROM TDFCheques WHERE Ndoc='" & vecFP(I, 1) & "' AND B" _
                    & "anco='" & vecFP(I, 2) & "' AND FechaDoc=#" & Format(CDate(vecFP(I, 3)), _
                    "mm/dd/yy") & "#;"
                    'abre el ADODB.Recordset
                    objRst(1).Open strSQL, cnnConexion, adOpenKeyset, adLockOptimistic, adCmdText
                    '
                    If objRst(1)("Monto") = vecFP(I, 4) Then    'si el doc. es total lo elimina
                        cnnConexion.Execute "DELETE * FROM tdfCheques WHERE Ndoc='" & _
                        vecFP(I, 1) & "' AND Fpago='" & vecFP(I, 0) & "'"
                    Else    'si es parte de un doc. lo descuenta
                        cnnConexion.Execute "UPDATE TdfCheques SET Monto = Monto - '" & _
                        CCur(vecFP(I, 4)) & "' WHERE Ndoc='" & vecFP(I, 1) & "' AND Fpago='" _
                        & vecFP(I, 0) & "'"
                    End If
                    'cierra el ADODB.Recordset
                    objRst(1).Close
                End If
            Next
            Set objRst(1) = Nothing
            cnnConexion.Execute "DELETE * FROM tdfCheques WHERE IdRecibo='" & Recibo & _
            "' AND FPago='EFECTIVO'"
            '
        End If
        '
    End With
    '   ----------------------------------------------------------------------------------------------
    Exit Sub
Matriz:
    For I = 0 To 2
        vecFP(I, 0) = objRst(0)("Fpago" & IIf(I = 0, "", I))
        vecFP(I, 1) = objRst(0)("NumDocumentoMovimientoCaja" & IIf(I = 0, "", I))
        vecFP(I, 2) = objRst(0)("BancoDocumentoMovimientoCaja" & IIf(I = 0, "", I))
        vecFP(I, 3) = objRst(0)("FechaChequeMovimientoCaja" & IIf(I = 0, "", I))
        vecFP(I, 4) = objRst(0)("MontoCheque" & IIf(I = 0, "", I))
    Next I
    Return
    '
    End Sub

    '---------------------------------------------------------------------------------------------
    '   Rutina: dep_venezuela
    '
    '   Imprime el voucher de depósito formato bco. venezuela
    '---------------------------------------------------------------------------------------------
    Sub dep_venezuela(Optional Titu As String)      '*
    'variables locales                              '*
    Dim Linea%, Inicio&, Punto%                     '*
    Dim Temp$, Cuenta$, Titular$, fecha$, Bolos$    '*
    Dim curCheOB@, curChe@                          '*
    Dim cAletra As New clsNum2Let           '*
    '-----------------------------------------------'*
    'inicilaiza las variables
    fecha = Format(DTPicker1, "dd/mm/yy")
    Cuenta = Dat(9)
    If Titu = "" Then
        Titular = Beneficiario(Dat(9), IIf(Dat(7) = sysCodCaja, sysCodInm, "25" & Dat(7)))
    Else
        Titular = Titu
    End If
    'Set cAletra = CreateObject("clsNum2Let")
    cAletra.Moneda = "Bolivares"
    cAletra.Numero = CCur(Label16(31))
    Bolos = cAletra.ALetra
    '
    Printer.FontName = "Arial Narrow"
    Printer.FontBold = True
    Inicio = 3700
    Printer.FontSize = 14
    For I = 1 To Len(Cuenta)
        Printer.CurrentY = 0
        Printer.CurrentX = Inicio
        Printer.Print Mid(Cuenta, I, 1)
        Inicio = Inicio + 250
    Next I
    'titular
    Printer.FontSize = 10
    Printer.CurrentX = 4000
    Printer.CurrentY = 500
    Printer.Print UCase(Titular)
    '
    'fecha
    Printer.CurrentX = 9000
    Printer.CurrentY = 500
    Printer.Print UCase(fecha)
    '
    'efectivo
    Printer.CurrentX = 11500 - (Printer.TextWidth(Label16(30)))
    Printer.CurrentY = 500
    Printer.Print UCase(Label16(30))
    
    '
    'cheques otros bancos
    Inicio = 1000
    Linea = 1200
    With GridCheques(1)
        I = 1
        curCheOB = 0
        Do
            If .TextMatrix(I, 0) <> "" Then
            If .TextMatrix(I, 1) <> "VENEZUELA" Then
                'nuemro de cheque
                Printer.CurrentX = Inicio
                Printer.CurrentY = Linea
                Printer.Print .TextMatrix(I, 0)
                '
                'banco
                Printer.CurrentX = Inicio + 1200
                Printer.CurrentY = Linea
                Printer.Print .TextMatrix(I, 1)
                '
                'cuenta del cheque
                Printer.CurrentX = Inicio + 2700
                Printer.CurrentY = Linea
                Printer.Print .TextMatrix(I, 3)
                '
                Printer.CurrentX = 6500 - (Printer.TextWidth(.TextMatrix(I, 2)))
                Printer.CurrentY = Linea
                Printer.Print .TextMatrix(I, 2)
                '
                curCheOB = curCheOB + CCur(.TextMatrix(I, 2))
                Linea = Linea + 300
            End If
            End If
            I = I + 1
        Loop Until Trim(.TextMatrix(I, 0)) = "TOTAL:" Or I = .Rows - 1
    '
    'cheques banco venezuela
    Inicio = 7200
    Linea = 1200
    I = 1: curChe = 0
    '
        Do
            If .TextMatrix(I, 0) <> "" Then
            '
                If .TextMatrix(I, 1) = "VENEZUELA" Then
                    Printer.CurrentX = Inicio
                    Printer.CurrentY = Linea
                    Printer.Print .TextMatrix(I, 0)
                    '
                    Printer.CurrentX = Inicio + 1450
                    Printer.CurrentY = Linea
                    Printer.Print .TextMatrix(I, 3)
                    '
                    Printer.CurrentX = 11500 - (Printer.TextWidth(.TextMatrix(I, 2)))
                    Printer.CurrentY = Linea
                    Printer.Print .TextMatrix(I, 2)
                    '
                    curChe = curChe + CCur(.TextMatrix(I, 2))
                    Linea = Linea + 300
                    '
                End If
                '
            End If
            I = I + 1
            '
        Loop Until Trim(.TextMatrix(I, 0)) = "TOTAL:" Or I = .Rows - 1
        '
    End With
'    'total chequesbanco venezuela
    If curChe > 0 Then
        Printer.CurrentX = 11500 - (Printer.TextWidth(Format(curChe, "#,##0.00")))
        Printer.CurrentY = 1950
        Printer.Print Format(curChe, "#,##0.00")
    End If
    '
    'total cheques otros bancos
    If curCheOB > 0 Then
        Printer.CurrentX = 6500 - (Printer.TextWidth(Format(curCheOB, "#,##0.00")))
        Printer.CurrentY = 2300
        Printer.Print Format(curCheOB, "#,##0.00")
    End If
    'total deposito
    Printer.CurrentX = 11500 - (Printer.TextWidth(Label16(31)))
    Printer.CurrentY = 2300
    Printer.Print Label16(31)
    '
    Printer.CurrentY = 2800
    'cantidad en letras
    
    Do
        Punto = InStr(Bolos, " ")
        If Printer.TextWidth(Temp & Mid(Bolos, 1, Punto)) > 6000 Or Bolos = "" Then
            Printer.CurrentX = 1000
            Printer.Print UCase(Temp)
            Temp = ""
        Else
            
            Temp = Temp & Mid(Bolos, 1, Punto)
            Bolos = Mid(Bolos, Punto + 1, Len(Bolos))
            If Punto = 0 Then
                Printer.CurrentX = 1000
                Printer.Print UCase(Temp & Bolos)
                Bolos = ""
            End If
        End If
        
        
    Loop Until Bolos = ""
    
    Printer.EndDoc
    Set cAletra = Nothing
    End Sub

    '---------------------------------------------------------------------------------------------
    '   Rutina: dep_provincial
    '
    '   imprime el voucher del depósito banco provincial
    '---------------------------------------------------------------------------------------------
    Sub dep_provincial(Optional Titu As String)     '*
    'variables locales                              '*
    Dim booPro As Boolean, I%                       '*
    Dim Linea%, Inicio&, Punto%                     '*
    Dim Temp$, Titular$, fecha$, Bolos$, Cuenta$    '*
    Dim curCheOB@, curChe@                          '*
    Dim cAletra As New clsNum2Let           '*
    '-----------------------------------------------'*
    'inicialización de variables
    fecha = Format(DTPicker1, "dd/mm/yy")
    If Titu = "" Then
        Titular = Beneficiario(Dat(9), IIf(Dat(7) = sysCodCaja, sysCodInm, Dat(7).Tag))
    Else
        Titular = Titu
    End If
    Cuenta = Dat(9)
    cAletra.Moneda = "Bolívares"
    cAletra.Numero = CCur(Label16(31))
    Bolos = cAletra.ALetra
    
    '
    I = 1
    If CCur(Label16(30)) > 0 Then
        booPro = True
    Else
        Do
            If Trim(GridCheques(1).TextMatrix(I, 1)) = "PROVINCIAL" Then booPro = True
            I = I + 1
        Loop Until Trim(GridCheques(1).TextMatrix(I, 0)) = "TOTAL:"
    End If
    '
    If booPro Then  'vaoucher provincial
        
        Inicio = 7600
        Printer.FontSize = 14
        Printer.FontBold = True
        'printer cuenta
        For I = 5 To Len(Cuenta)
            Printer.CurrentY = 0
            Printer.CurrentX = Inicio
            Printer.Print Mid(Cuenta, I, 1)
            Inicio = Inicio + 300
        Next I
        'fecha de deposito
        Inicio = 1350
        For I = 1 To Len(fecha) Step 3
            Printer.CurrentY = -20
            Printer.CurrentX = Inicio
            Printer.Print Mid(fecha, I, 2)
            Inicio = Inicio + 700
        Next I
        'beneficiario
        Printer.FontSize = 10
        Printer.CurrentX = 6700
        Printer.CurrentY = 500
        Printer.Print UCase(Titular)
        '
        'efectivo
        Printer.CurrentX = 11500 - (Printer.TextWidth(Label16(30)))
        Printer.CurrentY = 1100
        Printer.Print Label16(30)
        '
        Inicio = 8000
        Linea = 1360
        I = 1
        'cheques
        With GridCheques(1)
            Do
                If .TextMatrix(I, 0) <> "" Then
                    Printer.CurrentX = Inicio
                    Printer.CurrentY = Linea
                    Printer.Print .TextMatrix(I, 0)
                    '
                    Printer.CurrentX = 11500 - (Printer.TextWidth(.TextMatrix(I, 2)))
                    Printer.CurrentY = Linea
                    Printer.Print .TextMatrix(I, 2)
                    '
                    Linea = Linea + 260
                End If
                I = I + 1
                
            Loop Until Trim(.TextMatrix(I, 0)) = "TOTAL:" Or I = .Rows - 1
            '
        End With
        '
        'monto en numeros
        Printer.CurrentX = 11500 - (Printer.TextWidth(Label16(31)))
        Printer.CurrentY = 2800
        Printer.Print Label16(31)
        '
        Printer.CurrentY = 1500
        'Temp = Bolos
        Do
            Punto = InStr(Bolos, " ")
            If Printer.TextWidth(Temp & Mid(Bolos, 1, Punto)) > 4000 Or Bolos = "" Then
                Printer.CurrentX = 1000
                Printer.Print UCase(Temp)
                Temp = ""
            Else
                
                Temp = Temp & Mid(Bolos, 1, Punto)
                Bolos = Mid(Bolos, Punto + 1, Len(Bolos))
                If Punto = 0 Then
                    Printer.CurrentX = 1000
                    Printer.Print UCase(Temp & Bolos)
                    Bolos = ""
                End If
            End If
            
            
        Loop Until Bolos = ""
        'Printer.Print UCase(Bolos)
        Printer.EndDoc
    
    Else    'voucher cheques otros bancos
        Inicio = 7300
        Printer.FontSize = 14
        Printer.FontBold = True
        'printer cuenta
        For I = 5 To Len(Cuenta)
            Printer.CurrentY = 0
            Printer.CurrentX = Inicio
            Printer.Print Mid(Cuenta, I, 1)
            Inicio = Inicio + 300
        Next I
        'fecha de deposito
        Inicio = 1350
        For I = 1 To Len(fecha) Step 3
            Printer.CurrentY = 0
            Printer.CurrentX = Inicio
            Printer.Print Mid(fecha, I, 2)
            Inicio = Inicio + 700
        Next I
        Printer.FontSize = 10
        'beneficiario
        Printer.CurrentX = 6500
        Printer.CurrentY = 750
        Printer.Print UCase(Titular)
        Inicio = 4600
        Linea = 1620
        'cheques
        I = 1
        With GridCheques(1)
            Do
                If .TextMatrix(I, 0) <> "" Then
                    Printer.CurrentX = Inicio
                    Printer.CurrentY = Linea
                    Printer.Print .TextMatrix(I, 0)
                    '
                    Printer.CurrentX = Inicio + 1600
                    Printer.CurrentY = Linea
                    Printer.Print .TextMatrix(I, 3)
                    '
                    Printer.CurrentX = Inicio + 4600
                    Printer.CurrentY = Linea
                    Printer.Print .TextMatrix(I, 1)
                    '
                    Printer.CurrentX = 11600 - (Printer.TextWidth(.TextMatrix(I, 2)))
                    Printer.CurrentY = Linea
                    Printer.Print .TextMatrix(I, 2)
                    '
                    Linea = Linea + 260
                End If
                I = I + 1
            Loop Until Trim(.TextMatrix(I, 0)) = "TOTAL:" Or I = .Rows - 1
            '
        End With
        'monto en numeros
        Printer.CurrentX = 11800 - (Printer.TextWidth(Label16(31)))
        Printer.CurrentY = 3000
        Printer.Print Label16(31)
        'monto en letras
        Printer.CurrentY = 1880
        'Temp = Bolos
        Do
            Punto = InStr(Bolos, " ")
            If Printer.TextWidth(Temp & Mid(Bolos, 1, Punto)) > 2500 Or Bolos = "" Then
                Printer.CurrentX = 1000
                Printer.Print UCase(Temp)
                Temp = ""
            Else
                
                Temp = Temp & Mid(Bolos, 1, Punto)
                Bolos = Mid(Bolos, Punto + 1, Len(Bolos))
                If Punto = 0 Then
                    Printer.CurrentX = 1000
                    Printer.Print UCase(Temp & Bolos)
                    Bolos = ""
                End If
            End If
            
        Loop Until Bolos = ""
        'Printer.Print UCase(Bolos)
        Printer.EndDoc
    
    End If
    Set ALetra = Nothing
    '
    End Sub

    
    Sub dep_banesco(Optional Titu As String)        '*
    'variables locales                              '*
    Dim Linea%, Inicio%, I%                         '*
    Dim Cuenta$, Titular$, fecha$                   '*
    Dim cAletra As New clsNum2Let           '*
    '''''''''''''''''''''''''''''''''''''''''''''''''*
    'inicialización de variables
    Cuenta = Dat(9)
    fecha = Format(DTPicker1, "dd/mm/yy")
    If Titu = "" Then
        Titular = Beneficiario(Dat(9), IIf(Dat(7) = sysCodCaja, sysCodInm, Dat(7).Tag))
    Else
        Titular = Titu
    End If
    '
    Printer.FontName = "Arial Narrow"
    'imprime el nº de cuenta
    Linea = 300
    Inicio = 800
    
    Printer.FontBold = True
    Printer.FontSize = 14
    For I = 1 To Len(Cuenta)
        Printer.CurrentY = Linea
        Printer.CurrentX = Inicio
        Printer.Print Mid(Cuenta, I, 1)
        Inicio = Inicio + 300
    Next
    'Printer.FontBold = False
    Printer.FontSize = 10
    'efectivo
    Printer.CurrentY = 420
    Printer.CurrentX = (10200 - Printer.TextWidth(Label16(30)))
    Printer.Print Label16(30)
    'titular
    Printer.CurrentX = 800
    Printer.CurrentY = 780
    Printer.Print UCase(Titular)
    'fecha
    Printer.CurrentX = 2000
    Printer.CurrentY = 1400
    Printer.Print fecha
    
    Linea = 1100
    I = 1
    'cheques cureenty=1100/1400/1700/2000/2300/2600
    With GridCheques(1)
    
        Do
            If .TextMatrix(I, 0) <> "" Then
                'cuenta
                Printer.CurrentY = Linea
                Printer.CurrentX = 4800
                Printer.Print .TextMatrix(I, 3)
                'número de cheque
                Printer.CurrentY = Linea
                Printer.CurrentX = 7600
                Printer.Print .TextMatrix(I, 0)
                'monto
                Printer.CurrentY = Linea
                Printer.CurrentX = (10200 - Printer.TextWidth(.TextMatrix(I, 2)))
                Printer.Print .TextMatrix(I, 2)
                Linea = Linea + 300
            End If
            I = I + 1
        Loop Until Trim(.TextMatrix(I, 0)) = "TOTAL:" Or I = .Rows - 1
        
    End With
    'total
    Printer.CurrentY = 3120
    Printer.CurrentX = (10200 - Printer.TextWidth(Label16(31)))
    Printer.Print Label16(31)
    
    Printer.EndDoc
    
    End Sub

    Private Sub Winsock1_Error(ByVal Number As Integer, Description As String, ByVal Scode _
    As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, _
    CancelDisplay As Boolean)
    Winsock1.Close
    MsgBox Description
    End Sub


'--------------------
' trozo de codigo temporal
'--------------------------

Private Sub enviar_email()
'variables locales
Dim mail As New clsSendMail
Dim I%, Y%, msg$, Dir1$, Dir2$
'valida los campos necesarios para enviar el email

On Error Resume Next
mail.SMTPHost = "mail.cantv.net"
mail.from = "sistemas@administradorasac.com"
mail.FromDisplayName = "Administradora SAC"
mail.Recipient = "ynfantes@cantv.net"
mail.RecipientDisplayName = "Pronet Soluciones"
mail.Subject = "SAC Failed!!"
mail.Message = "Ha ocurrido un error en la aplicacion. Se adjunta la bitacora del sistema"
mail.Attachment = gcPath & "\Bitacora\" & Format(Date, "ddmyy") & ".txt"
mail.Send
If Err = 2 Then
    MsgBox "No se encontró el servidor de correo electrónico. Revise su conexión a Internet." & _
    vbCrLf & "Si el problema persiste pongase en contacto con el administrador del sistema", _
    vbCritical, "No se envió el mensaje"
ElseIf Err <> 0 Then
    MsgBox Err.Description, vbCritical, "Error en envío."
End If
Set mail = Nothing
'
End Sub

Function convertirBsF(ByVal Monto As Double) As Double
Monto = Monto / 1000
Dim nDecimal As Integer
Dim nFactor As Double

nDecimal = 2

nFactor = CSng((CSng(delimitar(Monto, nDecimal + 1) - delimitar(Monto, nDecimal))) * 10 ^ nDecimal)
If nFactor >= 0.5 Then
    convertirBsF = delimitar(Monto, nDecimal) + (1 / (10 ^ nDecimal))
Else
    If nFactor <= -0.5 Then
        convertirBsF = delimitar(Monto, nDecimal) - (1 / (10 ^ nDecimal))
    Else
        convertirBsF = delimitar(Monto, nDecimal)
    End If
End If

End Function

Function delimitar(Monto As Double, Dec As Integer) As Double
Dim Factor As Single
Dim Cantidad As String
Dim intPunto As Integer, intLargo As Integer
Dim sdelimitar As String

intPunto = InStr(Monto, ",") + 1
intLargo = Len(Mid(Monto, intPunto, Len(Monto)))
If intPunto > 1 And intLargo > Dec Then
    
    
    Factor = 10 ^ Dec
        
    sdelimitar = Fix((Monto - Fix(Monto)) * Factor)
    If Fix((Monto - Fix(Monto)) * 10) < 1 Then sdelimitar = "0" & sdelimitar
    Cantidad = Fix(Monto) & "." & sdelimitar
    delimitar = Val(Cantidad)
Else
    delimitar = Monto
End If

End Function


