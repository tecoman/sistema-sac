VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form FrmNoComunes 
   Caption         =   "Registro de gastos No Comunes"
   ClientHeight    =   15
   ClientLeft      =   60
   ClientTop       =   420
   ClientWidth     =   2700
   ControlBox      =   0   'False
   Icon            =   "FrmNoComunes.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   15
   ScaleWidth      =   2700
   Tag             =   "1"
   WindowState     =   2  'Maximized
   Begin ComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   870
      Left            =   0
      TabIndex        =   11
      Top             =   0
      Width           =   2700
      _ExtentX        =   4763
      _ExtentY        =   1535
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
            Object.Visible         =   0   'False
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
   Begin MSAdodcLib.Adodc adoNoComun 
      Height          =   330
      Index           =   2
      Left            =   5610
      Tag             =   "Select * from Propietarios WHERE Codigo Not Like 'U%' order by codigo"
      Top             =   5400
      Visible         =   0   'False
      Width           =   2325
      _ExtentX        =   4101
      _ExtentY        =   582
      ConnectMode     =   1
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
   Begin MSAdodcLib.Adodc adoNoComun 
      Height          =   330
      Index           =   1
      Left            =   285
      Tag             =   "select * from tgastos where comun = false"
      Top             =   5745
      Visible         =   0   'False
      Width           =   2685
      _ExtentX        =   4736
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   2
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
      Caption         =   "AdoCuentas"
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
   Begin MSAdodcLib.Adodc adoNoComun 
      Height          =   330
      Index           =   0
      Left            =   285
      Top             =   5400
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
      Caption         =   "AdoGastoNocomun"
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
      Height          =   7305
      Left            =   195
      TabIndex        =   7
      Top             =   600
      Width           =   10530
      _ExtentX        =   18574
      _ExtentY        =   12885
      _Version        =   393216
      Tabs            =   2
      Tab             =   1
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
      TabCaption(0)   =   "Datos Generales"
      TabPicture(0)   =   "FrmNoComunes.frx":000C
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "fraNoComun(0)"
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Lista"
      TabPicture(1)   =   "FrmNoComunes.frx":0028
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "DataGrid1"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "fraNoComun(1)"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).ControlCount=   2
      Begin VB.Frame fraNoComun 
         Caption         =   "Ordenar por:"
         Height          =   1335
         Index           =   1
         Left            =   165
         TabIndex        =   15
         Top             =   5565
         Width           =   3840
         Begin VB.OptionButton Opt 
            Caption         =   "Monto"
            Height          =   330
            Index           =   3
            Left            =   2040
            TabIndex        =   19
            Top             =   765
            Width           =   1530
         End
         Begin VB.OptionButton Opt 
            Caption         =   "Concepto"
            Height          =   330
            Index           =   2
            Left            =   2040
            TabIndex        =   18
            Top             =   330
            Width           =   1530
         End
         Begin VB.OptionButton Opt 
            Caption         =   "Cod. Gasto"
            Height          =   330
            Index           =   1
            Left            =   360
            TabIndex        =   17
            Top             =   765
            Width           =   1530
         End
         Begin VB.OptionButton Opt 
            Caption         =   "Apartamento"
            Height          =   330
            Index           =   0
            Left            =   360
            TabIndex        =   16
            Top             =   330
            Value           =   -1  'True
            Width           =   1530
         End
      End
      Begin MSDataGridLib.DataGrid DataGrid1 
         Height          =   5055
         Left            =   180
         TabIndex        =   10
         Top             =   450
         Width           =   9810
         _ExtentX        =   17304
         _ExtentY        =   8916
         _Version        =   393216
         AllowUpdate     =   -1  'True
         HeadLines       =   2
         RowHeight       =   19
         RowDividerStyle =   4
         FormatLocked    =   -1  'True
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Comic Sans MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "LISTADO DE GASTOS NO COMUNES"
         ColumnCount     =   6
         BeginProperty Column00 
            DataField       =   "CodApto"
            Caption         =   "CodApto."
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
            DataField       =   "CodGasto"
            Caption         =   "CodGasto"
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
            DataField       =   "Concepto"
            Caption         =   "Concepto"
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
            DataField       =   "Monto"
            Caption         =   "Monto"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   "#,##0.00 "
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   8202
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column04 
            DataField       =   "Periodo"
            Caption         =   "Periodo"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   "mmm-yyyy"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   8202
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column05 
            DataField       =   "Fecha"
            Caption         =   "Fecha"
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
               Alignment       =   2
               ColumnWidth     =   750,047
            EndProperty
            BeginProperty Column01 
               Alignment       =   2
               ColumnWidth     =   870,236
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   4050,142
            EndProperty
            BeginProperty Column03 
               Alignment       =   1
            EndProperty
            BeginProperty Column04 
               ColumnWidth     =   824,882
            EndProperty
            BeginProperty Column05 
               ColumnWidth     =   1200,189
            EndProperty
         EndProperty
      End
      Begin VB.Frame fraNoComun 
         Enabled         =   0   'False
         Height          =   2730
         Index           =   0
         Left            =   -74835
         TabIndex        =   8
         Top             =   420
         Width           =   10185
         Begin VB.ComboBox cmbNC 
            Height          =   315
            Index           =   1
            Left            =   3270
            Style           =   2  'Dropdown List
            TabIndex        =   5
            Top             =   1538
            Width           =   1065
         End
         Begin VB.ComboBox cmbNC 
            Height          =   315
            Index           =   0
            ItemData        =   "FrmNoComunes.frx":0044
            Left            =   2190
            List            =   "FrmNoComunes.frx":006C
            Style           =   2  'Dropdown List
            TabIndex        =   4
            Top             =   1545
            Width           =   1065
         End
         Begin VB.TextBox txtNoComun 
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
            Height          =   315
            Left            =   2190
            TabIndex        =   6
            Top             =   2078
            Width           =   1440
         End
         Begin MSDataListLib.DataCombo cmbNoComun 
            Bindings        =   "FrmNoComunes.frx":00AC
            Height          =   315
            Index           =   2
            Left            =   2190
            TabIndex        =   2
            Top             =   983
            Width           =   1050
            _ExtentX        =   1852
            _ExtentY        =   556
            _Version        =   393216
            ListField       =   "CodGasto"
            BoundColumn     =   "Titulo"
            Text            =   " "
         End
         Begin MSDataListLib.DataCombo cmbNoComun 
            Bindings        =   "FrmNoComunes.frx":00C8
            Height          =   315
            Index           =   3
            Left            =   3390
            TabIndex        =   3
            Top             =   983
            Width           =   4380
            _ExtentX        =   7726
            _ExtentY        =   556
            _Version        =   393216
            ListField       =   "Titulo"
            BoundColumn     =   "CodGasto"
            Text            =   " "
         End
         Begin MSDataListLib.DataCombo cmbNoComun 
            Bindings        =   "FrmNoComunes.frx":00E4
            Height          =   315
            Index           =   0
            Left            =   2190
            TabIndex        =   0
            Top             =   443
            Width           =   1035
            _ExtentX        =   1826
            _ExtentY        =   556
            _Version        =   393216
            ListField       =   "Codigo"
            BoundColumn     =   "Nombre"
            Text            =   " "
         End
         Begin MSDataListLib.DataCombo cmbNoComun 
            Bindings        =   "FrmNoComunes.frx":0100
            Height          =   315
            Index           =   1
            Left            =   3360
            TabIndex        =   1
            Top             =   443
            Width           =   4380
            _ExtentX        =   7726
            _ExtentY        =   556
            _Version        =   393216
            ListField       =   "Nombre"
            BoundColumn     =   "Codigo"
            Text            =   " "
         End
         Begin VB.Label Label1 
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
            Index           =   3
            Left            =   345
            TabIndex        =   14
            Top             =   2130
            Width           =   1290
         End
         Begin VB.Label Label1 
            Caption         =   "Período:(mm-aaaa)"
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
            Left            =   360
            TabIndex        =   13
            Top             =   1590
            Width           =   1680
         End
         Begin VB.Label Label1 
            Caption         =   "Cod.Gasto:"
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
            Left            =   345
            TabIndex        =   12
            Top             =   1035
            Width           =   1290
         End
         Begin VB.Label Label1 
            Caption         =   "Número Apto. :"
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
            Left            =   375
            TabIndex        =   9
            Top             =   495
            Width           =   1290
         End
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
            Picture         =   "FrmNoComunes.frx":011C
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmNoComunes.frx":029E
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmNoComunes.frx":0420
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmNoComunes.frx":05A2
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmNoComunes.frx":0724
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmNoComunes.frx":08A6
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmNoComunes.frx":0A28
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmNoComunes.frx":0BAA
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmNoComunes.frx":0D2C
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmNoComunes.frx":0EAE
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmNoComunes.frx":1030
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmNoComunes.frx":11B2
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "FrmNoComunes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
    'variables globles a nivel de módulo
    Dim cnnNoComun As New ADODB.Connection
    Dim rstFactura As New ADODB.Recordset
    Dim datPeriodo As Date
    
    Private Sub cmbNC_KeyPress(Index As Integer, KeyAscii As Integer)
    On Error Resume Next
    If Index = 0 Then cmbNC(1).SetFocus
    If Index = 1 Then txtNoComun.SetFocus
    End Sub

    '---------------------------------------------------------------------------------------------
    '   Contiene los procedimientos que responden al hacer click sobre un
    '   elemento de la lista desplegable de cualquier Datacombo de la matriz
    '   de controles
    '---------------------------------------------------------------------------------------------
    Private Sub cmbNoComun_Click(Index As Integer, Area As Integer)
    '
    If Area = 2 Then    'Click en la lista despleglabe
    '
        Select Case Index
        '
            Case 0, 1 'Codigo de Apartamento,'Nombre de Propietario
            '----------------
                cmbNoComun(IIf(Index = 0, 1, 0)) = cmbNoComun(Index).BoundText
                cmbNoComun(2).SetFocus
            
            Case 2, 3 'CodGasto,'Titulo del Gasto
            '--------------------
                On Error Resume Next
                cmbNoComun(IIf(Index = 2, 3, 2)) = cmbNoComun(Index).BoundText
                cmbNC(0).SetFocus
        End Select
        '
    End If
    '
    End Sub

    Private Sub cmbNoComun_KeyPress(Index As Integer, KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then
        If Index = 2 Or Index = 3 Then
            Call RtnBuscarCodigo(Index)
        Else
            Call cmbNoComun_Click(Index, 2)
        End If
    End If
    End Sub



    Private Sub DataGrid1_DblClick(): Call Llenar_Controls
    End Sub

    '---------------------------------------------------------------------------------------------
    Private Sub Form_Load() '
    '---------------------------------------------------------------------------------------------
    'Conecta los control Ado al origen de datos
    cnnNoComun.CursorLocation = adUseClient
    cnnNoComun.Open cnnOLEDB & mcDatos
    If Not gcCodInm = sysCodInm Then
        rstFactura.Open "SELECT MAX(Periodo) FROM Factura WHERE Fact Not Like 'CH%' or IsNull(F" _
        & "act)", cnnNoComun, adOpenKeyset, adLockOptimistic
        datPeriodo = IIf(IsNull(rstFactura.Fields(0)), DateAdd("YYYY", -10, Date), rstFactura.Fields(0))
        rstFactura.Close
        '
        For i = 1 To 2
        
           adoNoComun(i).ConnectionString = cnnNoComun.ConnectionString
           adoNoComun(i).RecordSource = adoNoComun(i).Tag
           adoNoComun(i).Refresh
           cmbNC(1).AddItem (Format(DateAdd("yyyy", i - 1, Date), "yyyy"))
           
        Next
        adoNoComun(0).ConnectionString = cnnNoComun.ConnectionString
        adoNoComun(0).RecordSource = "SELECT * FROM GastoNoComun WHERE Periodo>=#" & _
        Format(DateAdd("m", 1, datPeriodo), "mm/dd/yyyy") & "# ORDER BY CodApto"
        adoNoComun(0).Refresh
        Set DataGrid1.DataSource = adoNoComun(0)
        '
    End If
    rstFactura.Open "SELECT * FROM TGASTOS;", cnnNoComun, adOpenStatic, adLockReadOnly
    Set DataGrid1.HeadFont = LetraTitulo(LoadResString(527), 7.5, True)
    Set DataGrid1.Font = LetraTitulo(LoadResString(528), 8)
    '
    End Sub


    '---------------------------------------------------------------------------------------------
    Private Sub Form_Resize()   '
    '---------------------------------------------------------------------------------------------
    '
    If WindowState <> vbMinimized Then
        SSTab1.Left = (ScaleWidth / 2) - (SSTab1.Width / 2)
        
    End If
    '
    End Sub

    '---------------------------------------------------------------------------------------------
    Private Sub Form_Unload(Cancel As Integer)
    '---------------------------------------------------------------------------------------------
    'Rutina destructor
    cnnNoComun.Close
    Set cnnNoComun = Nothing
    Set rstFactura = Nothing
    Set FrmNoComunes = Nothing
    End Sub

    Private Sub Opt_Click(Index As Integer)
    adoNoComun(0).Recordset.Sort = DataGrid1.Columns(Index).DataField & ", CodGasto"
    End Sub

    Private Sub SSTab1_Click(PreviousTab As Integer)
    If SSTab1.Tab = 0 Then
        SSTab1.Height = 3480
    Else
       SSTab1.Height = Me.ScaleHeight - 1000
       adoNoComun(0).Recordset.Requery
    End If
    End Sub

    '---------------------------------------------------------------------------------------------
    '   Responde a los eventos click en la barra de herramientas
    '---------------------------------------------------------------------------------------------
    Private Sub Toolbar1_ButtonClick(ByVal Button As ComctlLib.Button)
    'variables locales
    'Dim Marca As Variant
    Dim Fila As Long
    Dim msg As String
    '
    With adoNoComun(0).Recordset
        
        Select Case UCase(Button.Key)
          
            Case "FIRST"    'Primer Registro
            '----------------
                If Not .EOF Or Not .BOF Then
                    .MoveFirst
                    Call Llenar_Controls
                End If
                
            Case "PREVIOUS"     'Registro Previo
            '----------------
                If Not .EOF Or Not .BOF Then
                    .MovePrevious
                    If .BOF Then .MoveLast
                    Call Llenar_Controls
                End If
            Case "NEXT"     'Siguiente Registro
            '----------------
                If Not .EOF Or Not .BOF Then
                    .MoveNext
                    If .EOF Then .MoveFirst
                    Call Llenar_Controls
                End If
                
            Case "END"  'Último Registro
            '----------------
                If Not .EOF Or Not .BOF Then
                    .MoveLast
                    Call Llenar_Controls
                End If
            Case "NEW"  'Nuevo Registro
            '----------------
                SSTab1.Tab = 0
                fraNoComun(0).Enabled = True
                Call Limpiar_Controls
                cmbNoComun(0).SetFocus
                
            Case "SAVE" 'Guardar Registr0
            '----------------
                If ftnValidar = True Then Exit Sub
                fraNoComun(0).Enabled = False
                cnnNoComun.Execute "INSERT INTO GastoNoComun(PF,CodApto,CodGasto,Concepto,Monto" _
                & ",Periodo,Fecha,Hora,Usuario) VALUES (False,'" & cmbNoComun(0) & "','" _
                & cmbNoComun(2) & "','" & cmbNoComun(3) & "','" & CCur(txtNoComun) & "','" _
                & "01/" & cmbNC(0) & "/" & cmbNC(1) & "',Date() ,Time() ,'" & gcUsuario & "');"
                
                Call Limpiar_Controls
                MsgBox "Registro Actualizado...", vbInformation, App.ProductName
                .Requery
                
            Case "FIND" 'Buscar Registro
            '----------------
                             
            Case "UNDO" 'Cancelar Registro
            '----------------
                fraNoComun(0).Enabled = False
                Call Limpiar_Controls
                
            Case "DELETE"   'Eliminar Registro
            '----------------
                On Error GoTo fallo:
                msg = "¿Desea eliminar el(los) registro(s) seleccinado(s)?"
                
                If Respuesta(msg) Then
                
                    cnnNoComun.BeginTrans
                    
'                    For Each Marca In DataGrid1.SelBookmarks
'
'                        .AbsolutePosition = Marca
                        'elimina el registro activo
                        '
                        cnnNoComun.Execute "DELETE * FROM GastoNoComun WHERE PF=False AND CodAp" _
                        & "to='" & !CodApto & "' AND CodGasto='" & !codGasto & "' AND Concepto=" _
                        & "'" & !Concepto & "' AND Monto=CCur('" & !Monto & "') AND Periodo=#" _
                        & Format(!Periodo, "mm/dd/yy") & "# AND Usuario='" & !Usuario & _
                        "' AND Hora='" & !Hora & "';"
                        'Escribir en la bitacora del sistema la acción eliminar
                        Fila = Fila + 1
                        
'                    Next
                    If Fila > 1 Then
                        msg = "(" & Fila & ") Registros Eliminados.."
                    Else
                        msg = "Registro Eliminado.."
                    End If
                    cnnNoComun.CommitTrans
                    MsgBox msg, vbInformation, App.ProductName
                    .Requery
                    Set DataGrid1.DataSource = Nothing
                    Set DataGrid1.DataSource = adoNoComun(0)
                    '
                End If
fallo:
                If Err.Number <> 0 Then
                    MsgBox Err.Description, vbInformation, App.ProductName
                    cnnNoComun.RollbackTrans
                End If
                           
            Case "CLOSE"    'Descargar Formulario
            '----------------
                Unload Me
                
                
            Case "PRINT"    'Imprimir Registro
            '----------------
                Dim strSql As String
                strSql = "SELECT * FROM gastoNoComun WHERE Periodo=#" & _
                Format(DateAdd("m", 1, datPeriodo), "mm/dd/yyyy") & "#"
                Call rtnGenerator(mcDatos, strSql, "gcmacGNC")
                mcTitulo = "Reporte de Gastos No Comunes"
                mcReport = "fact_lgnc.Rpt"
                mcOrdCod = "+ {gcmacGNC.CodApto}"
                mcOrdAlfa = "+ {gcmacGNC.Concepto}"
                mcCrit = ""
                FrmReport.Show
    
        End Select
        '
     End With
     
    End Sub

    '---------------------------------------------------------------------------------------------
    Private Sub Limpiar_Controls()  '
    '---------------------------------------------------------------------------------------------
    'Limpia los controles de captura
    For i = 0 To 3
        cmbNoComun(i) = ""
    Next
    cmbNC(0).ListIndex = -1
    cmbNC(1).ListIndex = -1
    txtNoComun = ""
    '
    End Sub

    
    '---------------------------------------------------------------------------------------------
    Private Sub Llenar_Controls()
    '---------------------------------------------------------------------------------------------
    'Llena el contenido de los controles de captura con la información del grid
    With DataGrid1
        cmbNoComun(0) = .Columns(0)
        cmbNoComun(1) = Buscar_Propietario(.Columns(0))
        cmbNoComun(2) = .Columns(1)
        cmbNoComun(3) = .Columns(2)
        cmbNC(0) = cmbNC(0).List(Month(.Columns(4)) - 1)
        If Year(.Columns(4)) > Year(Date) Then
            cmbNC(1) = cmbNC(1).List(0)
        ElseIf Year(.Columns(4)) = Year(Date) Then
            cmbNC(1) = cmbNC(1).List(1)
        Else
            cmbNC(1) = cmbNC(1).List(1)
        End If
        txtNoComun = Format(.Columns(3), "#,##0.00")
    End With
    '
    End Sub

    '---------------------------------------------------------------------------------------------
    '   Funcion:    Buscar_Propietario
    '
    '   Entrada:    {strPro} código de apartamento
    '
    '   Salida:     Nombre del propietario correspondiente al apartamento strprop
    '---------------------------------------------------------------------------------------------
    Private Function Buscar_Propietario(strPro) As String
    '
    With adoNoComun(2).Recordset
        .MoveFirst
        .Find "Codigo ='" & strPro & "'"
        If Not .EOF Then Buscar_Propietario = !Nombre
    End With
    '
    End Function

    '---------------------------------------------------------------------------------------------
    '
    '   Funcion:     ftnValidar
    '
    '   Devuelve True si se obtienen todos los datos necesarios para guardar un
    '   registro y si c/u corresponde con el tipo de dato, de lo contrario
    '   devuelve: False
    '---------------------------------------------------------------------------------------------
    Private Function ftnValidar() As Boolean
    'Variable local
    Dim strMsg As String
    
    If cmbNoComun(0) = "" Then
        strMsg = "- Falta Número de Apartamento del Propietario."
    ElseIf cmbNoComun(1) = "" Then
        strMsg = "- Falta Nombre del Propietario del apartamento " & cmbNoComun(0)
    ElseIf cmbNoComun(2) = "" Then
        strMsg = "- Falta el Código del Gasto."
    ElseIf cmbNoComun(3) = "" Then
        strMsg = "- Falta la descripción del Gasto."
    ElseIf txtNoComun = "" Then
        strMsg = "- Falta el monto del gasto."
    ElseIf datPeriodo >= "01/" & cmbNC(0) & "/" & cmbNC(1) Then
        strMsg = "- Período señalado ya está facturado..."
    End If
    If Not strMsg = "" Then ftnValidar = MsgBox(strMsg, vbInformation, "Error...")
    '
    End Function


    '---------------------------------------------------------------------------------------------
    Private Sub txtNoComun_KeyPress(KeyAscii As Integer)    '
    '---------------------------------------------------------------------------------------------
    If KeyAscii = 46 Then KeyAscii = 44 'CONVIERTE EL PUNTO EN COMA
    '
    Call Validacion(KeyAscii, "-1234567890,")    'Permite solo datos numéricos
    '
    If KeyAscii = 13 Then
    
        txtNoComun = Format(txtNoComun, "#,##0.00") 'Da formato al campo
        If Respuesta("¿Seguro desea agregar este gasto?") Then
        
            Dim boton As ComctlLib.Button
            
            Set boton = Toolbar1.Buttons("Save")
            Toolbar1_ButtonClick boton
            If Respuesta("¿Desea agregar otro gasto?") Then
                Set boton = Toolbar1.Buttons("New")
                Toolbar1_ButtonClick boton
            End If
            
        End If
        
    End If
        '
    End Sub


    Private Sub RtnBuscarCodigo(X As Integer)
    'variables locales
    Dim strCriterio As String
    '
    If X = 2 Then
        strCriterio = "CodGasto='" & cmbNoComun(2) & "'"
    Else
        strCriterio = "Titulo='" & cmbNoComun(3) & "'"
    End If
    '
    With rstFactura
        .Find strCriterio
        If .EOF Then
            .MoveFirst
            .Find strCriterio
        End If
        If Not .EOF Then
            cmbNoComun(2) = !codGasto
            cmbNoComun(3) = IIf(IsNull(!Titulo), "", !Titulo)
        Else
            MsgBox "No existe coincidencia", vbOKOnly + vbInformation
            cmbNoComun(IIf(X = 2, 3, 2)) = ""
            cmbNoComun(X).SetFocus
        End If
    End With
    '
    End Sub
