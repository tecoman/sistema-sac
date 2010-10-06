VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "msadodc.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmAsignaChequera 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Asignación de Chequeras"
   ClientHeight    =   4410
   ClientLeft      =   45
   ClientTop       =   405
   ClientWidth     =   7200
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   ForeColor       =   &H00000000&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4410
   ScaleWidth      =   7200
   Tag             =   "1"
   Begin VB.PictureBox Picture1 
      Height          =   1365
      Left            =   2445
      Picture         =   "FrmAsignaChequera.frx":0000
      ScaleHeight     =   5.438
      ScaleMode       =   4  'Character
      ScaleWidth      =   35.25
      TabIndex        =   18
      Top             =   1335
      Visible         =   0   'False
      Width           =   4290
   End
   Begin MSComctlLib.Toolbar tlbAsigna 
      Align           =   1  'Align Top
      Height          =   480
      Left            =   0
      TabIndex        =   17
      Top             =   0
      Width           =   7200
      _ExtentX        =   12700
      _ExtentY        =   847
      ButtonWidth     =   714
      ButtonHeight    =   688
      Appearance      =   1
      ImageList       =   "ImageList2"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   12
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "First"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Previous"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Next"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "End"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "New"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Save"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Find"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "UNDO"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Delete"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Edit"
            ImageIndex      =   10
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Print"
            ImageIndex      =   11
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   3
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "Todo"
                  Text            =   "Todas"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "Asig"
                  Text            =   "Asignadas"
               EndProperty
               BeginProperty ButtonMenu3 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "Regis"
                  Text            =   "Registradas"
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Close"
            ImageIndex      =   12
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin TabDlg.SSTab stbAsigna 
      Height          =   3645
      Left            =   60
      TabIndex        =   0
      Top             =   600
      Width           =   7035
      _ExtentX        =   12409
      _ExtentY        =   6429
      _Version        =   393216
      Tabs            =   2
      TabHeight       =   520
      ShowFocusRect   =   0   'False
      TabCaption(0)   =   "Datos Generales"
      TabPicture(0)   =   "FrmAsignaChequera.frx":121CE
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "fraAsigna(0)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Lista Chequeras"
      TabPicture(1)   =   "FrmAsignaChequera.frx":121EA
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "fraAsigna(1)"
      Tab(1).Control(1)=   "GridAsigna"
      Tab(1).ControlCount=   2
      Begin VB.Frame fraAsigna 
         Caption         =   "Ver Chequeras:"
         Height          =   885
         Index           =   1
         Left            =   -74790
         TabIndex        =   12
         Top             =   2580
         Width           =   6660
         Begin VB.OptionButton optAsigna 
            Caption         =   "Registradas"
            Height          =   285
            Index           =   2
            Left            =   2925
            TabIndex        =   15
            Top             =   345
            Width           =   1290
         End
         Begin VB.OptionButton optAsigna 
            Caption         =   "Asignadas"
            Height          =   285
            Index           =   1
            Left            =   1560
            TabIndex        =   14
            Top             =   345
            Width           =   1290
         End
         Begin VB.OptionButton optAsigna 
            Caption         =   "Todas"
            Height          =   285
            Index           =   0
            Left            =   345
            TabIndex        =   13
            Top             =   345
            Value           =   -1  'True
            Width           =   1020
         End
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid GridAsigna 
         Height          =   1980
         Left            =   -74835
         TabIndex        =   11
         Tag             =   "400|1500|1800|900|900|600"
         Top             =   510
         Width           =   6690
         _ExtentX        =   11800
         _ExtentY        =   3493
         _Version        =   393216
         Cols            =   6
         FixedCols       =   0
         BackColorFixed  =   -2147483646
         ForeColorFixed  =   -2147483639
         BackColorSel    =   65280
         ForeColorSel    =   -2147483640
         BackColorBkg    =   -2147483636
         GridColor       =   -2147483633
         GridColorFixed  =   -2147483633
         SelectionMode   =   1
         MousePointer    =   99
         FormatString    =   "ID|Nº CUENTA|BANCO|DESDE|HASTA|STATUS"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Comic Sans MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MouseIcon       =   "FrmAsignaChequera.frx":12206
         _NumberOfBands  =   1
         _Band(0).BandIndent=   2
         _Band(0).Cols   =   6
         _Band(0).GridLinesBand=   1
         _Band(0).TextStyleBand=   0
         _Band(0).TextStyleHeader=   0
      End
      Begin VB.Frame fraAsigna 
         Enabled         =   0   'False
         Height          =   2925
         Index           =   0
         Left            =   105
         TabIndex        =   1
         Top             =   435
         Width           =   6735
         Begin MSDataListLib.DataCombo dtcAsigna 
            Bindings        =   "FrmAsignaChequera.frx":12368
            DataField       =   "IDChequera"
            DataSource      =   "adoAsigna(0)"
            Height          =   315
            Left            =   1890
            TabIndex        =   16
            Top             =   383
            Width           =   1170
            _ExtentX        =   2064
            _ExtentY        =   556
            _Version        =   393216
            ListField       =   "IDChequera"
            Text            =   ""
         End
         Begin VB.CheckBox ChkActiva 
            Alignment       =   1  'Right Justify
            Caption         =   "Activar :"
            DataField       =   "Activa"
            DataSource      =   "adoAsigna(0)"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   285
            TabIndex        =   5
            Top             =   2280
            Width           =   1830
         End
         Begin VB.TextBox txtAsigna 
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
            Index           =   0
            Left            =   3240
            Locked          =   -1  'True
            MaxLength       =   15
            TabIndex        =   4
            Top             =   383
            Width           =   3345
         End
         Begin VB.TextBox txtAsigna 
            DataField       =   "Responsable"
            DataSource      =   "adoAsigna(0)"
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
            Left            =   1890
            MaxLength       =   15
            TabIndex        =   3
            Top             =   1328
            Width           =   4725
         End
         Begin VB.TextBox txtAsigna 
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
            Index           =   1
            Left            =   1890
            Locked          =   -1  'True
            TabIndex        =   2
            Top             =   818
            Width           =   4725
         End
         Begin MSMask.MaskEdBox MskFecha 
            Bindings        =   "FrmAsignaChequera.frx":12383
            DataField       =   "Fecha"
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "dd/MM/yyyy"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   8202
               SubFormatType   =   3
            EndProperty
            DataSource      =   "adoAsigna(0)"
            Height          =   315
            Left            =   1890
            TabIndex        =   6
            TabStop         =   0   'False
            Top             =   1793
            Width           =   1125
            _ExtentX        =   1984
            _ExtentY        =   556
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   10
            Format          =   "dd/MM/yyyy"
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin MSAdodcLib.Adodc adoAsigna 
            Height          =   360
            Index           =   1
            Left            =   3975
            Tag             =   $"FrmAsignaChequera.frx":123A5
            Top             =   2115
            Visible         =   0   'False
            Width           =   2655
            _ExtentX        =   4683
            _ExtentY        =   635
            ConnectMode     =   16
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
            Caption         =   "adoGrid"
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
         Begin MSAdodcLib.Adodc adoAsigna 
            Height          =   360
            Index           =   0
            Left            =   3975
            Tag             =   "SELECT * FROM AsignaChequera"
            Top             =   1740
            Visible         =   0   'False
            Width           =   2655
            _ExtentX        =   4683
            _ExtentY        =   635
            ConnectMode     =   16
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
            Caption         =   "AdoAsigna"
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
         Begin MSAdodcLib.Adodc adoAsigna 
            Height          =   360
            Index           =   2
            Left            =   3975
            Tag             =   "SELECT IDChequera FROM Chequera WHERE IDChequera Not In (SELECT IdChequera FROM AsignaChequera) ORDER BY IDChequera"
            Top             =   2475
            Visible         =   0   'False
            Width           =   2655
            _ExtentX        =   4683
            _ExtentY        =   635
            ConnectMode     =   16
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
            Caption         =   "adoChequeras"
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
         Begin VB.Label lblAsigna 
            Caption         =   "Fecha Asignación :"
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
            Left            =   360
            TabIndex        =   10
            Top             =   1845
            Width           =   1305
         End
         Begin VB.Label lblAsigna 
            Caption         =   "Responsable :"
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
            TabIndex        =   9
            Top             =   1380
            Width           =   1305
         End
         Begin VB.Label lblAsigna 
            Caption         =   "Num. Cuenta :"
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
            Left            =   360
            TabIndex        =   8
            Top             =   870
            Width           =   1305
         End
         Begin VB.Label lblAsigna 
            Caption         =   "Cod Chequera :"
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
            Left            =   360
            TabIndex        =   7
            Top             =   450
            Width           =   1305
         End
      End
   End
   Begin MSComctlLib.ImageList ImageList2 
      Left            =   690
      Top             =   315
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
            Picture         =   "FrmAsignaChequera.frx":1251D
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmAsignaChequera.frx":126AD
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmAsignaChequera.frx":1283D
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmAsignaChequera.frx":129CD
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmAsignaChequera.frx":12B5D
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmAsignaChequera.frx":12CED
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmAsignaChequera.frx":12E7D
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmAsignaChequera.frx":1300D
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmAsignaChequera.frx":13121
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmAsignaChequera.frx":132B1
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmAsignaChequera.frx":13441
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmAsignaChequera.frx":135D1
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "FrmAsignaChequera"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
        'SINAI TECH-Modulo Banco-Asignación de Chequera
    '
    Option Explicit
    Dim cnnAsigna_C As ADODB.Connection
    Const iTODO% = 0
    Const iASIGNADAS% = 1
    Const iREGISTRADAS% = 2
    
    '---------------------------------------------------------------------------------------------
    Private Sub dtcAsigna_Click(Area As Integer)    '
    '---------------------------------------------------------------------------------------------
    '
    If Area = 2 Then
    '
        With adoAsigna(1).Recordset
            .MoveFirst
            .Find "IDChequera=" & dtcAsigna
            txtAsigna(0) = !NombreBanco
            txtAsigna(1) = !NumCuenta
        End With
        '
    End If
    '
    End Sub

    '---------------------------------------------------------------------------------------------
    Private Sub dtcAsigna_KeyPress(KeyAscii As Integer) '
    '---------------------------------------------------------------------------------------------
    Call Validacion(KeyAscii, "")
    End Sub

    '---------------------------------------------------------------------------------------------
    Private Sub Form_Load() '
    '---------------------------------------------------------------------------------------------
    'varibales locales
    Dim I%
    '
    CenterForm Me
    
    Set cnnAsigna_C = New ADODB.Connection
    cnnAsigna_C.Open cnnOLEDB & mcDatos
    For I = 0 To 2
        adoAsigna(I).ConnectionString = cnnOLEDB + mcDatos
        adoAsigna(I).RecordSource = adoAsigna(I).Tag
        adoAsigna(I).Refresh
    Next
    Call rtnRepint
    '
    With GridAsigna
        Set .FontFixed = LetraTitulo(LoadResString(527), 6.5, , True)
        Set .Font = LetraTitulo(LoadResString(528), 8)
        .RowHeight(0) = 315
        .ColAlignment(2) = flexAlignCenterCenter
        .ColAlignment(3) = flexAlignCenterCenter
        Call centra_titulo(GridAsigna, True)
    End With
    '
    End Sub

    '---------------------------------------------------------------------------------------------
    Private Sub optAsigna_Click(Index As Integer)   '
    '---------------------------------------------------------------------------------------------
    '
    Select Case Index
        'todas las chequeras
        Case 0: adoAsigna(1).Recordset.Filter = 0
        'Chequeras asignadas
        Case 1: adoAsigna(1).Recordset.Filter = "Activa=True"
        'chequeras registradas
        Case 2: adoAsigna(1).Recordset.Filter = "Activa=False"
    End Select
    Call rtnRepint
    '
    End Sub

    '---------------------------------------------------------------------------------------------
    Private Sub stbAsigna_Click(PreviousTab As Integer) '
    '---------------------------------------------------------------------------------------------
    '
    Select Case stbAsigna.Tab
    '
        Case 0  'Ficha datos generales
            optAsigna_Click 0
            If dtcAsigna <> "" Then dtcAsigna_Click 2
            '
    End Select
    '
    End Sub

    
    '---------------------------------------------------------------------------------------------
    '   Rutina: rtnRepint
    '
    '   Reconstruye el grid y los datos de acuerdo a los parametros
    '   solicitados por el usuario
    '---------------------------------------------------------------------------------------------
    Private Sub rtnRepint()
    'variables locales
    Dim I%, strStatus$
    
    Call rtnLimpiar_Grid(GridAsigna)
    
    If Not adoAsigna(1).Recordset.EOF Then
        '
        With adoAsigna(1).Recordset
            .MoveFirst
            GridAsigna.Rows = .RecordCount + 1
            I = 0
            Do Until .EOF
               I = I + 1
               GridAsigna.TextMatrix(I, 0) = Format(!IDchequera, "00")
               GridAsigna.TextMatrix(I, 1) = !NumCuenta
               GridAsigna.TextMatrix(I, 2) = !NombreBanco
               GridAsigna.TextMatrix(I, 3) = Format(!Desde, "000000")
               GridAsigna.TextMatrix(I, 4) = Format(!Hasta, "000000")
               If IsNull(!activa) Then
                    strStatus = ""
               ElseIf Not !activa Then
                    strStatus = "U"
               Else
                    strStatus = "A"
               End If
               GridAsigna.TextMatrix(I, 5) = strStatus
               .MoveNext
            Loop
        End With
        GridAsigna.Col = 0
        GridAsigna.Row = 1
        GridAsigna.ColSel = GridAsigna.Cols - 1
    End If

    '
    End Sub
    
    '---------------------------------------------------------------------------------------------
    Private Sub tlbAsigna_ButtonClick(ByVal Button As MSComctlLib.Button)   '
    '---------------------------------------------------------------------------------------------
    Dim I As Integer
    '
    With adoAsigna(0).Recordset
    '
        Select Case UCase(Button.Key)
    '
            Case "FIRST"    'mover al inicio
                .MoveFirst
    '
            Case "NEXT" 'mover al siguiente
                .MoveNext
                If .EOF Then .MoveFirst
    '
            Case "PREVIOUS" 'mover al anterior
                .MovePrevious
                If .BOF Then .MoveLast
    '
            Case "END"  'mover al final
                .MoveLast
    '
            Case "NEW"  'Asignar chequera
                'adoAsigna(2).Refresh
                .AddNew
                txtAsigna(0) = ""
                txtAsigna(1) = ""
                MskFecha = Date
                ChkActiva.Value = vbChecked
                fraAsigna(0).Enabled = True
    '
            Case "SAVE" 'guardar registro
                !Hora = Time()
                !Usuario = gcUsuario
                MskFecha.PromptInclude = True
                .Update
                MskFecha.PromptInclude = False
                fraAsigna(0).Enabled = False
                Call rtnBitacora("Update Registrar Chequera " & dtcAsigna)
                MsgBox "Registro Actualizado...", vbInformation, App.ProductName
    '
            Case "EDIT" 'editar
                fraAsigna(0).Enabled = True
    '
            Case "UNDO" 'deshacer los cambios
                fraAsigna(0).Enabled = False
                .CancelUpdate
                MsgBox "Registro Cancelado...", vbInformation, App.ProductName
    '
            Case "PRINT"    'imprimir
                For I = 0 To 2: If optAsigna(I).Value Then Call rtnPrinter(I)
                Next
    '
            Case "CLOSE"    'cerrar formulario
                Unload Me
                Set FrmAsignaChequera = Nothing
        End Select
        
    End With
    '
    End Sub

    '---------------------------------------------------------------------------------------------
    '   Rutina:     rtnPrinter
    '
    '   diseña el reporte de las chequeras según las opciones
    '   señaladas por el usuario {todas/registradas/asignadas}
    '---------------------------------------------------------------------------------------------
    Private Sub rtnPrinter(intOpcion As Integer)
    'variables locales
    Dim strTitulo$
    '
    optAsigna_Click intOpcion
    With adoAsigna(1).Recordset
        If .RecordCount < 1 Then
            MsgBox "No existen datos que imprimir", vbInformation, App.ProductName
        Else
            Printer.PaintPicture Picture1, 0, 0, , , , , , , vbSrcCopy
            Printer.FontSize = 14
            Printer.FontBold = True
            strTitulo = gcCodInm & " - " & gcNomInm
            Printer.ScaleMode = Picture1.ScaleMode
            Printer.Print
            Printer.CurrentX = Printer.ScaleLeft + Picture1.ScaleWidth
            Printer.Print strTitulo
            Printer.FontSize = 12
            strTitulo = "Registro de chequeras " & IIf(intOpcion = 0, "", _
            optAsigna(intOpcion).Caption)
            Printer.CurrentX = Printer.ScaleLeft + Picture1.ScaleWidth
            Printer.Print strTitulo
            Printer.FontSize = 10
            Printer.FontBold = False
            .MoveFirst
            Printer.CurrentY = Picture1.ScaleHeight
            Printer.ScaleMode = 3
            Printer.DrawWidth = 4
            Printer.Line (Printer.ScaleLeft, Printer.CurrentY)-(Printer.ScaleWidth, _
            Printer.CurrentY + (TextHeight(strTitulo) / 2)), , BF
            'Imprime los títulos de las columnas
            Printer.FontBold = True
            Printer.ForeColor = &HFFFFFF
            Printer.CurrentY = Printer.CurrentY - TextHeight(strTitulo)
            Printer.Print Tab(20); "ID."; Tab(30); "N° CUENTA"; Tab(60); "DESDE"; Tab(75); "HASTA"
            Printer.FontBold = False
            Printer.ForeColor = &H80000012
            Printer.CurrentY = Printer.CurrentY + (TextHeight(strTitulo) / 2)
            'Printer.Line (Printer.ScaleLeft, Printer.CurrentY)-(Printer.ScaleWidth, Printer.CurrentY)
            Do Until .EOF
                Printer.Print Tab(21); Format(!IDchequera, "00"); Tab(30); !NumCuenta; Tab(65); _
                Format(!Desde, "000000"); Tab(81); Format(!Hasta, "000000")
                .MoveNext
            Loop
            Printer.EndDoc
        End If
    End With
    End Sub

    Private Sub tlbAsigna_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
    '
    Select Case UCase(ButtonMenu.Key)
        '
        Case "TODO": Call rtnPrinter(iTODO)
        '
        Case "REGIS": Call rtnPrinter(iREGISTRADAS)
        '
        Case "ASIG": Call rtnPrinter(iASIGNADAS)
        '
    End Select
    '
    End Sub

    Private Sub txtAsigna_KeyPress(Index As Integer, KeyAscii As Integer)
    If Index = 2 Then KeyAscii = Asc(UCase(Chr(KeyAscii)))
    End Sub
