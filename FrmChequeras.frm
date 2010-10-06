VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form FrmChequeras 
   Caption         =   "Registro de Chequeras"
   ClientHeight    =   3135
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   4680
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   3135
   ScaleWidth      =   4680
   Tag             =   "1"
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
            Object.ToolTipText     =   "Salir"
            Object.Tag             =   ""
            ImageIndex      =   10
         EndProperty
         BeginProperty Button11 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Print"
            Object.ToolTipText     =   "Cerrar"
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
   End
   Begin VB.Frame fraChequera 
      Caption         =   "Ordenar por:"
      Height          =   2100
      Index           =   1
      Left            =   8445
      TabIndex        =   18
      Top             =   1125
      Width           =   1860
      Begin VB.OptionButton optOrden 
         Caption         =   "Banco"
         Height          =   225
         Index           =   3
         Left            =   240
         TabIndex        =   22
         Tag             =   "NombreBanco, Desde"
         Top             =   1215
         Width           =   1485
      End
      Begin VB.OptionButton optOrden 
         Caption         =   "Chequera"
         Height          =   225
         Index           =   0
         Left            =   240
         TabIndex        =   21
         Tag             =   "IDChequera"
         Top             =   525
         Value           =   -1  'True
         Width           =   1485
      End
      Begin VB.OptionButton optOrden 
         Caption         =   "Cuenta"
         Height          =   225
         Index           =   1
         Left            =   240
         TabIndex        =   20
         Tag             =   "NumCuenta, Desde"
         Top             =   870
         Width           =   1485
      End
      Begin VB.OptionButton optOrden 
         Caption         =   "Fecha Registro"
         Height          =   225
         Index           =   2
         Left            =   240
         TabIndex        =   19
         Tag             =   "Fecha, IDChequera"
         Top             =   1560
         Width           =   1485
      End
   End
   Begin MSAdodcLib.Adodc adochequera 
      Height          =   345
      Index           =   2
      Left            =   5220
      Tag             =   $"FrmChequeras.frx":0000
      Top             =   2595
      Visible         =   0   'False
      Width           =   2400
      _ExtentX        =   4233
      _ExtentY        =   609
      ConnectMode     =   3
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   5
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
      BackColor       =   -2147483646
      ForeColor       =   -2147483639
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
      Caption         =   "AdoGRID"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.Frame fraChequera 
      BackColor       =   &H80000004&
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
      Height          =   2835
      Index           =   0
      Left            =   1403
      TabIndex        =   1
      Top             =   570
      Width           =   9075
      Begin VB.TextBox txtchequera 
         DataField       =   "Hasta"
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "000000"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   8202
            SubFormatType   =   0
         EndProperty
         DataSource      =   "adochequera(0)"
         Height          =   315
         Index           =   2
         Left            =   2055
         Locked          =   -1  'True
         MaxLength       =   6
         TabIndex        =   6
         Text            =   " "
         Top             =   2205
         Width           =   1125
      End
      Begin VB.TextBox txtchequera 
         DataField       =   "DESDE"
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "000000"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   8202
            SubFormatType   =   0
         EndProperty
         DataSource      =   "adochequera(0)"
         Height          =   315
         Index           =   1
         Left            =   2055
         MaxLength       =   6
         TabIndex        =   5
         Text            =   " "
         Top             =   1830
         Width           =   1125
      End
      Begin VB.TextBox txtchequera 
         DataField       =   "IDchequera"
         DataSource      =   "adochequera(0)"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   315
         Index           =   0
         Left            =   2055
         MaxLength       =   5
         TabIndex        =   4
         Top             =   270
         Width           =   1305
      End
      Begin VB.ComboBox cmbchequera 
         Height          =   315
         ItemData        =   "FrmChequeras.frx":00C2
         Left            =   4890
         List            =   "FrmChequeras.frx":00CC
         TabIndex        =   3
         Top             =   1830
         Width           =   810
      End
      Begin VB.CommandButton cmdchequera 
         Height          =   270
         Left            =   3090
         Picture         =   "FrmChequeras.frx":00D8
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   1485
         Width           =   255
      End
      Begin MSDataListLib.DataCombo dtcChequera 
         Bindings        =   "FrmChequeras.frx":01DA
         DataField       =   "NombreBanco"
         Height          =   330
         Index           =   1
         Left            =   2055
         TabIndex        =   7
         Top             =   1065
         Width           =   4740
         _ExtentX        =   8361
         _ExtentY        =   582
         _Version        =   393216
         BackColor       =   -2147483643
         ListField       =   "NombreBanco"
         BoundColumn     =   "NumCuenta"
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
      Begin MSMask.MaskEdBox mskChequera 
         Bindings        =   "FrmChequeras.frx":01F7
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
         DataSource      =   "adochequera(0)"
         Height          =   315
         Left            =   2055
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   1455
         Width           =   1320
         _ExtentX        =   2328
         _ExtentY        =   556
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   10
         Format          =   "dd/MM/yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSAdodcLib.Adodc adochequera 
         Height          =   345
         Index           =   0
         Left            =   3810
         Tag             =   "Chequera"
         Top             =   1650
         Visible         =   0   'False
         Width           =   2400
         _ExtentX        =   4233
         _ExtentY        =   609
         ConnectMode     =   16
         CursorLocation  =   3
         IsolationLevel  =   -1
         ConnectionTimeout=   15
         CommandTimeout  =   5
         CursorType      =   1
         LockType        =   3
         CommandType     =   2
         CursorOptions   =   0
         CacheSize       =   50
         MaxRecords      =   0
         BOFAction       =   0
         EOFAction       =   0
         ConnectStringType=   1
         Appearance      =   1
         BackColor       =   -2147483646
         ForeColor       =   -2147483639
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
         Caption         =   "AdoChequera"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _Version        =   393216
      End
      Begin MSDataListLib.DataCombo dtcChequera 
         Bindings        =   "FrmChequeras.frx":0219
         DataField       =   "NumCuenta"
         Height          =   330
         Index           =   0
         Left            =   2055
         TabIndex        =   9
         Top             =   660
         Width           =   4740
         _ExtentX        =   8361
         _ExtentY        =   582
         _Version        =   393216
         BackColor       =   -2147483643
         ListField       =   "NumCuenta"
         BoundColumn     =   "NombreBanco"
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
      Begin MSAdodcLib.Adodc adochequera 
         Height          =   345
         Index           =   1
         Left            =   3825
         Tag             =   $"FrmChequeras.frx":0236
         Top             =   2385
         Visible         =   0   'False
         Width           =   2400
         _ExtentX        =   4233
         _ExtentY        =   609
         ConnectMode     =   16
         CursorLocation  =   3
         IsolationLevel  =   -1
         ConnectionTimeout=   15
         CommandTimeout  =   5
         CursorType      =   2
         LockType        =   3
         CommandType     =   1
         CursorOptions   =   0
         CacheSize       =   50
         MaxRecords      =   0
         BOFAction       =   0
         EOFAction       =   0
         ConnectStringType=   1
         Appearance      =   1
         BackColor       =   -2147483646
         ForeColor       =   -2147483639
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
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _Version        =   393216
      End
      Begin VB.Label lblchequera 
         Caption         =   "Desde Cheque :"
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
         Left            =   300
         TabIndex        =   16
         Top             =   1882
         Width           =   1335
      End
      Begin VB.Label lblchequera 
         Caption         =   "Cuenta N°:"
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
         Left            =   300
         TabIndex        =   15
         Top             =   742
         Width           =   1335
      End
      Begin VB.Label lblchequera 
         Caption         =   "Hasta Cheque :"
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
         Left            =   300
         TabIndex        =   14
         Top             =   2257
         Width           =   1335
      End
      Begin VB.Label lblchequera 
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
         Index           =   1
         Left            =   300
         TabIndex        =   13
         Top             =   1110
         Width           =   1335
      End
      Begin VB.Label lblchequera 
         Caption         =   "Cod. Chequera :"
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
         Left            =   300
         TabIndex        =   12
         Top             =   367
         Width           =   1335
      End
      Begin VB.Label lblchequera 
         Caption         =   "Fecha Registro:"
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
         Left            =   300
         TabIndex        =   11
         Top             =   1507
         Width           =   1335
      End
      Begin VB.Label lblchequera 
         AutoSize        =   -1  'True
         Caption         =   "Num. Cheques :"
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
         Left            =   3450
         TabIndex        =   10
         Top             =   1882
         Width           =   1320
      End
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid msgChequera 
      Height          =   3600
      Left            =   1380
      TabIndex        =   17
      Tag             =   "500|2000|2000|1200|1200|1200"
      Top             =   3630
      Width           =   9135
      _ExtentX        =   16113
      _ExtentY        =   6350
      _Version        =   393216
      ForeColor       =   -2147483646
      Cols            =   6
      FixedCols       =   0
      BackColorFixed  =   -2147483646
      ForeColorFixed  =   -2147483639
      BackColorSel    =   65280
      ForeColorSel    =   0
      AllowBigSelection=   0   'False
      FocusRect       =   0
      ScrollBars      =   2
      SelectionMode   =   1
      AllowUserResizing=   1
      MousePointer    =   99
      FormatString    =   "ID|CUENTA|BANCO|DESDE|HASTA|ULTIMO"
      BandDisplay     =   1
      BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MouseIcon       =   "FrmChequeras.frx":02CA
      _NumberOfBands  =   1
      _Band(0).Cols   =   6
      _Band(0).GridLinesBand=   0
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   0
   End
   Begin ComctlLib.ImageList ImageList1 
      Left            =   120
      Top             =   3750
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
            Picture         =   "FrmChequeras.frx":042C
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmChequeras.frx":05AE
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmChequeras.frx":0730
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmChequeras.frx":08B2
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmChequeras.frx":0A34
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmChequeras.frx":0BB6
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmChequeras.frx":0D38
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmChequeras.frx":0EBA
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmChequeras.frx":103C
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmChequeras.frx":11BE
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmChequeras.frx":1340
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmChequeras.frx":14C2
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "FrmChequeras"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

    Dim ChqrsCnn As New ADODB.Connection
    Dim datAhora As Date

    Private Sub CmbNumCheques_Click(): TxtHasta = CDbl(TxtDesde) + CInt(CmbNumCheques)
    End Sub

    '---------------------------------------------------------------------------------------------
    Private Sub cmbchequera_Click() '
    '---------------------------------------------------------------------------------------------
    '
        txtchequera(2) = Format(CLng(txtchequera(1)) + CLng(cmbchequera - 1), "000000")
        txtchequera(1) = Format(txtchequera(1), "000000")
    '
    End Sub

    '---------------------------------------------------------------------------------------------
    Private Sub cmbchequera_KeyPress(KeyAscii As Integer)   '
    '---------------------------------------------------------------------------------------------
    '
    Call Validacion(KeyAscii, "1234567890")
    If KeyAscii = 13 Then
        txtchequera(2) = Format(CLng(txtchequera(1)) + CLng(cmbchequera - 1), "000000")
        txtchequera(1) = Format(txtchequera(1), "000000")
    End If
    '
    End Sub

    '---------------------------------------------------------------------------------------------
    Private Sub dtcChequera_Click(Index As Integer, Area As Integer)    '
    '---------------------------------------------------------------------------------------------
    '
    If Area = 2 Then
        Select Case Index
            Case 0: dtcChequera(1) = dtcChequera(0).BoundText
            Case 1: dtcChequera(0) = dtcChequera(1).BoundText
        End Select
        With mskChequera
            .SelStart = 0
            .SelLength = 20
            .SetFocus
        End With
    End If
    '
    End Sub

    '---------------------------------------------------------------------------------------------
    Private Sub Form_Load() '
    '---------------------------------------------------------------------------------------------
    '
    datAhora = Now()
    'CONECTA CON EL ORIGEN DE LOS DATOS
    ChqrsCnn.CursorLocation = adUseClient
    ChqrsCnn.Open cnnOLEDB + mcDatos
    For i = 0 To 2
        adochequera(i).ConnectionString = cnnOLEDB + mcDatos
        adochequera(i).RecordSource = adochequera(i).Tag
        adochequera(i).Refresh
    Next
    Call RtnEstado(6, Toolbar1)
    If adochequera(1).Recordset.RecordCount < 1 Then
        MsgBox "Este Inmueble no tiene cuentas registradas......" & vbCrLf & _
        "Registrelas primero e intentelo nuevamente", vbInformation, App.ProductName
        Exit Sub
    End If
        Call centra_titulo(msgChequera, True)
    '---------------------------------------------------------------------------------------------
    '
    Call rtnRepint
    '
    End Sub

    
    Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    ChqrsCnn.Close
    Set ChqrsCnn = Nothing
    End Sub


    Private Sub msgChequera_EnterCell()
    If msgChequera.Row = 0 Or msgChequera.TextMatrix(msgChequera.RowSel, 0) = "" Then Exit Sub
    With adochequera(0).Recordset
            If .EOF Then Exit Sub
            .MoveFirst
            .Find "IDChequera ='" & msgChequera.TextMatrix(msgChequera.RowSel, 0) & "'"
            If Not .EOF Then
                dtcChequera(0) = msgChequera.TextMatrix(msgChequera.RowSel, 1)
                dtcChequera(1) = msgChequera.TextMatrix(msgChequera.RowSel, 2)
            End If
        End With
    End Sub

    '---------------------------------------------------------------------------------------------
    Private Sub mskChequera_KeyPress(KeyAscii As Integer)   '
    '---------------------------------------------------------------------------------------------
    '
    Call Validacion(KeyAscii, "1234567890")
    If KeyAscii = 13 Then txtchequera(1).SetFocus
    '
    End Sub
    

Private Sub optOrden_Click(Index As Integer): Call rtnRepint(optOrden(Index).Tag)
End Sub

    '---------------------------------------------------------------------------------------------
    Private Sub Toolbar1_ButtonClick(ByVal Button As ComctlLib.Button)  '
    '---------------------------------------------------------------------------------------------
    '
    With adochequera(0).Recordset
    '
        Select Case UCase(Button.Key)
    '
            Case "FIRST  'Primer Registro"
    '       ---------------------------
            .MoveFirst
            Call rtnDtc(!IDCuenta)
            cmbchequera = ftnNCheque
            
            Case "NEXT"  'Registro Previo
    '   ---------------------------
            .MoveNext
            If .EOF Then .MoveFirst
            Call rtnDtc(!IDCuenta)
            cmbchequera = ftnNCheque
            
            Case "PREVIOUS"  'Siguiente Registro
    '   --------------------------
            .MovePrevious
            If .BOF Then .MoveLast
            Call rtnDtc(!IDCuenta)
            cmbchequera = ftnNCheque
            
            Case "END"  'Último Registro
    '       ---------------------------
            .MoveLast
            Call rtnDtc(!IDCuenta)
            cmbchequera = ftnNCheque
            
            Case "NEW"  'Nuevo Registro
    '       ---------------------------

                .AddNew
                For i = 0 To 1
                    dtcChequera(i) = ""
                Next
                fraChequera(0).Enabled = True
                mskChequera = Date
                txtchequera(0) = Format(FntMax, "000")
                Call RtnEstado(5, Toolbar1)
                
            Case "SAVE"  'Actualizar Registro
    '       --------------------------
                mskChequera.PromptInclude = True
                !IDCuenta = ftnIDCta(dtcChequera(0))
                !Ultimo = IIf(.EditMode = adEditAdd, 0, !Ultimo)
                !fecha = mskChequera
                !Hora = Time
                !Usuario = gcUsuario
                .Update
                MsgBox "Registro actualizado...", vbInformation, App.ProductName
                Call rtnRepint
                mskChequera.PromptInclude = False
                fraChequera(0).Enabled = False
                Call RtnEstado(6, Toolbar1)
            
            Case "FIND"  'Buscar Registro
    '       --------------------------
        
            Case "UNDO"  'Cancelar Registro
    '       --------------------------
                If .EditMode = adEditAdd Then
                    mskChequera.PromptInclude = True
                    .CancelUpdate
                    mskChequera.PromptInclude = False
                    For i = 0 To 1
                        dtcChequera(i) = ""
                    Next
                    fraChequera(0).Enabled = False
                    With adochequera(0).Recordset
                        If .RecordCount > 0 Then Call rtnDtc(!IDCuenta): cmbchequera = ftnNCheque
                    End With
                Else
                    fraChequera(0).Enabled = False
                End If
                Call RtnEstado(6, Toolbar1)
            
            Case "DELETE"  'Eliminar Registro
    '   --------------------------
            
            If Respuesta(LoadResString(526)) Then
                .Delete
                .MoveNext
                If .EOF Then
                    .MovePrevious
                End If
                Call rtnRepint
                MsgBox "Registro Eliminado", vbInformation, App.ProductName
            End If
    
            Case "EDIT" 'Editar
    '       ---------------------------
                fraChequera(0).Enabled = True
                Call RtnEstado(5, Toolbar1)
                    
            Case "PRINT" 'Imprimir
    '        --------------------------
                Dim rpReporte As ctlReport
                Set rpReporte = New ctlReport
                With rpReporte
                    .Reporte = gcReport + "chequera_reg.rpt"
                    .OrigenDatos(0) = mcDatos
                    .Formulas(0) = "Inmueble='" & gcNomInm & "'"
                    .TituloVentana = "Reporte Chequeras Inm.:" & gcCodInm
                    If optOrden(0) Then
                        .OrdenRegistros(0) = "+Chequera.IDchequera"
                    ElseIf optOrden(1) Then
                        .OrdenRegistros(0) = "+Cuentas.NumCuenta"
                        .OrdenRegistros(1) = "+Chequera.DESDE"
                    ElseIf optOrden(3) Then
                        .OrdenRegistros(0) = "+Bancos.NombreBanco"
                        .OrdenRegistros(1) = "+Chequera.DESDE"
                    Else
                        .OrdenRegistros(0) = "+Chequera.Fecha"
                    End If
'
                    .Imprimir
                End With
                Set rpReporte = Nothing
            Case "CLOSE"    'Cerrar
    '       --------------------------
                Unload Me
    End Select
    '
    End With
    '
    End Sub
    
    '---------------------------------------------------------------------------------------------
    Private Function FntMax()   '
    '---------------------------------------------------------------------------------------------
    '
    Dim rstMax As ADODB.Recordset
    Set rstMax = New ADODB.Recordset
    '
    rstMax.Open "SELECT max(idchequera) FROM Chequera;", ChqrsCnn, adOpenKeyset, adLockOptimistic
    FntMax = IIf(IsNull(rstMax.Fields(0)), 0, rstMax.Fields(0)) + 1
    rstMax.Close: Set rstMax = Nothing
    '
    End Function

    
    Private Sub txtchequera_KeyPress(Index As Integer, KeyAscii As Integer)
    '
    Call Validacion(KeyAscii, "1234567890")
    If KeyAscii = 13 Then
        Select Case Index
            Case 1  'dESDE
    '       -----------------
                If cmbchequera <> "" Then
                    txtchequera(2) = Format(CLng(txtchequera(1)) + CLng(cmbchequera - 1), "000000")
                Else
                    cmbchequera.SetFocus
                End If
                txtchequera(1) = Format(txtchequera(1), "000000")
        End Select
        '
    End If
    '
    End Sub
    
    '---------------------------------------------------------------------------------------------
    Private Sub rtnDtc(intVal As Integer)   '
    '---------------------------------------------------------------------------------------------
    '
   With adochequera(1).Recordset
        .MoveFirst
        .Find "IDCuenta =" & intVal
        dtcChequera(0) = !NumCuenta
        dtcChequera(1) = !NombreBanco
    End With
    '
    End Sub
    
    '---------------------------------------------------------------------------------------------
    Private Sub rtnRepint(Optional orden As String) 'Redistribuye el ADODB.Recordset en el Grid
    '---------------------------------------------------------------------------------------------
    With msgChequera
        adochequera(2).Recordset.Requery
        adochequera(2).Recordset.Sort = orden
        If Not adochequera(2).Recordset.EOF Then
            Call rtnDtc(adochequera(0).Recordset!IDCuenta)
            Call rtnLimpiar_Grid(msgChequera)
            .Rows = adochequera(2).Recordset.RecordCount + 1
            
            'For I = 0 To 1
            adochequera(2).Recordset.MoveFirst: J = 0
            .Refresh
            Do Until adochequera(2).Recordset.EOF
                J = J + 1
                .TextMatrix(J, 0) = adochequera(2).Recordset!IDchequera
                .TextMatrix(J, 1) = adochequera(2).Recordset!NumCuenta
                .TextMatrix(J, 2) = adochequera(2).Recordset!NombreBanco
                .TextMatrix(J, 3) = Format(adochequera(2).Recordset!Desde, "000000")
                .TextMatrix(J, 4) = Format(adochequera(2).Recordset!Hasta, "000000")
                .TextMatrix(J, 5) = Format(adochequera(2).Recordset!Ultimo, "000000")
                adochequera(2).Recordset.MoveNext
            Loop
            '
        End If
    '
    End With
    '
    End Sub
    
    '---------------------------------------------------------------------------------------------
    '   Funtion:    ftnNCheque
    '
    '   Devuelve la diferencia entre #cheque inicial y #cheque final
    '---------------------------------------------------------------------------------------------
    Private Function ftnNCheque()
    ftnNCheque = CLng(txtchequera(2) - CLng(txtchequera(1))) + 1
    '
    End Function
    
    '---------------------------------------------------------------------------------------------
    '   Function:   ftnIDCta
    '
    '   Devuelve el identificardor de determinado número de cuenta
    '---------------------------------------------------------------------------------------------
    Private Function ftnIDCta(strNumCuenta As String) As Long
    '
    Dim rstCta As New ADODB.Recordset
    rstCta.Open "Cuentas", ChqrsCnn, adOpenKeyset, adLockOptimistic, cmdtable
    If Not rstCta.EOF Then
        rstCta.MoveFirst
        rstCta.Find "NumCuenta='" & strNumCuenta & "'"
        ftnIDCta = rstCta!IDCuenta
    End If
    rstCta.Close: Set rstCta = Nothing
    '
    End Function
