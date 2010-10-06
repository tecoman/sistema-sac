VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmNEmp 
   Caption         =   "Nº Empresa en S.S.O"
   ClientHeight    =   6225
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6285
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   6225
   ScaleWidth      =   6285
   StartUpPosition =   3  'Windows Default
   Begin MSDataGridLib.DataGrid Dtg 
      Bindings        =   "frmNEmp.frx":0000
      Height          =   4125
      Left            =   255
      TabIndex        =   2
      Top             =   1305
      Width           =   5730
      _ExtentX        =   10107
      _ExtentY        =   7276
      _Version        =   393216
      AllowUpdate     =   -1  'True
      BorderStyle     =   0
      HeadLines       =   2
      RowHeight       =   19
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
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   3
      BeginProperty Column00 
         DataField       =   "CodInm"
         Caption         =   "Código"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   " "
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   8202
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   "Nombre"
         Caption         =   "Inmueble"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   " $"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   8202
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column02 
         DataField       =   "Campo3"
         Caption         =   "Nº Empresa"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   " $"
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
            DividerStyle    =   6
            Locked          =   -1  'True
            ColumnWidth     =   689,953
         EndProperty
         BeginProperty Column01 
            Locked          =   -1  'True
            ColumnWidth     =   3000,189
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   1319,811
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton Command1 
      Cancel          =   -1  'True
      Caption         =   "&Cerrar"
      Default         =   -1  'True
      Height          =   450
      Left            =   4755
      TabIndex        =   3
      Top             =   5610
      Width           =   1215
   End
   Begin MSAdodcLib.Adodc Ado 
      Height          =   330
      Left            =   285
      Top             =   5715
      Visible         =   0   'False
      Width           =   2640
      _ExtentX        =   4657
      _ExtentY        =   582
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
      Caption         =   "NEmpresa"
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
      Caption         =   "Solo puede editar el Nº de empresa."
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
      Index           =   1
      Left            =   360
      TabIndex        =   1
      Top             =   795
      Width           =   3870
   End
   Begin VB.Label Lbl 
      Caption         =   "Para editar esta información escriba directamente sobre cada celda. "
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
      Index           =   0
      Left            =   360
      TabIndex        =   0
      Top             =   315
      Width           =   5565
   End
End
Attribute VB_Name = "frmNEmp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
Unload Me
Set frmNEmp = Nothing
End Sub

Private Sub Dtg_HeadClick(ByVal ColIndex As Integer)
'-------------------------------------------------------------------------------------------------
'ordena la lista por el encabezado de la columna
'   ASC 0
'   DES 1
'-------------------------------------------------------------------------------------------------
Static orden(1) As Integer
'
If ColIndex < 2 Then
    '
    Ado.Recordset.Sort = Dtg.Columns(ColIndex).DataField & " " & IIf(orden(ColIndex) = 0, _
    "DESC", "ASC")
    orden(ColIndex) = IIf(orden(ColIndex) = 0, 1, 0)
    '
End If
'
End Sub

Private Sub Dtg_KeyPress(KeyAscii As Integer)
'
KeyAscii = Asc(UCase(Chr(KeyAscii)))
Call Validacion(KeyAscii, "D0123456789")
'
End Sub

Private Sub Form_Load()
Call Configurar_ADO(Ado, adCmdText, "SELECT * FROM Inmueble WHERE Inactivo=False", gcPath _
& "\sac.mdb", "CodInm")
'With Ado
'
'    .CursorLocation = adUseClient
'    .CommandType = adCmdText
'    .LockType = adLockOptimistic
'    .ConnectionString = cnnConexion.ConnectionString
'    .CursorType = adOpenKeyset
'    .RecordSource = "SELECT * FROM Inmueble WHERE Inactivos = False"
'    .Refresh
'End With
Call CenterForm(frmNEmp)
Me.Show vbModeless, FrmAdmin
End Sub
