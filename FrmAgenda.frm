VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form FrmAgenda 
   Caption         =   "Agenda Telefónica"
   ClientHeight    =   30
   ClientLeft      =   60
   ClientTop       =   420
   ClientWidth     =   2580
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   30
   ScaleWidth      =   2580
   Tag             =   "1"
   WindowState     =   2  'Maximized
   Begin VB.Frame fraPro 
      Height          =   6975
      Index           =   0
      Left            =   30
      TabIndex        =   1
      Top             =   30
      Width           =   11865
      Begin MSAdodcLib.Adodc AdoProp 
         Height          =   360
         Left            =   360
         Top             =   5610
         Visible         =   0   'False
         Width           =   3180
         _ExtentX        =   5609
         _ExtentY        =   635
         ConnectMode     =   0
         CursorLocation  =   2
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
         Caption         =   "AdoProp"
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
      Begin VB.Frame fraPro 
         Caption         =   "Imprimir Reporte     y     Salir al Menú"
         ClipControls    =   0   'False
         Height          =   1335
         Index           =   1
         Left            =   8370
         TabIndex        =   8
         Top             =   5490
         Width           =   3120
         Begin VB.CommandButton cmdPro 
            Cancel          =   -1  'True
            Caption         =   "Imprimir"
            Height          =   765
            Index           =   0
            Left            =   390
            Style           =   1  'Graphical
            TabIndex        =   10
            Top             =   360
            Width           =   1065
         End
         Begin VB.CommandButton cmdPro 
            Caption         =   "Salir"
            Height          =   765
            Index           =   1
            Left            =   1695
            Style           =   1  'Graphical
            TabIndex        =   9
            Top             =   360
            Width           =   1065
         End
      End
      Begin VB.Frame fraPro 
         ClipControls    =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1335
         Index           =   2
         Left            =   4020
         TabIndex        =   3
         Top             =   5490
         Width           =   4305
         Begin VB.CommandButton cmdPro 
            Height          =   400
            Index           =   2
            Left            =   3300
            Style           =   1  'Graphical
            TabIndex        =   5
            ToolTipText     =   "Buscar"
            Top             =   700
            Width           =   690
         End
         Begin VB.TextBox txtPro 
            BackColor       =   &H00FFFFFF&
            Height          =   315
            Index           =   0
            Left            =   1365
            TabIndex        =   0
            Top             =   270
            Width           =   2640
         End
         Begin VB.TextBox txtPro 
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
            Index           =   1
            Left            =   2205
            Locked          =   -1  'True
            TabIndex        =   4
            Top             =   743
            Width           =   690
         End
         Begin VB.Label lblPro 
            Caption         =   "Total Propietarios:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   1
            Left            =   285
            TabIndex        =   7
            Top             =   780
            Width           =   1875
         End
         Begin VB.Label lblPro 
            Caption         =   "Buscar a:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   0
            Left            =   285
            TabIndex        =   6
            Top             =   315
            Width           =   1005
         End
      End
      Begin MSDataGridLib.DataGrid DataGrid1 
         Bindings        =   "FrmAgenda.frx":0000
         Height          =   5055
         Left            =   255
         TabIndex        =   2
         Top             =   345
         Width           =   11250
         _ExtentX        =   19844
         _ExtentY        =   8916
         _Version        =   393216
         AllowUpdate     =   0   'False
         BackColor       =   16777215
         HeadLines       =   1,5
         RowHeight       =   15
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
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Agenda Telefónica Propietarios"
         ColumnCount     =   9
         BeginProperty Column00 
            DataField       =   "Codigo"
            Caption         =   "Codigo"
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
            Caption         =   "Nombre"
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
            DataField       =   "Telefonos"
            Caption         =   "Telefonos"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   "(0000)-###-##-##"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   8202
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column03 
            DataField       =   "Celular"
            Caption         =   "Celular"
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
            DataField       =   "ExtOfc"
            Caption         =   "Ext.Ofc."
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
            DataField       =   "TelfHab"
            Caption         =   "TelfHab"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   "(0000)-###-##-##"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   8202
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column06 
            DataField       =   "Fax"
            Caption         =   "Fax"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   "(0000)-###-##-##"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   8202
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column07 
            DataField       =   "Email"
            Caption         =   "Email"
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
         BeginProperty Column08 
            DataField       =   ""
            Caption         =   ""
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
            AllowRowSizing  =   0   'False
            AllowSizing     =   0   'False
            BeginProperty Column00 
               ColumnWidth     =   659,906
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   3734,929
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   1785,26
            EndProperty
            BeginProperty Column03 
               ColumnWidth     =   1665,071
            EndProperty
            BeginProperty Column04 
               ColumnWidth     =   810,142
            EndProperty
            BeginProperty Column05 
            EndProperty
            BeginProperty Column06 
               ColumnWidth     =   1920,189
            EndProperty
            BeginProperty Column07 
               ColumnWidth     =   2970,142
            EndProperty
            BeginProperty Column08 
               ColumnWidth     =   14,74
            EndProperty
         EndProperty
      End
   End
End
Attribute VB_Name = "FrmAgenda"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

    Private Sub cmdPro_Click(Index As Integer)  'Matríz de Contoles Button
    '
    Select Case Index
        '
        Case 0  'Botón Imprimir
        '--------------------
            mcTitulo = "Agenda de Propietarios"
            mcReport = "AgendaProp.Rpt"
            mcOrdCod = ""
            mcOrdAlfa = ""
            mcCrit = ""
            FrmReport.Show
            
        Case 1  'Botón Salir
        '--------------------
            Unload Me
            Set FrmAgenda = Nothing
            '
        Case 2  'Botón Buscar
        '--------------------
            Call Buscar_Propietario
            '
    End Select
    '
    End Sub

    Private Sub Form_Load()
    'carga el formulario de agenda telf. propietarios
    '
    cmdPro(0).Picture = LoadResPicture("Print", vbResIcon)
    cmdPro(1).Picture = LoadResPicture("Salir", vbResIcon)
    cmdPro(2).Picture = LoadResPicture("Buscar", vbResIcon)
    AdoProp.ConnectionString = cnnOLEDB & mcDatos
    AdoProp.CursorLocation = adUseClient
    AdoProp.CommandType = adCmdTable
    AdoProp.RecordSource = "Propietarios"
    AdoProp.Refresh
    AdoProp.Recordset.Filter = "Codigo <> 'U" & gcCodInm & "'"
    AdoProp.Recordset.Sort = "Codigo"
    '
    End Sub



    Private Sub Form_Resize()
    'configura la presentación de los controles en pantalla
    fraPro(0).Top = Me.ScaleTop
    fraPro(0).Left = Me.ScaleLeft
    fraPro(0).Height = Me.ScaleHeight
    fraPro(0).Width = Me.ScaleWidth
    '
    DataGrid1.Width = fraPro(0).Width - (DataGrid1.Left * 2)
    DataGrid1.Height = fraPro(0).Height - DataGrid1.Top - 200 - fraPro(2).Height
    '
    fraPro(2).Left = (DataGrid1.Width + DataGrid1.Left) - (fraPro(2).Width + 100 + _
    fraPro(1).Width)
    fraPro(2).Top = DataGrid1.Top + DataGrid1.Height + 100
    fraPro(1).Left = Me.fraPro(2).Left + Me.fraPro(2).Width + 100
    fraPro(1).Top = DataGrid1.Top + DataGrid1.Height + 100
    '
    End Sub
    

    Private Sub txtPro_KeyPress(Index As Integer, KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then Call Buscar_Propietario
    End Sub

    '---------------------------------------------------------------------------------------------
    '   Rutina: Buscar_Propietario
    '
    '   Busqueda de Propietario por parte del nombre, busca todas las coincidencias
    '   dentro del ADODB.Recordset.
    '---------------------------------------------------------------------------------------------
    Private Sub Buscar_Propietario()
    '
    With AdoProp.Recordset
        .MoveNext
        
        If .EOF Then .MoveFirst
        .Find "Nombre LIKE '%" & txtPro(0) & "%'"
        If .EOF Then
            .Find "Nombre LIKE '%" & txtPro(0) & "%'"
            If .EOF Then
                MsgBox "Propietario No Registrado...", vbInformation, App.ProductName
                .MoveFirst
            End If
            '
        End If
        '
    End With
    '
    End Sub
