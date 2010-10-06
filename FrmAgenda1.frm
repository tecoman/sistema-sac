VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form FrmAgenda1 
   Caption         =   "Agenda Telefónica"
   ClientHeight    =   90
   ClientLeft      =   165
   ClientTop       =   285
   ClientWidth     =   2415
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   90
   ScaleWidth      =   2415
   Tag             =   "1"
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      Height          =   7245
      Left            =   -15
      TabIndex        =   1
      Top             =   -75
      Width           =   12255
      Begin MSDataGridLib.DataGrid DataGrid1 
         Bindings        =   "FrmAgenda1.frx":0000
         Height          =   4755
         Left            =   315
         TabIndex        =   3
         Top             =   585
         Width           =   11340
         _ExtentX        =   20003
         _ExtentY        =   8387
         _Version        =   393216
         AllowUpdate     =   0   'False
         BorderStyle     =   0
         HeadLines       =   1,5
         RowHeight       =   15
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
         Caption         =   "Agenda Telefónica Proveedores"
         ColumnCount     =   6
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
            DataField       =   "NombProv"
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
            DataField       =   "Ramo"
            Caption         =   "Ramo"
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
         BeginProperty Column04 
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
         BeginProperty Column05 
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
         SplitCount      =   1
         BeginProperty Split0 
            BeginProperty Column00 
               ColumnWidth     =   1005,165
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   3720,189
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   3720,189
            EndProperty
            BeginProperty Column03 
               ColumnWidth     =   3435,024
            EndProperty
            BeginProperty Column04 
               ColumnWidth     =   1739,906
            EndProperty
            BeginProperty Column05 
               ColumnWidth     =   2940,095
            EndProperty
         EndProperty
      End
      Begin VB.Frame fraAgenda1 
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
         Index           =   0
         Left            =   4185
         TabIndex        =   6
         Top             =   5550
         Width           =   4305
         Begin VB.TextBox txtAgenda1 
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
            TabIndex        =   8
            Top             =   743
            Width           =   690
         End
         Begin VB.TextBox txtAgenda1 
            BackColor       =   &H00FFFFFF&
            Height          =   315
            Index           =   0
            Left            =   1365
            TabIndex        =   0
            Top             =   270
            Width           =   2640
         End
         Begin VB.CommandButton cmdAgenda1 
            Height          =   400
            Index           =   2
            Left            =   3300
            Style           =   1  'Graphical
            TabIndex        =   7
            ToolTipText     =   "Buscar"
            Top             =   700
            Width           =   690
         End
         Begin VB.Label lblAgenda 
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
            TabIndex        =   10
            Top             =   315
            Width           =   1005
         End
         Begin VB.Label lblAgenda 
            Caption         =   "Total Proveedores:"
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
            TabIndex        =   9
            Top             =   780
            Width           =   1875
         End
      End
      Begin VB.Frame fraAgenda1 
         Caption         =   "Imprimir Reporte     y     Salir al Menú"
         ClipControls    =   0   'False
         Height          =   1335
         Index           =   1
         Left            =   8535
         TabIndex        =   2
         Top             =   5550
         Width           =   3120
         Begin VB.CommandButton cmdAgenda1 
            Caption         =   "Salir"
            Height          =   765
            Index           =   1
            Left            =   1695
            MaskColor       =   &H80000004&
            Style           =   1  'Graphical
            TabIndex        =   5
            Top             =   360
            Width           =   1065
         End
         Begin VB.CommandButton cmdAgenda1 
            Cancel          =   -1  'True
            Caption         =   "Imprimir"
            Height          =   765
            Index           =   0
            Left            =   390
            Style           =   1  'Graphical
            TabIndex        =   4
            Top             =   360
            Width           =   1065
         End
      End
      Begin MSAdodcLib.Adodc AdoProv 
         Height          =   360
         Left            =   4605
         Top             =   6810
         Visible         =   0   'False
         Width           =   2130
         _ExtentX        =   3757
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
         Caption         =   "AdoProv"
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
Attribute VB_Name = "FrmAgenda1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
    

    Private Sub cmdAgenda1_Click(Index As Integer)
    '
    Select Case Index
        '
        Case 0  'Printer
        '--------------------
            mcTitulo = "Agenda de Proveedores"
            mcReport = "AgendaProv.Rpt"
            mcOrdCod = ""
            mcOrdAlfa = ""
            mcCrit = ""
            FrmReport.Show
        Case 1  'Salir
        '--------------------
            Unload Me
            Set FrmAgenda1 = Nothing
        Case 2  'Buscar
        '--------------------
            With AdoProv.Recordset
                .MoveNext
                If .EOF Then .MoveFirst
                .Find "NombProv LIKE '%" & txtAgenda1(0) & "%'"
                If .EOF Then
                    .Find "NombProv LIKE '%" & txtAgenda1(0) & "%'"
                    If .EOF Then
                        MsgBox "Proveedor No Registrado...", vbInformation, App.ProductName
                        .MoveFirst
                    End If
                    '
                End If
                '
            End With
            '
    End Select
    '
    End Sub


    Private Sub Form_Load()
    '
    AdoProv.ConnectionString = cnnOLEDB + gcPath & "\sac.mdb"
    AdoProv.CommandType = adCmdTable
    AdoProv.CursorLocation = adUseClient
    AdoProv.RecordSource = "Proveedores"
    AdoProv.Refresh
    AdoProv.Recordset.Sort = "Codigo"
    cmdAgenda1(0).Picture = LoadResPicture("Print", vbResIcon)
    cmdAgenda1(1).Picture = LoadResPicture("salir", vbResIcon)
    cmdAgenda1(2).Picture = LoadResPicture("Buscar", vbResIcon)
    txtAgenda1(1) = AdoProv.Recordset.RecordCount
    '
    End Sub

    Private Sub Form_Resize()
    'configura la presentación de los controles en pantalla
    Frame1.Top = Me.ScaleTop
    Frame1.Left = Me.ScaleLeft
    Frame1.Height = Me.ScaleHeight
    Frame1.Width = Me.ScaleWidth
    
    DataGrid1.Width = Frame1.Width - (DataGrid1.Left * 2)
    DataGrid1.Height = Frame1.Height - DataGrid1.Top - 200 - fraAgenda1(0).Height
    '
    fraAgenda1(0).Left = (DataGrid1.Width + DataGrid1.Left) - (fraAgenda1(0).Width + 100 + _
    fraAgenda1(1).Width)
    fraAgenda1(0).Top = DataGrid1.Top + DataGrid1.Height + 100
    fraAgenda1(1).Left = fraAgenda1(0).Left + fraAgenda1(0).Width + 100
    fraAgenda1(1).Top = DataGrid1.Top + DataGrid1.Height + 100
    '
    End Sub

    Private Sub txtAgenda1_KeyPress(Index%, KeyAscii%)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    End Sub
