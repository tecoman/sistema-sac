VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form FrmAgenda2 
   Caption         =   "Agenda Telefónica Junta de Condominio"
   ClientHeight    =   135
   ClientLeft      =   60
   ClientTop       =   420
   ClientWidth     =   5160
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   135
   ScaleWidth      =   5160
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      Height          =   6975
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   11895
      Begin VB.Frame fraAgenda1 
         Caption         =   "Imprimir Reporte     y     Salir al Menú"
         ClipControls    =   0   'False
         Height          =   1335
         Index           =   1
         Left            =   8565
         TabIndex        =   8
         Top             =   4920
         Width           =   3120
         Begin VB.CommandButton cmdAgenda1 
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
         Begin VB.CommandButton cmdAgenda1 
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
         Left            =   4215
         TabIndex        =   3
         Top             =   4920
         Width           =   4305
         Begin VB.CommandButton cmdAgenda1 
            Height          =   400
            Index           =   2
            Left            =   3315
            Style           =   1  'Graphical
            TabIndex        =   5
            ToolTipText     =   "Buscar"
            Top             =   700
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
            TabIndex        =   4
            Top             =   743
            Width           =   690
         End
         Begin VB.Label lblAgenda 
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
            TabIndex        =   6
            Top             =   315
            Width           =   1005
         End
      End
      Begin MSAdodcLib.Adodc AdoProp 
         Height          =   360
         Left            =   4635
         Top             =   6330
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
      Begin MSDataGridLib.DataGrid DataGrid1 
         Bindings        =   "FrmAgenda2.frx":0000
         Height          =   4575
         Left            =   225
         TabIndex        =   2
         Top             =   255
         Width           =   11445
         _ExtentX        =   20188
         _ExtentY        =   8070
         _Version        =   393216
         AllowUpdate     =   0   'False
         BackColor       =   16777215
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
         Caption         =   "Agenda Telefónica Junta de Condominio"
         ColumnCount     =   10
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
            DataField       =   "Cedula"
            Caption         =   "Cedula"
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
            DataField       =   "CarJunta"
            Caption         =   "Cargo"
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
            DataField       =   "Telefonos"
            Caption         =   "Teléfono"
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
         BeginProperty Column06 
            DataField       =   "TelfHab"
            Caption         =   "Telf. Hab."
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
            DataField       =   "Celular"
            Caption         =   "Nº Celular"
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
         BeginProperty Column08 
            DataField       =   "Fax"
            Caption         =   "Nº Fax"
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
         BeginProperty Column09 
            DataField       =   "Email"
            Caption         =   "Email"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   """<"" $"
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
               ColumnWidth     =   840,189
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   3000,189
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   1739,906
            EndProperty
            BeginProperty Column03 
               ColumnWidth     =   1739,906
            EndProperty
            BeginProperty Column04 
               ColumnWidth     =   1739,906
            EndProperty
            BeginProperty Column05 
               ColumnWidth     =   1739,906
            EndProperty
            BeginProperty Column06 
               ColumnWidth     =   1739,906
            EndProperty
            BeginProperty Column07 
               ColumnWidth     =   1739,906
            EndProperty
            BeginProperty Column08 
               ColumnWidth     =   1739,906
            EndProperty
            BeginProperty Column09 
               ColumnWidth     =   2594,835
            EndProperty
         EndProperty
      End
   End
End
Attribute VB_Name = "FrmAgenda2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

    Private Sub cmdAgenda1_Click(Index As Integer)
    '
    Select Case Index
    '
        Case 0 'Boton imprimir
        '--------------------
            mcTitulo = "Agenda Junta Condominio"
            mcReport = "agendajuntacon.rpt"
            mcOrdCod = ""
            mcOrdAlfa = ""
            mcCrit = ""
            FrmReport.Show
            '
        Case 1  'Boton Salir
        '--------------------
            Unload Me: Set FrmAgenda2 = Nothing
        Case 2  'Boton Buscar
        '--------------------
            Call Buscar_Propietario
    End Select
    '
    End Sub

    Private Sub Form_Load()
    AdoProp.ConnectionString = cnnOLEDB + mcDatos
    AdoProp.CommandType = adCmdText
    AdoProp.RecordSource = "Select * FROM Propietarios WHERE CarJunta<>'' AND Codigo NOT Like '" _
    & "U%'  ORDER BY codigo"
    AdoProp.Refresh
    
    cmdAgenda1(0).Picture = LoadResPicture("Print", vbResIcon)
    cmdAgenda1(1).Picture = LoadResPicture("salir", vbResIcon)
    cmdAgenda1(2).Picture = LoadResPicture("Buscar", vbResIcon)
    AdoProp.ConnectionString = cnnOLEDB & mcDatos
    AdoProp.Refresh
    txtAgenda1(1) = AdoProp.Recordset.RecordCount
    End Sub

    '---------------------------------------------------------------------------------------------
    '   Rutina: Buscar_Propietario
    '
    '   Busqueda de Propietario por parte del nombre, busca todas las coincidencias
    '   dentro del ADODB.Recordset.
    '---------------------------------------------------------------------------------------------
    Private Sub Buscar_Propietario()
    '
    If txtAgenda1(0) = "" Then Exit Sub
    With AdoProp.Recordset
    
        If Not .EOF And Not .BOF Then
            .MoveNext
            If .EOF Then .MoveFirst
            .Find "Nombre LIKE '%" & txtAgenda1(0) & "%'"
            
            If .EOF Then
            
                .Find "Nombre LIKE '%" & txtAgenda1(0) & "%'"
                
                If .EOF Then
                
                    MsgBox "Propietario No Registrado...", vbInformation, App.ProductName
                    .MoveFirst
                    
                End If
            '
            End If
            
        End If
        '
    End With
    '
    End Sub


    Private Sub Form_Resize()
    'configura la presentación de los controles en pantalla
    '
    Frame1.Top = Me.ScaleTop
    Frame1.Left = Me.ScaleLeft
    Frame1.Height = Me.ScaleHeight
    Frame1.Width = Me.ScaleWidth
    DataGrid1.Width = Frame1.Width - (DataGrid1.Left * 2)
    DataGrid1.Height = Frame1.Height - DataGrid1.Top - 200 - fraAgenda1(0).Height
    fraAgenda1(0).Left = (DataGrid1.Width + DataGrid1.Left) - (fraAgenda1(0).Width + 100 + _
    fraAgenda1(1).Width)
    fraAgenda1(0).Top = DataGrid1.Top + DataGrid1.Height + 100
    fraAgenda1(1).Left = fraAgenda1(0).Left + fraAgenda1(0).Width + 100
    fraAgenda1(1).Top = DataGrid1.Top + DataGrid1.Height + 100
    '
    End Sub

    Private Sub txtAgenda1_KeyPress(Index%, KeyAscii%)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then Call Buscar_Propietario
    End Sub
