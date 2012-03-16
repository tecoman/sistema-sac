VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{65E121D4-0C60-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCHRT20.OCX"
Begin VB.Form frmIEco 
   Caption         =   "Informe Económico"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   Tag             =   "1"
   WindowState     =   2  'Maximized
   Begin VB.TextBox txtCopy 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   5
      Left            =   5610
      TabIndex        =   51
      Top             =   8445
      Width           =   1080
   End
   Begin VB.TextBox txtCopy 
      Height          =   315
      Index           =   4
      Left            =   510
      TabIndex        =   50
      Top             =   8490
      Width           =   4725
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   8895
      Left            =   90
      TabIndex        =   0
      Top             =   150
      Width           =   12825
      _ExtentX        =   22622
      _ExtentY        =   15690
      _Version        =   393216
      TabHeight       =   520
      ShowFocusRect   =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Selección"
      TabPicture(0)   =   "frmIEco.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "fraIECO(0)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "txtCopy(2)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "txtCopy(3)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "Comparativo Fondo - Deuda"
      TabPicture(1)   =   "frmIEco.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "SSTab2"
      Tab(1).Control(1)=   "grid(3)"
      Tab(1).Control(2)=   "cmdIECO(5)"
      Tab(1).Control(3)=   "opt(0)"
      Tab(1).Control(4)=   "opt(1)"
      Tab(1).ControlCount=   5
      TabCaption(2)   =   "Informe Económico"
      TabPicture(2)   =   "frmIEco.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "grid(4)"
      Tab(2).Control(1)=   "cmdIECO(3)"
      Tab(2).Control(2)=   "cmdIECO(4)"
      Tab(2).Control(3)=   "fraIECO(2)"
      Tab(2).ControlCount=   4
      Begin VB.TextBox txtCopy 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   3
         Left            =   5520
         TabIndex        =   49
         Top             =   7815
         Width           =   1080
      End
      Begin VB.TextBox txtCopy 
         Height          =   315
         Index           =   2
         Left            =   435
         TabIndex        =   48
         Top             =   7815
         Width           =   4725
      End
      Begin VB.OptionButton opt 
         Alignment       =   1  'Right Justify
         Caption         =   "SAC"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Index           =   1
         Left            =   -66240
         TabIndex        =   37
         Top             =   7140
         Width           =   1785
      End
      Begin VB.OptionButton opt 
         Alignment       =   1  'Right Justify
         Caption         =   "Inversiones"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Index           =   0
         Left            =   -66240
         TabIndex        =   36
         Top             =   6795
         Value           =   -1  'True
         Width           =   1785
      End
      Begin VB.CommandButton cmdIECO 
         Caption         =   "Imprimir"
         Height          =   795
         Index           =   5
         Left            =   -64200
         Picture         =   "frmIEco.frx":0054
         Style           =   1  'Graphical
         TabIndex        =   35
         Top             =   6690
         Width           =   1395
      End
      Begin VB.Frame fraIECO 
         Caption         =   "Aplicar Filtro:"
         Height          =   780
         Index           =   2
         Left            =   -72690
         TabIndex        =   28
         Top             =   6810
         Width           =   5880
         Begin VB.OptionButton optIECO 
            Caption         =   "Gastos variables"
            Height          =   420
            Index           =   2
            Left            =   3915
            TabIndex        =   31
            Tag             =   "AND Tgastos.Fijo = False"
            Top             =   210
            Width           =   1680
         End
         Begin VB.OptionButton optIECO 
            Caption         =   "Gastos Fijos"
            Height          =   420
            Index           =   1
            Left            =   2400
            TabIndex        =   30
            Tag             =   "AND Tgastos.Fijo=True"
            Top             =   210
            Width           =   1320
         End
         Begin VB.OptionButton optIECO 
            Caption         =   "Todos los Gastos"
            Height          =   420
            Index           =   0
            Left            =   420
            TabIndex        =   29
            Top             =   210
            Value           =   -1  'True
            Width           =   2010
         End
      End
      Begin VB.CommandButton cmdIECO 
         Caption         =   "Exportar Excel"
         Height          =   705
         Index           =   4
         Left            =   -66510
         TabIndex        =   24
         Top             =   6870
         Width           =   1410
      End
      Begin VB.CommandButton cmdIECO 
         Caption         =   "Imprimir"
         Height          =   705
         Index           =   3
         Left            =   -64980
         TabIndex        =   21
         Top             =   6870
         Width           =   1410
      End
      Begin VB.Frame fraIECO 
         Height          =   7215
         Index           =   0
         Left            =   165
         TabIndex        =   1
         Top             =   360
         Width           =   11295
         Begin VB.Frame fraIECO 
            Caption         =   "Nº Copias"
            Height          =   1035
            Index           =   3
            Left            =   7455
            TabIndex        =   41
            Top             =   4815
            Width           =   3690
            Begin VB.TextBox txtCopy 
               Alignment       =   1  'Right Justify
               Height          =   315
               Index           =   1
               Left            =   2265
               TabIndex        =   45
               Top             =   600
               Width           =   1080
            End
            Begin VB.TextBox txtCopy 
               Alignment       =   1  'Right Justify
               Height          =   315
               Index           =   0
               Left            =   2265
               TabIndex        =   44
               Top             =   180
               Width           =   1080
            End
            Begin VB.Label lblIeco 
               AutoSize        =   -1  'True
               Caption         =   "Económico:"
               Height          =   195
               Index           =   12
               Left            =   960
               TabIndex        =   43
               Top             =   660
               Width           =   870
            End
            Begin VB.Label lblIeco 
               AutoSize        =   -1  'True
               Caption         =   "Fondo - Deuda:"
               Height          =   195
               Index           =   11
               Left            =   960
               TabIndex        =   42
               Top             =   240
               Width           =   1140
            End
         End
         Begin VB.Frame fraIECO 
            BorderStyle     =   0  'None
            Height          =   1110
            Index           =   1
            Left            =   120
            TabIndex        =   25
            Top             =   240
            Width           =   3300
            Begin VB.TextBox txtIEco 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   915
               TabIndex        =   26
               Top             =   585
               Width           =   1350
            End
            Begin VB.Label lblIeco 
               Caption         =   "1.- Señale el Nº de facturaciones a incluir en el reporte:"
               Height          =   420
               Index           =   3
               Left            =   45
               TabIndex        =   27
               Top             =   360
               Width           =   3150
            End
         End
         Begin VB.CheckBox chk 
            Caption         =   "Mostrar intereses moratotios"
            Height          =   495
            Left            =   7875
            TabIndex        =   23
            Top             =   1650
            Width           =   2535
         End
         Begin VB.CommandButton cmdIECO 
            Cancel          =   -1  'True
            Caption         =   "&Informe Económico"
            Height          =   705
            Index           =   0
            Left            =   9930
            TabIndex        =   16
            Top             =   6195
            Width           =   1260
         End
         Begin VB.CommandButton cmdIECO 
            Caption         =   "&Comparativo Fondo - Deuda"
            Height          =   705
            Index           =   1
            Left            =   8670
            TabIndex        =   3
            Top             =   6195
            Width           =   1260
         End
         Begin VB.CommandButton cmdIECO 
            Caption         =   "&Salir"
            Height          =   705
            Index           =   2
            Left            =   7410
            TabIndex        =   2
            Top             =   6195
            Width           =   1260
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid grid 
            Height          =   2085
            Index           =   0
            Left            =   210
            TabIndex        =   4
            Tag             =   "1000|1000|0"
            Top             =   1995
            Width           =   2955
            _ExtentX        =   5212
            _ExtentY        =   3678
            _Version        =   393216
            Cols            =   3
            FixedCols       =   0
            BackColorFixed  =   -2147483646
            ForeColorFixed  =   -2147483639
            BackColorSel    =   -2147483643
            ForeColorSel    =   -2147483646
            BackColorBkg    =   -2147483636
            FormatString    =   "Cod.Gasto|Selección|"
            _NumberOfBands  =   1
            _Band(0).Cols   =   3
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid grid 
            Height          =   2085
            Index           =   1
            Left            =   165
            TabIndex        =   5
            Tag             =   "1000|1000|0"
            Top             =   4935
            Width           =   2955
            _ExtentX        =   5212
            _ExtentY        =   3678
            _Version        =   393216
            Cols            =   3
            FixedCols       =   0
            BackColorFixed  =   -2147483646
            ForeColorFixed  =   -2147483639
            BackColorSel    =   -2147483643
            ForeColorSel    =   -2147483646
            BackColorBkg    =   -2147483636
            FormatString    =   "Cod.Gasto|Selección|"
            _NumberOfBands  =   1
            _Band(0).Cols   =   3
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid grid 
            Height          =   6030
            Index           =   2
            Left            =   3990
            TabIndex        =   6
            Tag             =   "1000|1000"
            Top             =   960
            Width           =   2955
            _ExtentX        =   5212
            _ExtentY        =   10636
            _Version        =   393216
            FixedCols       =   0
            BackColorFixed  =   -2147483646
            ForeColorFixed  =   -2147483639
            BackColorSel    =   -2147483643
            ForeColorSel    =   -2147483646
            BackColorBkg    =   -2147483636
            FormatString    =   "Cod.Gasto|Selección"
            _NumberOfBands  =   1
            _Band(0).Cols   =   2
         End
         Begin MSComCtl2.DTPicker dtp 
            Height          =   315
            Index           =   0
            Left            =   7875
            TabIndex        =   7
            Top             =   780
            Width           =   1395
            _ExtentX        =   2461
            _ExtentY        =   556
            _Version        =   393216
            Format          =   84803585
            CurrentDate     =   37909
         End
         Begin MSMask.MaskEdBox MskFecha 
            Bindings        =   "frmIEco.frx":035E
            DataSource      =   "DEFrmFactura"
            Height          =   315
            Index           =   0
            Left            =   9360
            TabIndex        =   17
            Top             =   3345
            Width           =   1410
            _ExtentX        =   2487
            _ExtentY        =   556
            _Version        =   393216
            MaxLength       =   7
            Format          =   "MM/yyyy"
            Mask            =   "##/####"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MskFecha 
            Bindings        =   "frmIEco.frx":0380
            DataSource      =   "DEFrmFactura"
            Height          =   315
            Index           =   1
            Left            =   9360
            TabIndex        =   18
            Top             =   3795
            Width           =   1410
            _ExtentX        =   2487
            _ExtentY        =   556
            _Version        =   393216
            MaxLength       =   7
            Format          =   "MM/yyyy"
            Mask            =   "##/####"
            PromptChar      =   "_"
         End
         Begin MSComCtl2.DTPicker dtp 
            Height          =   315
            Index           =   1
            Left            =   9405
            TabIndex        =   47
            Top             =   4260
            Width           =   1395
            _ExtentX        =   2461
            _ExtentY        =   556
            _Version        =   393216
            Format          =   84803585
            CurrentDate     =   37909
         End
         Begin VB.Label lblIeco 
            Caption         =   "8.- Fecha de la Asamblea:"
            Height          =   420
            Index           =   10
            Left            =   7485
            TabIndex        =   46
            Top             =   4305
            Width           =   2025
         End
         Begin VB.Label lblIeco 
            Caption         =   "6.- Desea que aparezcan los intereses moratorios"
            Height          =   420
            Index           =   9
            Left            =   7635
            TabIndex        =   22
            Top             =   1260
            Width           =   3375
         End
         Begin VB.Label lblIeco 
            Caption         =   "En cada página del repote el máximo de meses a presetar es de doce (12)"
            Height          =   420
            Index           =   8
            Left            =   7980
            TabIndex        =   15
            Top             =   2775
            Width           =   3090
         End
         Begin VB.Label lblIeco 
            AutoSize        =   -1  'True
            Caption         =   "Hasta:"
            Height          =   195
            Index           =   7
            Left            =   8640
            TabIndex        =   14
            Top             =   3855
            Width           =   465
         End
         Begin VB.Label lblIeco 
            AutoSize        =   -1  'True
            Caption         =   "Desde:"
            Height          =   195
            Index           =   6
            Left            =   8640
            TabIndex        =   13
            Top             =   3405
            Width           =   510
         End
         Begin VB.Label lblIeco 
            Caption         =   "7.- Selecciones un rango de fechas para imprimir el Informe Económico:"
            Height          =   420
            Index           =   2
            Left            =   7635
            TabIndex        =   12
            Top             =   2295
            Width           =   3495
         End
         Begin VB.Line ln 
            BorderColor     =   &H00FFFFFF&
            Index           =   3
            X1              =   7320
            X2              =   7320
            Y1              =   465
            Y2              =   7020
         End
         Begin VB.Line ln 
            BorderColor     =   &H00808080&
            BorderWidth     =   2
            Index           =   2
            X1              =   7320
            X2              =   7320
            Y1              =   465
            Y2              =   7020
         End
         Begin VB.Line ln 
            BorderColor     =   &H00FFFFFF&
            Index           =   1
            X1              =   3555
            X2              =   3555
            Y1              =   465
            Y2              =   7005
         End
         Begin VB.Line ln 
            BorderColor     =   &H00808080&
            BorderWidth     =   2
            Index           =   0
            X1              =   3555
            X2              =   3555
            Y1              =   450
            Y2              =   6990
         End
         Begin VB.Label lblIeco 
            Caption         =   "2.- Seleccione los códigos de los fondos que se aparecerán en el reporte:"
            Height          =   420
            Index           =   4
            Left            =   135
            TabIndex        =   11
            Top             =   1470
            Width           =   3150
         End
         Begin VB.Label lblIeco 
            Caption         =   "3.- Seleccione los códigos de las cuotas especiales que aparecerán en el reporte:"
            Height          =   480
            Index           =   5
            Left            =   135
            TabIndex        =   10
            Top             =   4425
            Width           =   3150
         End
         Begin VB.Label lblIeco 
            Caption         =   "4.- Seleccione los códigos de los gastos para el gráfico:"
            Height          =   420
            Index           =   0
            Left            =   3885
            TabIndex        =   9
            Top             =   495
            Width           =   3150
         End
         Begin VB.Label lblIeco 
            Caption         =   "5.- Indique la fecha de la última asamblea"
            Height          =   420
            Index           =   1
            Left            =   7635
            TabIndex        =   8
            Top             =   495
            Width           =   3375
         End
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid grid 
         Height          =   2265
         Index           =   3
         Left            =   -74580
         TabIndex        =   19
         Tag             =   "1000|1000|0"
         Top             =   495
         Width           =   12060
         _ExtentX        =   21273
         _ExtentY        =   3995
         _Version        =   393216
         Cols            =   3
         FixedCols       =   0
         BackColorFixed  =   -2147483646
         ForeColorFixed  =   -2147483639
         BackColorSel    =   -2147483643
         ForeColorSel    =   -2147483646
         BackColorBkg    =   -2147483636
         GridColor       =   0
         _NumberOfBands  =   1
         _Band(0).Cols   =   3
         _Band(0).GridLinesBand=   0
         _Band(0).TextStyleBand=   0
         _Band(0).TextStyleHeader=   0
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid grid 
         Height          =   6135
         Index           =   4
         Left            =   -74760
         TabIndex        =   20
         Top             =   600
         Width           =   11190
         _ExtentX        =   19738
         _ExtentY        =   10821
         _Version        =   393216
         Cols            =   3
         BackColorFixed  =   -2147483646
         ForeColorFixed  =   -2147483639
         BackColorSel    =   -2147483643
         ForeColorSel    =   -2147483646
         BackColorBkg    =   -2147483636
         GridColor       =   -2147483638
         AllowUserResizing=   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _NumberOfBands  =   1
         _Band(0).Cols   =   3
      End
      Begin TabDlg.SSTab SSTab2 
         Height          =   4785
         Left            =   -74640
         TabIndex        =   32
         Top             =   2865
         Width           =   12030
         _ExtentX        =   21220
         _ExtentY        =   8440
         _Version        =   393216
         Tabs            =   2
         TabsPerRow      =   2
         TabHeight       =   520
         ShowFocusRect   =   0   'False
         TabCaption(0)   =   "Gráfico Promedio de Facturación"
         TabPicture(0)   =   "frmIEco.frx":03A2
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "Graf"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "pic(0)"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).Control(2)=   "txtProm"
         Tab(0).Control(2).Enabled=   0   'False
         Tab(0).ControlCount=   3
         TabCaption(1)   =   "Gráfico Promedio de Gastos"
         TabPicture(1)   =   "frmIEco.frx":03BE
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "graf1"
         Tab(1).Control(1)=   "pic(1)"
         Tab(1).ControlCount=   2
         Begin VB.TextBox txtProm 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   1080
            TabIndex        =   40
            Top             =   720
            Visible         =   0   'False
            Width           =   1995
         End
         Begin VB.PictureBox pic 
            AutoSize        =   -1  'True
            Height          =   1815
            Index           =   1
            Left            =   -74850
            Picture         =   "frmIEco.frx":03DA
            ScaleHeight     =   1755
            ScaleWidth      =   3900
            TabIndex        =   39
            Top             =   1260
            Visible         =   0   'False
            Width           =   3960
         End
         Begin VB.PictureBox pic 
            AutoSize        =   -1  'True
            Height          =   1035
            Index           =   0
            Left            =   7185
            Picture         =   "frmIEco.frx":2D3F
            ScaleHeight     =   975
            ScaleWidth      =   3960
            TabIndex        =   38
            Top             =   405
            Visible         =   0   'False
            Width           =   4020
         End
         Begin MSChart20Lib.MSChart graf1 
            CausesValidation=   0   'False
            Height          =   4275
            Left            =   -74685
            OleObjectBlob   =   "frmIEco.frx":4F45
            TabIndex        =   34
            Top             =   390
            Width           =   10935
         End
         Begin MSChart20Lib.MSChart Graf 
            Height          =   4170
            Left            =   300
            OleObjectBlob   =   "frmIEco.frx":6CF0
            TabIndex        =   33
            Top             =   360
            Width           =   10935
         End
      End
   End
   Begin VB.Image Img 
      Enabled         =   0   'False
      Height          =   150
      Index           =   1
      Left            =   315
      Top             =   0
      Visible         =   0   'False
      Width           =   165
   End
   Begin VB.Image Img 
      Enabled         =   0   'False
      Height          =   150
      Index           =   0
      Left            =   0
      Top             =   15
      Visible         =   0   'False
      Width           =   165
   End
End
Attribute VB_Name = "frmIEco"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------------
    '   módulo condominio
    '
    '   ventana Informe económico.
    '
    '   Permite al ususario diseñar el informe económico siguiendo los parámetros
    '   dados por este.-
    '---------------------------------------------------------------------------------------------
    'variables locales a nivel de módulo
    Dim rstCGastos As ADODB.Recordset
    Enum Ruburo
        cargado = 2
        Recaudado
        xRecaudar
        Pagado
        Saldo
    End Enum
    
    Private Sub cmdIEco_Click(Index As Integer)
    'matriz de controles command
    Select Case Index
    
        Case 0: Call IE 'informe económico
        
        Case 1: Call View 'VER
            Call mostrarInformeEconomico
            
        Case 2: Unload Me 'cerrar formulario
        
        Case 3: Call PrinterIE  'imprime el informe económico
        
        Case 4 'exportar a una hoja de calculo
        
        Case 5  'Imprimir informe económico
            Call Imprimir_ieco
        
    End Select
    '
    End Sub

    '---------------------------------------------------------------------------------------------
    '   Rutina:     View
    '
    '   Genera un archivo en formato .html y lo muestra en el browser
    '---------------------------------------------------------------------------------------------
    Private Sub View()
    'variables locales
    Dim strSQL As String, strIN As String
    Dim rstlocal As New ADODB.Recordset
    Dim Fila As Long
    Const m$ = "#,##0.00"
    Dim curTFondo@, curTemp@
    Dim varDatos()
    Dim maxFact As Currency, Por_Cobrar As Currency
    Dim datDesde As Date, datHasta As Date
    Dim rstTemp As New ADODB.Recordset
    Dim strFiltro As String, Total As Currency, Saldo_Fecha As Currency
    Dim GranTotal As Currency
    '
    MousePointer = vbHourglass
    
    'selecciona las cuentas de fondos que aparecerán en el informe
    Grid(0).Col = 1
    
    For I = 1 To Grid(0).Rows - 1
        Grid(0).Row = I
        If Grid(0).CellPicture = img(1) Then
            strIN = strIN & IIf(strIN = "", "('", "','") & Grid(0).TextMatrix(I, 0)
        End If
    Next I
    '
    Fila = 1
    Grid(3).Rows = 2
    
    Call rtnLimpiar_Grid(Grid(3))
    
    If strIN <> "" Then
    
        strIN = strIN & "')"
        strSQL = "SELECT * FROM Tgastos WHERE CodGasto IN " & strIN
        rstlocal.Open strSQL, cnnOLEDB + mcDatos, adOpenKeyset, adLockOptimistic, adCmdText
        If Not (rstlocal.EOF And rstlocal.BOF) Then
            rstlocal.MoveFirst
            Do
                Grid(3).TextMatrix(Fila, 0) = rstlocal!Titulo
                Grid(3).TextMatrix(Fila, 1) = Format(rstlocal!SaldoActual, m)
                curTFondo = curTFondo + rstlocal!SaldoActual
                rstlocal.MoveNext
                Grid(3).AddItem ("")
                Fila = Fila + 1
            Loop Until rstlocal.EOF
        End If
        rstlocal.Close
    End If
    '
    'seleeciona los intereses cargados (menos) intereses descontados
    If chk.Value = vbChecked Then
    
        strSQL = "SELECT Sum(Monto) FROM DetFact WHERE Fecha Between #" & _
        Format(dtp(0), "mm/dd/yy") & "# AND #" & Format(Date, "mm/dd/yy") & "# AND CodGasto=(SELEC" _
        & "T CodIntMora FROM Inmueble IN '" & gcPath & "\sac.mdb' WHERE CodInm='" & gcCodInm & _
        "')"
        rstlocal.Open strSQL, cnnOLEDB + mcDatos, adOpenKeyset, adLockOptimistic, adCmdText
        If Not IsNull(rstlocal.Fields(0)) Then
            curTemp = rstlocal.Fields(0)
        End If
        rstlocal.Close
        'le resta lo descontado
        strSQL = "SELECT Sum(Monto) FROM DetFact WHERE Fecha Between #" & _
        Format(dtp(0), "mm/dd/yy") & "# AND #" & Format(Date, "mm/dd/yy") & "# AND CodGasto=(SELEC" _
        & "T Left(CodIntMora,2) & 2 " _
        & " & Right(CodIntMora,3) FROM Inmueble IN '" & gcPath & "\sac.mdb' WHERE CodInm='" & _
        gcCodInm & "')"
        rstlocal.Open strSQL, cnnOLEDB + mcDatos, adOpenKeyset, adLockOptimistic, adCmdText
        If Not IsNull(rstlocal.Fields(0)) Then
            curTemp = curTemp + rstlocal.Fields(0)
        End If
        rstlocal.Close
        Grid(3).TextMatrix(Fila, 0) = "INT. MORATORIOS CARGADOS " & Format(dtp(0), "dd/mm/yyyy") _
        & " AL " & Format(Date, "dd/mm/yyyy")
        Grid(3).TextMatrix(Fila, 1) = Format(curTemp, m)
        Fila = Fila + 1
        Grid(3).AddItem ("")
        '
        curTFondo = curTFondo + curTemp
    
    End If
    '
    'selecciona las cuentas de cuotas especiales
    Grid(1).Col = 1
    
    For I = 1 To Grid(1).Rows - 1
    
        Grid(1).Row = I
        
        If Grid(1).CellPicture = img(1) Then
            'strIN = strIN & IIf(strIN = "", "('", "','") & Grid(1).TextMatrix(i, 0)
            '
            
            rstlocal.Open "SELECT Sum(MovFondo.Debe) AS D,Sum(MovFondo.Haber) AS H,Tgastos.Titu" _
            & "lo FROM MovFondo INNER JOIN Tgastos ON MovFondo.CodGasto = Tgastos.CodGasto WHER" _
            & "E MovFondo.CodGasto ='" & Grid(1).TextMatrix(I, 0) & "' AND MovFondo.Del= False GROUP BY MovFondo.CodGas" _
            & "to, Tgastos.Titulo;", cnnOLEDB + mcDatos, adOpenKeyset, adLockOptimistic, adCmdText
            If Not (rstlocal.EOF And rstlocal.BOF) Then
            '
                Saldo_Fecha = rstlocal("H") - rstlocal("D")
                '
                Grid(3).TextMatrix(Fila, 0) = "SALDO A LA FECHA " & rstlocal("Titulo")
                Grid(3).TextMatrix(Fila, 1) = Format(Saldo_Fecha, m)
                Fila = Fila + 1
                curTFondo = curTFondo + Saldo_Fecha
                Grid(3).AddItem ("")
                rstlocal.Close
                'monto por recaudar
                
                rstlocal.Open "SELECT Sum(DF.Monto) AS Total FROM DetFact AS DF INNER JOIN Factura " _
                & "AS F ON DF.Fact = F.FACT WHERE F.Saldo>0 AND DF.CodGasto='" & _
                Grid(1).TextMatrix(I, 0) & "';"
                If Not (rstlocal.EOF And rstlocal.BOF) Then Por_Cobrar = IIf(IsNull(rstlocal("Total")), 0, rstlocal("Total"))
            End If
            rstlocal.Close
            
            
        End If
        
    Next I
    '
    Grid(3).TextMatrix(Fila, 0) = "TOTAL FONDOS"
    Grid(3).TextMatrix(Fila, 2) = Format(curTFondo, m)
    Grid(3).AddItem ("")
    Grid(3).AddItem ("")
    Fila = Fila + 2
    '
    'imprime la deuda actual del condominio
    rstlocal.Open "SELECT * FROM Inmueble WHERE CodInm='" & gcCodInm & "'", _
    cnnConexion, adOpenKeyset, adLockOptimistic, adCmdText
    curTemp = rstlocal!DeudaAct
    rstlocal.Close
    
    '
    Grid(3).TextMatrix(Fila, 0) = "DEUDA DE CONDOMINIO"
    Grid(3).TextMatrix(Fila, 1) = Format(curTemp, m)
    Grid(3).AddItem ("")
    Fila = Fila + 1
    '
    'imprime monTo por facturar
    rstlocal.Open "SELECT Sum(Monto) FROM AsignaGasto WHERE Cargado IN (SELECT dateadd('m',1,Ma" _
    & "x(Periodo)) FROM Factura WHERE Fact Not LIKE 'CH%')", cnnOLEDB + mcDatos, adOpenKeyset, _
    adLockOptimistic, adCmdText
    If Not IsNull(rstlocal.Fields(0)) Then
        Grid(3).TextMatrix(Fila, 0) = "GASTOS POR FACTURAR"
        Grid(3).TextMatrix(Fila, 1) = Format(CLng(rstlocal.Fields(0)), "#,##0.00")
        curTemp = curTemp + CLng(rstlocal.Fields(0))
        Grid(3).AddItem ("")
        Fila = Fila + 1
    End If
    rstlocal.Close
    '
    Grid(3).TextMatrix(Fila, 0) = "DEUDA TOTAL CONDOMINIO"
    Grid(3).TextMatrix(Fila, 2) = Format(curTemp, m)
    Grid(3).AddItem ("")
    Fila = Fila + 1
    Grid(3).TextMatrix(Fila, 0) = "TOTAL FONDO EDIFICIO"
    Grid(3).TextMatrix(Fila, 2) = Format(curTFondo - curTemp, m)
    Grid(3).AddItem ("")
    Fila = Fila + 1
    '
    'selecciona el detalle de las cuentas de cuotas especiales
    Grid(1).Col = 1
    
    For I = 1 To Grid(1).Rows - 1
    
        Grid(1).Row = I
        
        If Grid(1).CellPicture = img(1) Then
            '
            
            rstlocal.Open "SELECT Sum(MovFondo.Debe) AS D,Sum(MovFondo.Haber) AS H,Tgastos.Titu" _
            & "lo FROM MovFondo INNER JOIN Tgastos ON MovFondo.CodGasto = Tgastos.CodGasto WHER" _
            & "E MovFondo.CodGasto ='" & Grid(1).TextMatrix(I, 0) & "' AND MovFondo.Del= False GROUP BY MovFondo.CodGas" _
            & "to, Tgastos.Titulo;", cnnOLEDB + mcDatos, adOpenKeyset, adLockOptimistic, adCmdText
            If Not (rstlocal.EOF And rstlocal.BOF) Then
                Grid(3).AddItem ("")
                Fila = Fila + 1
                '
                Grid(3).Row = Fila
                Grid(3).Col = 0
                
                Grid(3).TextMatrix(Fila, 0) = "CUOTA ESPECIAL " & rstlocal("Titulo")
                Grid(3).CellFontBold = True
                Grid(3).AddItem ("")
                Fila = Fila + 1
                '
                'total credito
                Grid(3).TextMatrix(Fila, 0) = "TOTAL CREDITO"
                Grid(3).TextMatrix(Fila, 1) = Format(rstlocal("H"), m)
                Grid(3).AddItem ("")
                Fila = Fila + 1
                '
                'todal debito
                Grid(3).TextMatrix(Fila, 0) = "TOTAL DEBITO"
                Grid(3).TextMatrix(Fila, 1) = Format(rstlocal("d"), m)
                Grid(3).AddItem ("")
                Fila = Fila + 1
                '
                'saldo a la fecha
                Saldo_Fecha = rstlocal("H") - rstlocal("D")
                Grid(3).TextMatrix(Fila, 0) = "SALDO A LA FECHA"
                Grid(3).TextMatrix(Fila, 1) = Format(Saldo_Fecha, m)
                Grid(3).AddItem ("")
                Fila = Fila + 1
                '
            
                'monto por recaudar
                rstlocal.Close
                rstlocal.Open "SELECT Sum(DF.Monto) AS Total FROM DetFact AS DF INNER JOIN Factura " _
                & "AS F ON DF.Fact = F.FACT WHERE F.Saldo>0 AND DF.CodGasto='" & _
                Grid(1).TextMatrix(I, 0) & "';"
                If Not rstlocal.EOF And Not rstlocal.BOF Then Por_Cobrar = IIf(IsNull(rstlocal("Total")), 0, rstlocal("Total"))
                'por recaudar
                Grid(3).TextMatrix(Fila, 0) = "SALDO POR RECAUDAR"
                Grid(3).TextMatrix(Fila, 1) = Format(Por_Cobrar, m)
                Grid(3).AddItem ("")
                Fila = Fila + 1
                '
                'SALDO REAL
                Grid(3).TextMatrix(Fila, 0) = "SALDO REAL"
                Grid(3).TextMatrix(Fila, 1) = Format(Saldo_Fecha - Por_Cobrar, m)
                Grid(3).AddItem ("")
                Fila = Fila + 1
            End If
            rstlocal.Close
            
            
            
        End If
        
    Next I
    Grid(3).Col = 2
    Grid(3).ColSel = 2
    Grid(3).Row = 1
    Grid(3).RowSel = Grid(3).Rows - 1

    Grid(3).FillStyle = flexFillRepeat
    Grid(3).CellFontBold = True
    Grid(3).Col = 0
    Grid(3).Row = 1
    
    'selecciona ahora el total facturado para un determinado N períodos
    If txtIEco = "" Then txtIEco = 0
    
    If txtIEco > 0 Then
        
        'strSql = "SELECT TOP " & txtIEco & " Periodo, Sum(Facturado) AS TM From Factura WHERE Fact Not Like 'CH%' GRO" _
        & "UP BY Periodo ORDER BY Periodo DESC"
        strSQL = "SELECT Max(Periodo) as M FROM Factura WHERE Fact Not Like 'CH%'"
        rstlocal.Open strSQL, cnnOLEDB + mcDatos, adOpenDynamic, adLockReadOnly, adCmdText
        
        datHasta = rstlocal("M")
        datDesde = DateAdd("m", (txtIEco - 1) * -1, datHasta)
        datDesde = Format(datDesde, "MM/DD/YY")
        datHasta = Format(datHasta, "MM/DD/YY")
        rstlocal.Close
        
        strSQL = "SELECT Cargado, Sum(Monto) AS Facturado FROM AsignaGasto WHERE Cargado BETWEEN #" & datDesde & "# and #" & datHasta & "# GROUP BY Cargado ORDER BY Cargado"

        rstlocal.Open strSQL, cnnOLEDB + mcDatos, adOpenKeyset, adLockOptimistic, adCmdText
        
        Graf.ColumnCount = 1
        Graf.RowCount = txtIEco
        '
        If Not rstlocal.EOF And Not rstlocal.BOF Then
        '
            rstlocal.MoveFirst: I = 1
            
            '
            Do
                
                Graf.Row = I
                Graf.data = Format(rstlocal("Facturado"), "#,##0.00")
                Graf.RowLabel = UCase(Format(rstlocal("Cargado"), "mmm'yy"))
                maxFact = IIf((rstlocal("Facturado")) > maxFact, rstlocal("Facturado"), maxFact)
                Graf.Plot.SeriesCollection(1).DataPoints(-1).DataPointLabel.LocationType = VtChLabelLocationTypeInside
                Graf.Plot.SeriesCollection(1).DataPoints(-1).DataPointLabel.Component = VtChLabelComponentValue
                GranTotal = GranTotal + rstlocal("Facturado")
                rstlocal.MoveNext: I = I + 1
                
                
            Loop Until rstlocal.EOF
            GranTotal = GranTotal / (I - 1)
            txtProm = Format(GranTotal, "#,##0.00")
            If (maxFact - Fix(maxFact)) < 0.5 Then
            maxFact = Fix(maxFact) + 0.6
            End If
            maxFact = CLng(maxFact)
'            For i = maxFact To 2 Step -2
'                K = i - 2
'            Next
            'If K = 1 Then maxFact = maxFact + 1
            For I = 10 To 1 Step -1
                If maxFact Mod I = 0 Then K = I: Exit For
            Next
            Graf.Plot.Axis(VtChAxisIdY).ValueScale.Auto = False
            Graf.Plot.Axis(VtChAxisIdY).ValueScale.Maximum = maxFact
            Graf.Plot.Axis(VtChAxisIdY).ValueScale.Minimum = 0
            Graf.Plot.Axis(VtChAxisIdY).ValueScale.MajorDivision = K
            Graf.FootnoteText = "Expresado en Bs." '/ Promedio de Facturacion Bs. " & Format(GranTotal, "#,##0.00")
            GranTotal = 0
        End If
        
        rstlocal.Close
        '
        'grafica el promedio de gastos de los ùltimos tres perìodos
        
        
        '
        With Grid(2)
            .Col = 1
            For I = 1 To .Rows - 1
                .Row = I
                '
                If .CellPicture = img(1) Then
                '
                    strFiltro = strFiltro & IIf(strFiltro = "", "(", " or ") & "AsignaGasto.Cod" _
                    & "Gasto = '" & .TextMatrix(I, 0) & "'"
                    '
                End If
                
            Next
            strFiltro = IIf(strFiltro = "", "", strFiltro + ")")
            '
        End With
'        strSql = "SELECT Max(Periodo) as M FROM Factura WHERE Fact Not Like 'CH%'"
'        rstTemp.Open strSql, cnnOLEDB + mcDatos, adOpenDynamic, adLockReadOnly, adCmdText
'
'        datHasta = rstTemp("M")
        datHasta = Format(datHasta, "MM/DD/YY")
        datDesde = DateAdd("m", -2, datHasta)
        datHasta = Format(datHasta, "mm/dd/yy")
        datDesde = Format(datDesde, "mm/dd/yy")
        'rstTemp.Close
        
        strSQL = "TRANSFORM Sum(AsignaGasto.Monto) AS SubTotal SELECT AsignaGasto.CodGasto, Tga" _
        & "stos.Titulo AS Detalle, Sum(AsignaGasto.Monto) AS Total FROM AsignaGasto LEFT JOIN T" _
        & "gastos ON AsignaGasto.CodGasto = Tgastos.CodGasto Where (((AsignaGasto.Cargado) >= #" _
        & datDesde & "# And (AsignaGasto.Cargado) <= #" & datHasta & "#)) GROUP BY AsignaGasto." _
        & "CodGasto, Tgastos.Titulo PIVOT AsignaGasto.Cargado;"
        
        rstTemp.Open strSQL, cnnOLEDB + mcDatos, adOpenKeyset, adLockOptimistic, adCmdText
        
        If Not rstTemp.EOF And Not rstTemp.BOF Then
            rstTemp.MoveFirst
            Do
                GranTotal = GranTotal + rstTemp("Total")
                rstTemp.MoveNext
            Loop Until rstTemp.EOF
        End If
        'GranTotal = CLng(GranTotal)
        rstTemp.Close
      If strFiltro <> "" Then
        '
        strSQL = "TRANSFORM Sum(AsignaGasto.Monto) AS SubTotal SELECT AsignaGasto.CodGasto, Tga" _
        & "stos.Titulo AS Detalle, Sum(AsignaGasto.Monto) AS Total FROM AsignaGasto LEFT JO" _
        & "IN Tgastos ON AsignaGasto.CodGasto = Tgastos.CodGasto WHERE (AsignaGasto.Cargado >= " _
        & "#" & datDesde & "# And AsignaGasto.Cargado <= #" & datHasta & "#) " & _
        IIf(strFiltro = "", "", " AND " & strFiltro) & " GROUP BY AsignaGasto.CodGasto, Tgastos.Titulo  ORDER" _
        & " BY AsignaGAsto.CodGasto ASC PIVOT AsignaGasto.Cargado;"
        '
        rstTemp.CursorLocation = adUseClient
        
        rstTemp.Open strSQL, cnnOLEDB + mcDatos, adOpenKeyset, adLockOptimistic, adCmdText
        'rstTemp.Sort = "total ASC"
        graf1.Column = 1
                '
        If Not rstTemp.EOF And Not rstTemp.BOF Then
            graf1.Column = 1    'series
            graf1.RowCount = rstTemp.RecordCount + 1    'punto de datos

            
            rstTemp.MoveFirst: I = 1
            Do
              graf1.Row = I
              graf1.data = rstTemp("Total") * 100 / GranTotal
              
              graf1.RowLabel = Descrip(rstTemp("Detalle")) & Format(graf1.data, " #,##0.00 ") & "%"
              graf1.Plot.SeriesCollection(I).DataPoints(-1).DataPointLabel.VtFont.Size = 7
              graf1.Plot.SeriesCollection(I).DataPoints(-1).DataPointLabel.LocationType = VtChLabelLocationTypeOutside
              graf1.Plot.SeriesCollection(I).DataPoints(-1).DataPointLabel.Component = VtChLabelComponentSeriesName
              
'              With graf1.Plot.SeriesCollection(1).DataPoints(-1)
'                .Brush.Style = VtBrushStyleNull
'
'                '.DataPointLabel.Text = 0
'              End With
              rstTemp.MoveNext: I = I + 1
                
            Loop Until rstTemp.EOF
            
            rstTemp.Close
      
            '
            strSQL = "TRANSFORM Sum(AsignaGasto.Monto) AS SubTotal SELECT AsignaGasto.CodGasto," _
            & "Tgastos.Titulo AS Detalle, Sum(AsignaGasto.Monto) AS Total FROM AsignaGasto " _
            & "LEFT JOIN Tgastos ON AsignaGasto.CodGasto = Tgastos.CodGasto WHERE (AsignaGasto." _
            & "Cargado >= #" & datDesde & "# And AsignaGasto.Cargado <= #" & datHasta & "#) " & _
            IIf(strFiltro = "", "", "AND Not " & strFiltro) & " GROUP BY AsignaGasto.CodGasto, Tgas" _
            & "tos.Titulo  ORDER BY AsignaGAsto.CodGasto ASC PIVOT AsignaGasto.Cargado;"
            
            rstTemp.Open strSQL, cnnOLEDB + mcDatos, adOpenKeyset, adLockOptimistic, adCmdText
            rstTemp.MoveFirst
            Do
                Total = Total + rstTemp("Total")
                rstTemp.MoveNext
            Loop Until rstTemp.EOF
            graf1.Row = I
            graf1.data = Total * 100 / GranTotal
            graf1.RowLabel = "OTROS " & Format(graf1.data, " #,##0.00 ") & "%"
            graf1.Plot.SeriesCollection(I).DataPoints(-1).DataPointLabel.VtFont.Size = 7
            graf1.Plot.SeriesCollection(I).DataPoints(-1).DataPointLabel.LocationType = VtChLabelLocationTypeOutside
            graf1.Plot.SeriesCollection(I).DataPoints(-1).DataPointLabel.Component = VtChLabelComponentSeriesName
            rstTemp.Close
            
            '
            graf1.Plot.Axis(VtChAxisIdY).ValueScale.Auto = False
            
            graf1.Plot.Axis(VtChAxisIdY).ValueScale.Maximum = 40
            graf1.Plot.Axis(VtChAxisIdY).ValueScale.Minimum = 0
            graf1.Plot.Axis(VtChAxisIdY).ValueScale.MajorDivision = 5
            graf1.Plot.Axis(VtChAxisIdY).ValueScale.MinorDivision = 0
            graf1.Visible = True
        Else
            graf1.Visible = False
        End If
        
      Else
        graf1.Visible = False
        
      End If
        '
        Set rstlocal = Nothing
        Set rstTemp = Nothing
        '
    End If
    '
    MousePointer = vbDefault
    End Sub


    Private Sub Form_Load()
    '
    SSTab1.TabEnabled(2) = False
    
    img(0).Picture = LoadResPicture("UnChecked", vbResBitmap)
    img(1).Picture = LoadResPicture("Checked", vbResBitmap)
    'crea una instrancia del objeto ADODB.Recordset
    Set rstCGastos = New ADODB.Recordset
    rstCGastos.Open "SELECT CodGasto,'',Titulo FROM Tgastos WHERE Fondo=True", cnnOLEDB + _
    mcDatos, adOpenKeyset, adLockOptimistic, adCmdText
    Set Grid(0).DataSource = rstCGastos
    Set Grid(1).DataSource = rstCGastos
    '
    'crea una nueva instancia del mismo objeto
    Set rstCGastos = New ADODB.Recordset
    rstCGastos.Open "SELECT DISTINCT CodGasto,'' FROM AsignaGasto WHERE Len(CodGasto)>4", _
    cnnOLEDB + mcDatos, adOpenKeyset, adLockOptimistic, adCmdText
    '
    Set Grid(2).DataSource = rstCGastos
    '
    For I = 0 To 2
        Set Grid(I).FontFixed = LetraTitulo(LoadResString(527), 7.5, , True)
        Set Grid(I).Font = LetraTitulo(LoadResString(528), 8)
        Call centra_titulo(Grid(I), True)
        Grid(I).Col = 1
        For K = 1 To Grid(I).Rows - 1   'IMAGEN
            Grid(I).Row = K
            Set Grid(I).CellPicture = img(0)
            Grid(I).CellPictureAlignment = flexAlignCenterCenter
        Next K
    Next I
    Grid(3).ColWidth(0) = 7000
    Grid(3).ColWidth(1) = 1500
    Grid(3).ColWidth(2) = 1500
    dtp(1) = Date
    '
    End Sub

    Private Sub Form_Resize()
    'configura la presentacion en la ventana del formulario
    With SSTab1
        .Height = Me.ScaleHeight - SSTab1.top
        fraIECO(0).Height = .Height - fraIECO(0).top - 200
    End With
    End Sub

    Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    rstCGastos.Close
    Set rstCGastos = Nothing
    Set frmIEco = Nothing
    End Sub


    Private Sub grid_Click(Index As Integer)
    If Grid(Index).ColSel = 1 And Index < 3 Then
        Grid(Index).Row = Grid(Index).RowSel
        Set Grid(Index).CellPicture = IIf(Grid(Index).CellPicture = img(0), img(1), img(0))
        Grid(Index).CellPictureAlignment = flexAlignCenterCenter
    End If
    End Sub


    Private Sub MskFecha_KeyPress(Index As Integer, KeyAscii As Integer)
    Call Validacion(KeyAscii, "0123456789")
    If KeyAscii = 13 Then
        If Index = 0 Then
            MskFecha(1).SetFocus
        Else
            
        End If
    End If
    End Sub

    Private Sub optIECO_Click(Index As Integer): Call IE(optIECO(Index).Tag)
    End Sub

    Private Sub SSTab1_Click(PreviousTab As Integer)
    'VARIABLES LOCALES
    Select Case SSTab1.tab
        Case 0: SSTab1.TabEnabled(2) = False
        Case 1: SSTab1.TabEnabled(2) = False
        Case 2
    End Select
    '
    End Sub

    Private Sub txtCopy_KeyPress(Index%, KeyAscii%)
    Select Case Index
        Case 2, 4
            KEYASCCi = Asc(UCase(Chr(KeyAscii)))
        Case 0, 1
            If KeyAscii > 26 Then If InStr("0123456789", Chr(KeyAscii)) = 0 Then KeyAscii = 0
        Case 3, 5
            If KeyAscii > 26 Then If InStr("0123456789,-.", Chr(KeyAscii)) = 0 Then KeyAscii = 0
    End Select
    End Sub

    Private Sub txtIEco_KeyPress(KeyAscii As Integer)
    Call Validacion(KeyAscii, "0123456789")
    End Sub

    '---------------------------------------------------------------------------------------------
    '   Rutina: IE
    '
    '   Genera la consulta de datos del informe económico según parámetros seleccionados
    '---------------------------------------------------------------------------------------------
    Private Sub IE(Optional Filtro As String)
    'variables locales
    Dim strSQL As String
    Dim Desde$, Hasta$
    '
    Desde = "01/" & MskFecha(0)
    Hasta = "01/" & MskFecha(1)
    'valida los datos necesario
    If Not IsDate("01/" & MskFecha(0)) Or Not IsDate("01/" & MskFecha(1)) Then
        MsgBox "Verifique los datos que introdujo en el paso Nº 6 e intentelo nuevamente", _
        vbInformation, App.ProductName
        Exit Sub
    End If
    '
    Desde = Format(Desde, "mm/dd/yy")
    Hasta = Format(Hasta, "mm/dd/yy")
    '
'    '***********************************************
'    '   Si el inmueble es el 2553
'    '   eliminamos los registros del gasto 910002
'    '   de la tabla AsignaGasto
'    '***********************************************
'    If gcCodInm = "2553" Then
'        strSQL = "DELETE FROM AsignaGasto IN '" + mcDatos + "' WHERE CodGasto='910002'"
'        cnnConexion.Execute strSQL, N
'    End If
    
    strSQL = "TRANSFORM Sum(AsignaGasto.Monto) AS SubTotal SELECT AsignaGasto.CodGasto, Tgastos" _
    & ".Titulo AS Detalle, Sum(AsignaGasto.Monto) AS Total FROM AsignaGasto LEFT JOIN Tgastos " _
    & "ON AsignaGasto.CodGasto = Tgastos.CodGasto WHERE AsignaGasto.Cargado>= #" & _
    Desde & "# And AsignaGasto.Cargado<=#" & Hasta & "# " & Filtro & " GROUP BY Asi" _
    & "gnaGasto.CodGasto, Tgastos.Titulo PIVOT AsignaGasto.Cargado;"
    'strSql = "TRANSFORM Sum(DetFact.Monto) AS SubTotal SELECT DetFact.CodGasto, Tgastos.Titulo " _
    & " AS Detalle, Sum(DetFact.Monto) AS Total FROM DetFact LEFT JOIN Tgastos ON DetFact.CodGa" _
    & "sto = Tgastos.CodGasto Where (DetFact.Periodo >= #" & Desde & "# And DetFact.Periodo" _
    & " <= #" & Hasta & "#) AND DetFact.CodGasto Not In (SELECT CodGasto FROM GastoIECO IN '" & gcPath & "\sac.mdb') " & Filtro & " GROUP BY DetFact.CodGasto, Tgastos.Titulo PIVOT DetFact.Periodo;"

    '
    'crea la consulta
    Call rtnGenerator(mcDatos, strSQL, "IEconomico")
    Call Grid_IE
    SSTab1.TabEnabled(2) = True
    SSTab1.tab = 2
    End Sub

    '---------------------------------------------------------------------------------------------
    '   Rutina PrinterIE
    '
    '   Esta rutina configura la impresión del informe económico
    '---------------------------------------------------------------------------------------------
    Sub PrinterIE()
    Dim rst As New ADODB.Recordset
    Dim Periodo$, Linea%, subT$
    Dim Pos&, Monto$, INI%, Fin%, Left%, Contador&
    Dim varTotal()
    '
    rst.Open "ieconomico", cnnOLEDB & mcDatos, adOpenKeyset, adLockOptimistic, adCmdTable
    
    If Not rst.EOF And Not rst.BOF Then
        'configura la presentacion en el papel
        
        Periodo = UCase("Económico desde: " & MskFecha(0) & " Hasta:" & MskFecha(1))
        '
        Printer.Orientation = vbPRORLandscape
        Printer.TrackDefault = True
        If txtCopy(1) = "" Then txtCopy(1) = 0
        Printer.Copies = IIf(txtCopy(1) = 0, 1, txtCopy(1))
        
        Printer.FontName = "Times New Roman"
        Printer.FontBold = True
        If rst.RecordCount + 7 < 81 Then
            For I = 1 To CLng((81 - (rst.RecordCount + 7)) / 2)
                Printer.Print
            Next
        Else
            Printer.Print
        End If
        Left = (Printer.ScaleWidth - ((rst.Fields.Count - 2) * 900 + 3400)) / 2
        If Left <= 270 Then Left = 100
        For I = 0 To 2
            If optIECO(I) Then subT = "INFORME ECONOMICO (" & UCase(optIECO(I).Caption) & ")"
        Next
        'Printer.Print
        Printer.CurrentX = (Printer.ScaleWidth - Printer.TextWidth(subT)) / 2
        Printer.Print subT
        '
        Printer.FontName = "Arial Narrow"
        Printer.FontSize = 6
        Printer.CurrentX = (Printer.ScaleWidth - Printer.TextWidth(UCase(gcNomInm))) / 2
        Printer.Print UCase(gcNomInm)
        Printer.CurrentX = (Printer.ScaleWidth - Printer.TextWidth(Periodo)) / 2
        Printer.Print Periodo
        If rst.RecordCount + 7 < 81 Then Printer.Print
        '
        'imprime el detalle
        With rst
            ReDim varTotal(.Fields.Count)
            .MoveFirst
            INI = Printer.CurrentY
            'Printer.Line (Printer.ScaleLeft, Printer.CurrentY)-(Printer.ScaleWidth, Printer.CurrentY)
            'Printer.Line (Printer.ScaleLeft, Printer.CurrentY)-(Printer.ScaleWidth, _
            Printer.CurrentY + (TextHeight(strTitulo) / 2)), , BF
            Pos = Left + 3600 + ((.Fields.Count - 2) * 900)
            Printer.Line (150 + Left, INI)-(Pos, INI + 120), 8421504, BF
            'Pos = 0-
            Linea = INI
            Printer.CurrentX = Left + (550 - Printer.TextWidth("COD.")) / 2 + 100
            Printer.CurrentY = Linea
            Printer.Print "COD."
            '
            Printer.CurrentX = Left + (3600 - Printer.TextWidth("DESCRIPCION")) / 2
            Printer.CurrentY = Linea
            Printer.Print "DESCRIPCION"
            '
            Pos = 4500
            For I = 3 To .Fields.Count - 1
                
                Monto = Format(.Fields(I).Name, "mm'yy")
                Printer.CurrentY = Linea
                Printer.CurrentX = Left + ((Pos - 900) + (Pos - (Pos - 900)) / 2) - (Printer.TextWidth(Monto) / 2)
                Printer.Print Monto
                Pos = Pos + 900
            Next
            Printer.CurrentY = Linea
            Printer.CurrentX = Left + ((Pos - 900) + (Pos - (Pos - 900)) / 2) - (Printer.TextWidth("TOTAL") / 2)
            Printer.Print "TOTAL"
            Printer.FontBold = False
            Do
                'If !CodGasto = "212007" Then Stop
                Linea = Printer.CurrentY
                Printer.CurrentX = Left + 200
                Printer.CurrentY = Linea
                Printer.Print !codGasto
                '
                Printer.CurrentY = Linea
                Printer.CurrentX = Left + 600
                Printer.Print Mid(!Detalle, 1, 50)
                '
                Pos = 4400
                For I = 3 To .Fields.Count - 1
                    Monto = Format(IIf(IsNull(.Fields(I)), 0, .Fields(I)), "#,##0.00")
                    Printer.CurrentY = Linea
                    Printer.CurrentX = Left + Pos - Printer.TextWidth(Monto)
                    Printer.Print Monto
                    varTotal(I) = CCur(Monto) + CCur(varTotal(I))
                    Pos = Pos + 900
                Next
                Printer.FontBold = True
                varTotal(2) = varTotal(2) + IIf(IsNull(.Fields(2)), 0, .Fields(2))
                Monto = Format(IIf(IsNull(.Fields(2)), 0, .Fields(2)), "#,##0.00")
                Printer.CurrentY = Linea
                Printer.CurrentX = Left + Pos - Printer.TextWidth(Monto)
                Printer.Print Monto
                Printer.FontBold = False
                'Printer.Line (Printer.ScaleLeft, Printer.CurrentY)-(Printer.ScaleWidth, Printer.CurrentY)
                .MoveNext
                Linea = Linea + 1
                Contador = Contador + 1
                If Contador >= 75 Then
                    GoSub Cuadricula
                    Printer.NewPage
                    Printer.CurrentY = 653
                    Contador = 0
                    'imprime las cuadriculas
                   
                End If
                
            Loop Until .EOF
            GoSub Cuadricula
            Printer.EndDoc
            MsgBox "Reporte impreso con éxito", vbInformation, App.ProductName
            Exit Sub
Cuadricula:
            Linea = Printer.CurrentY
            'totaliza las columnas
            Pos = 4400
            Printer.FontBold = True
            Printer.CurrentX = (3600 - Printer.TextWidth("TOTAL"))
            Printer.CurrentY = Linea
            Printer.Print "TOTAL"
            For I = 3 To .Fields.Count - 1
                Monto = Format(varTotal(I), "#,##0.00")
                Printer.CurrentY = Linea
                Printer.CurrentX = Left + Pos - Printer.TextWidth(Monto)
                Printer.Print Monto
                Pos = Pos + 900
            Next
            Monto = Format(varTotal(2), "#,##0.00")
            Printer.CurrentY = Linea
            Printer.CurrentX = Left + Pos - Printer.TextWidth(Monto)
            Printer.Print Monto
            Printer.FontBold = False
            '
            Fin = Printer.CurrentY
            Printer.Line (150 + Left, INI)-(150 + Left, Fin)
            Printer.Line (550 + Left, INI)-(550 + Left, Fin)
            Pos = 3600 + Left
            For I = 3 To .Fields.Count + 1
                Pos = Left + 3600 + (900 * (I - 3))
                Printer.Line (Pos, INI)-(Pos, Fin)
            Next
            Printer.CurrentY = INI
            'Printer.Line (150 + Left, Printer.CurrentY)-(Pos, Printer.CurrentY), , BF
            Do
                Printer.Line (150 + Left, Printer.CurrentY)-(Pos, Printer.CurrentY)
                Printer.Print
            Loop Until Printer.CurrentY >= Fin
            Printer.Line (150 + Left, Printer.CurrentY)-(Pos, Printer.CurrentY)
            Return
        End With
    
    Else
        MsgBox "No existe información", vbInformation, App.ProductName
    End If
    rst.Close
    Set rst = Nothing
    End Sub

    Private Sub Grid_IE()
    'variable locales
    Dim rst As New ADODB.Recordset
    Dim FS As String, ancho As String
    Const m$ = "#,##0.00"
    Dim I%, j%, K%
    Dim varTotal() As Currency
    '
    Grid(4).Rows = 2
    Call rtnLimpiar_Grid(Grid(4))
    With rst
        .Open "ieconomico", cnnOLEDB & mcDatos, adOpenKeyset, adLockOptimistic, adCmdTable
        'aplica el filtro seleccionado
'        If optIECO(0) Then
'            .Filter = 0
'        ElseIf optIECO(1) Then
'            .Filter = "Fijo = True"
'        ElseIf optIECO(2) Then
'            .Filter = "Fijo = False"
'        End If
        If Not .EOF And Not .BOF Then
        
            ReDim varTotal(.Fields.Count)
            .MoveFirst
            'configura el encabezado del grid
            Grid(4).Cols = .Fields.Count - 1
            Grid(4).Rows = .RecordCount + 2
            FS = "^Código|<Descripción"
            ancho = "800|4000"
            For I = 3 To .Fields.Count - 1
              FS = FS + "|>" & Format(.Fields(I).Name, "mmm-yy")
              ancho = ancho + "|1000"
            Next
            FS = FS + "|>Total"
            ancho = ancho + "|1200"
            Grid(4).FormatString = FS
            Grid(4).Tag = ancho
            Call centra_titulo(Grid(4), True)
            I = 1
            Do
                
                For j = 0 To (.Fields.Count - 1)
                    If j = 2 Then
                        'Grid(4).Font.Bold = True
                        Grid(4).TextMatrix(I, .Fields.Count - 1) = Format(.Fields(2), m)
                    
                    Else
                        If j > 2 Then
                            K = j - 1
                        Else
                            K = j
                        End If
                        Grid(4).TextMatrix(I, K) = IIf(j >= 3, IIf(IsNull(.Fields(j)), "0,00", _
                        Format(.Fields(j), m)), IIf(IsNull(.Fields(j)), "", .Fields(j)))
                    End If
                    
                    If j > 1 Then varTotal(j) = IIf(IsNull(.Fields(j)), 0, .Fields(j)) + varTotal(j)
                    
                Next j
                .MoveNext: I = I + 1
                
            Loop Until .EOF
            .Close
            'On Error Resume Next
            With Grid(4)
            
                .TextMatrix(I, 1) = "TOTALES:"
                For j = 2 To .Cols - 1
                    .TextMatrix(I, j) = Format(varTotal(j + 1), m)
                Next
                .TextMatrix(I, rst.Fields.Count - 1) = Format(varTotal(2), m)
'                .Col = .Cols - 1
'                .Row = 1
'                .RowSel = .Rows - 1
'                .FillStyle = flexFillRepeat
'                .CellFontBold = True
'                .CellFontSize = 9
'                .Row = .Rows - 1
'                .Col = 0
'                .ColSel = .Cols - 1
'                .CellFontBold = True
'                .CellFontSize = 9
'                .FillStyle = flexFillSingle
            End With
            '
        End If
        '
    End With
    '
    End Sub

    
    '---------------------------------------------------------------------------------------------
    '
    '   Rutina: Export
    '
    '   Exporta el conetino del grid a una hoja de un libro excel
    '---------------------------------------------------------------------------------------------
    Private Sub Export()
    'variables locales
    'dim Libro as exce
    End Sub
    
    
    '---------------------------------------------------------------------------------------------
    '
    '   Function: Descrip
    '
    '   Entrada: N, contiene la descripcion del gasto tal como esta registrada en el catàlogo de
    '   gastos.
    '
    '   Esta función trata la variable recibida y devuelve una ajustada a la aplicacion (resumina)
    '---------------------------------------------------------------------------------------------
    Private Function Descrip(n As String) As String
    'variables locales
    Dim INI As Integer  'contiene el punto de inicio
    Dim K As Integer
    
    INI = InStr(n, " ")
    If INI = 0 Then
    
        Descrip = n
        
    Else
        If Left(n, INI - 1) <> "SUELDO" And Left(n, INI - 1) <> "APARTADO" Then
            Descrip = Left(n, INI - 1)
            K = 1
        End If
        n = Mid(n, INI + 1, Len(n))
        
        Do
            
            INI = InStr(n, " ")
        
            If INI = 0 Then
                If n <> "SOCIALES" Then
                    Descrip = Descrip & " " & n
                    n = ""
                End If
            Else
                
                If Left(n, INI - 1) = "SOCIALES" Then
                    n = ""
                Else
                    Descrip = Descrip & " " & Trim(Left(n, INI - 1))
                    n = Mid(n, INI + 1, Len(n))
                    K = K + 1
                End If
                
            End If
            
            
        Loop Until n = "" Or K = 2
        
    End If
    '
    End Function

    Private Sub Imprimir_ieco()
    'variables locales
    Dim strLogo As String
    Dim subT As String
    Dim rst As New ADODB.Recordset
    Dim Linea As Long
    Const LeftMargin& = 1500
    '
    MousePointer = vbHourglass
    
    For E = 0 To 2
        '
        'Printer.PaperBin = 15
        Printer.PrintQuality = vbPRPQHigh
        If txtCopy(0) = "" Then txtCopy(0) = 0
        If E > 1 Then Printer.Copies = IIf(txtCopy(0) = 0, 1, txtCopy(0))
        If E = txtCopy(0) Then Exit For
        '
        Printer.Font.Name = "TAHOMA"
        Printer.FontSize = 10
        Printer.FontBold = True
        '
        If E = 0 Then   'REPORTE DE LA ADMINISTRADORA
            Printer.CurrentX = Printer.ScaleWidth - Printer.TextWidth("ADMINISTRADORA")
            Printer.Print "ADMINISTRADORA"
        End If
        'imprime el encabezado
        If opt(0) Then  'logo inversiones
            Printer.PaintPicture pic(0), ScaleLeft, ScaleTop
            Printer.CurrentY = pic(0).Height
        Else    'logo sac
            Printer.PaintPicture pic(1), ScaleLeft, ScaleTop
            Printer.CurrentY = pic(1).Height
        End If
        Printer.Print
        subT = "INFORME ECONOMICO"
        Printer.CurrentX = (Printer.ScaleWidth - Printer.TextWidth(subT)) / 2

        Printer.Print subT
        Printer.CurrentX = (Printer.ScaleWidth - Printer.TextWidth(gcNomInm)) / 2
        Printer.Print gcNomInm
        
        rst.Open "SELECT Max(Periodo) FROM Factura WHERE Fact Not Like 'CH%'", cnnOLEDB & gcPath _
        & gcUbica & "\inm.mdb", adOpenKeyset, adLockOptimistic, adCmdText
        
        If Not rst.EOF And Not rst.BOF Then
            subT = "A " & UCase(Format(rst.Fields(0), "MMMM/YYYY"))
            Printer.CurrentX = (Printer.ScaleWidth - Printer.TextWidth(subT)) / 2
            Printer.Print subT
        End If
        '
        rst.Close
        Printer.Print
        Printer.FontSize = 8
        Printer.FontBold = False
    
        '
        'imprime la información del grid
        With Grid(3)
        
            For I = 1 To .Rows - 1
                Linea = Printer.CurrentY
                Printer.CurrentX = LeftMargin
                Printer.Print .TextMatrix(I, 0)
                'columna 2
                If .TextMatrix(I, 1) <> "" Then
                    subT = .TextMatrix(I, 1)
                    Printer.CurrentY = Linea
                    Printer.CurrentX = 8000 - Printer.TextWidth(subT)
                    Printer.Print subT
                End If
                'columna 3
                If .TextMatrix(I, 2) <> "" Then
                    subT = .TextMatrix(I, 2)
                    Printer.CurrentY = Linea
                    Printer.CurrentX = 10000 - Printer.TextWidth(subT)
                    Printer.Print subT
                End If
                
            Next
        End With
        'imprime los graficos
        'grafico deuda
        Linea = 8000
        Printer.CurrentX = LeftMargin
        Printer.CurrentY = Linea
        Printer.Print "Promedio de Facturacion Bs. " & txtProm
        'Printer.CurrentX = LeftMargin
        'Printer.Print "Auto Gestión Bs. " & Format(CCur(txtProm) * 2, "#,##0.00")
        If E = 0 Then
            Printer.CurrentX = LeftMargin
            'subT
            Printer.Print "Saldo en Banco Bs. " & Format(InputBox("Introduzca el saldo en banco" _
            & ":", App.ProductName, "0,00"), "#,##0.00")
        End If
        Graf.EditCopy
        Printer.PaintPicture Clipboard.GetData, 10000 - 4500, Linea, 4500, 2000
        '
        Linea = 7000
        graf1.EditCopy
        Printer.PaintPicture Clipboard.GetData, Printer.ScaleLeft, Linea + _
        2500, Printer.ScaleWidth, CLng((Printer.ScaleWidth / graf1.Width) * graf1.Height)
        'Printer.Print
        '
        Set rst = Nothing
        If E = 1 Then   'REPORTE DE LA JUNTA"
            subT = "Recibido por la Junta:________________________________"
            Printer.CurrentX = (Printer.ScaleWidth - Printer.TextWidth(subT)) / 2
            Printer.CurrentY = Printer.ScaleHeight - 800
            Printer.Print subT
        End If
        '
        'imprime la fecha de impresión
        Printer.FontName = "Arial Narrow"
        Printer.FontSize = 8
        Printer.CurrentY = Printer.ScaleHeight - 400
        Printer.Print "Fecha de Impresión: " & Format(Date, "dd/mm/yyyy")
        '
        'fecha de la asamblea
        subT = "Entragado en asamblea de fecha: " & dtp(1)
        Printer.CurrentX = Printer.ScaleWidth - Printer.TextWidth(subT)
        Printer.CurrentY = Printer.ScaleHeight - 400
        Printer.Print subT
        '
        'IMPRIME EL PIE DE PAGINA
        Printer.FontSize = 6
        subT = "Av. Caurimare Centro Caroní Modulo A Piso 1 Ofc. 14. Teléfonos (0212).751.99.92/ 37" _
        & ".64 / 67-961, Fax:(0212).753.81.11 email: administracion@administradorasac.com"
        Printer.CurrentX = (Printer.ScaleWidth - Printer.TextWidth(subT)) / 2
        Printer.CurrentY = Printer.ScaleHeight - 200
        Printer.Print subT
        '
        Printer.EndDoc
        '
    Next
    '
    MsgBox "Reporte impreso con éxito", vbInformation, App.ProductName
    MousePointer = vbDefault
    '
    End Sub

Private Sub mostrarInformeEconomico()
Dim cuentasDeFondo() As String
Dim cuentasCuotasEspeciales() As String
Dim numeroDeFacturaciones As Integer
Dim strCodigos As String, mostrarIntereses As Boolean
Dim Desde As Date, Hasta As Date, fechaUltimaAsamblea As Date
Dim deudaCondominio As Double
If Not (IsDate("01/" & MskFecha(0)) Or IsDate(MskFecha(1))) Then
    MsgBox "Complete la la información del paso Nº 7.", vbCritical, App.ProductName
    Exit Sub
End If

cuentasDeFondo = Split(getCuentasDe("Fondo"), ",")
cuentasCuotasEspeciales = Split(getCuentasDe("CuotasEspeciales"), ",")
numeroDeFacturaciones = txtIEco
Desde = "01/" & MskFecha(0)
Hasta = "01/" & MskFecha(1)
fechaUltimaAsamblea = dtp(0).Value
mostrarIntereses = chk.Value = vbChecked
deudaCondominio = FrmAdmin.objRst("deudaAct")

Call ModGeneral.imprimirInformeEconomico(cuentasDeFondo, cuentasCuotasEspeciales, _
mostrarIntereses, numeroDeFacturaciones, deudaCondominio, Desde, Hasta, fechaUltimaAsamblea, _
dtp(1).Value, crPantalla, txtCopy(2), txtCopy(3), txtCopy(4), txtCopy(5))

End Sub

Function getCuentasDe(tipoCuenta As String) As String
Dim Index As Integer
Dim strIN As String
Index = IIf(tipoCuenta = "Fondo", 0, 1)

Grid(Index).Col = 1
    
For I = 1 To Grid(Index).Rows - 1
    Grid(Index).Row = I
    If Grid(Index).CellPicture = img(1) Then
        strIN = strIN & IIf(strIN = "", "", ",") & Grid(0).TextMatrix(I, 0)
    End If
Next I
getCuentasDe = strIN

End Function
