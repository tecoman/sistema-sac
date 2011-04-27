VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form FrmSelCon 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Seleccionar Condominio"
   ClientHeight    =   6555
   ClientLeft      =   3675
   ClientTop       =   3165
   ClientWidth     =   4980
   Icon            =   "FrmSelCon.frx":0000
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6555
   ScaleWidth      =   4980
   ShowInTaskbar   =   0   'False
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   4815
      Left            =   30
      TabIndex        =   0
      Top             =   240
      Width           =   4935
      _ExtentX        =   8705
      _ExtentY        =   8493
      _Version        =   393216
      AllowUpdate     =   -1  'True
      BackColor       =   16777215
      BorderStyle     =   0
      HeadLines       =   1
      RowHeight       =   15
      RowDividerStyle =   4
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "C O N D O M I N I O S"
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   "CodInm"
         Caption         =   "Código"
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
      BeginProperty Column01 
         DataField       =   "Nombre"
         Caption         =   "Nombre del Inmueble"
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
         ScrollBars      =   2
         AllowRowSizing  =   0   'False
         AllowSizing     =   0   'False
         BeginProperty Column00 
            Alignment       =   2
            Locked          =   -1  'True
            ColumnWidth     =   615,118
         EndProperty
         BeginProperty Column01 
            Locked          =   -1  'True
            ColumnWidth     =   3764,977
         EndProperty
      EndProperty
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
      Height          =   1140
      Left            =   1590
      TabIndex        =   7
      Top             =   5175
      Width           =   3375
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
         Height          =   285
         Left            =   1530
         Locked          =   -1  'True
         TabIndex        =   9
         Top             =   735
         Width           =   690
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   255
         Left            =   135
         Locked          =   -1  'True
         TabIndex        =   8
         Text            =   "Total Registros"
         Top             =   765
         Width           =   1305
      End
      Begin VB.TextBox TxtBus 
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   780
         TabIndex        =   2
         Top             =   300
         Width           =   2490
      End
      Begin VB.CommandButton BotBusca 
         Appearance      =   0  'Flat
         Height          =   400
         Left            =   2520
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Buscar"
         Top             =   650
         Width           =   615
      End
      Begin VB.Label Label2 
         Caption         =   "B&uscar:"
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
         TabIndex        =   1
         Top             =   315
         Width           =   630
      End
   End
   Begin VB.Frame FrmBusca 
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
      Height          =   1155
      Left            =   45
      TabIndex        =   6
      Top             =   5160
      Width           =   1530
      Begin VB.OptionButton OptBusca 
         Caption         =   "Por &Nombre"
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
         Left            =   105
         TabIndex        =   4
         Top             =   795
         Width           =   1335
      End
      Begin VB.OptionButton OptBusca 
         Caption         =   "Por &Código"
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
         Left            =   105
         TabIndex        =   3
         Top             =   405
         Value           =   -1  'True
         Width           =   1335
      End
   End
End
Attribute VB_Name = "FrmSelCon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

    Private Sub BotBusca_Click(): Call RtnBuscaInmueble
    End Sub
    
    'DOBLE CLICK DATAGRID
    Private Sub DataGrid1_DblClick(): If Not FrmAdmin.objRst.EOF And Not FrmAdmin.objRst.BOF Then Call RtnSelCon
    End Sub

    Private Sub DataGrid1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then Call DataGrid1_DblClick
    End Sub

    Private Sub Form_Activate(): TxtBus.SetFocus
    End Sub

    Private Sub Form_Load()
    '
    BotBusca.Picture = LoadResPicture("Buscar", vbResIcon)
    Set DataGrid1.DataSource = FrmAdmin.objRst
    TxtTot = FrmAdmin.objRst.RecordCount
    FrmSelCon.Top = 700
    FrmSelCon.Left = 3000
    CenterForm Me
    '
    End Sub

    Private Sub OptBusca_Click(Index As Integer)
    TxtBus.SetFocus
    End Sub

    
    Private Sub TxtBus_KeyPress(KeyAscii As Integer)
    '
    If KeyAscii = 27 Then
        Call RtnSelCon
        Exit Sub
    End If
    '
    
    KeyAscii = Asc(UCase(Chr(KeyAscii))) 'Convierte en Mayuscula
    '
    If OptBusca(0).Value = True Then
        Call Validacion(KeyAscii, "1234567890")
    Else
        Call Validacion(KeyAscii, " ABCDEFGHIJKLMNÑOPQRSTUWVXYZ-/.")
    End If
    '
    If KeyAscii = 13 Then 'Permite Avanzar de campo con Enter
         
        If TxtBus.Text = "" Then
            MsgBox "Debe Ingresar la Expresión", vbInformation, App.ProductName
            TxtBus.SetFocus
        Else
            Call RtnBuscaInmueble
        End If
            
    End If
    '
    End Sub


    '---------------------------------------------------------------------------------------------
    '   Rutina:     RtnSelCon
    '
    '   Asigna valores a las variables globales s/inmueble seleccionado
    '---------------------------------------------------------------------------------------------
    Sub RtnSelCon()
    '
    With FrmAdmin.objRst
    
        If .State = 0 Then .Open
        If .RecordCount <= 0 Then
            MsgBox "No existen inmueble registrado...", vbInformation, App.ProductName
            Exit Sub
        End If
        If .EOF Or .BOF Then .MoveFirst
        Call rtnBitacora("Selección Inmueble " & !CodInm)
        gcCodInm = !CodInm 'Codigo del inmueble
        gcNomInm = !Nombre 'Nombre del inmueble
        gcUbica = !Ubica   'Carpeta inf. adicional
        gcCodFondo = IIf(IsNull(!CodFondo), 0, !CodFondo) 'Cod. fondo de reserva
        gnPorcFondo = !PorcFondo   '% por fondo de reserva
        gnPorIntMora = !HonoMorosidad   '
        gnMesesMora = !MesesMora    'meses para cobrar honorarios de abogado
        mcDatos = gcPath + gcUbica + "Inm.mdb"  'carpeta inf. adic. inmueble seleccionado
        '
        If !Caja = sysCodCaja Then
            gnCta = CUENTA_POTE
        Else
            gnCta = CUENTA_INMUEBLE
        End If
        '
    End With
    FrmAdmin.Caption = Trim(gcCodInm) + "  " + Trim(gcNomInm)
    TxtBus.SetFocus
    FrmSelCon.Hide
    TxtBus.Text = ""
    '
    End Sub

    Private Sub RtnBuscaInmueble()
    With FrmAdmin.objRst
        .MoveFirst
        If OptBusca(0).Value = True Then    'Si la busqueda es por codigo
            .Find "CodInm = '" & Left(TxtBus.Text, 4) & "'"
            If Not .EOF Then
                Call RtnSelCon
            Else
                With TxtBus
                    MsgBox "Verifique el Codigo introducido" _
                    & vbCrLf & "Propietario no Registrado '" & .Text & "'", vbCritical, App.ProductName
                    .SetFocus
                    .SelStart = 0
                    .SelLength = Len(.Text)
                End With
            End If
            Exit Sub
            
        Else
            
            .Find "Nombre Like '*" & Trim(TxtBus.Text) & "*'"
            
        End If
        
        If .EOF Then
            MsgBox "No se encontró el condominio  " & TxtBus, vbCritical + vbOKOnly, _
            "Seleccion de Inmueble"
            TxtBus.SetFocus
        Else
            Call RtnSelCon
        End If
        End With
    End Sub
