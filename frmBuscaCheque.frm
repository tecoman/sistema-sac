VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmBuscaCheque 
   Caption         =   "Busquesa Avenzada::.."
   ClientHeight    =   7410
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10215
   Icon            =   "frmBuscaCheque.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7410
   ScaleWidth      =   10215
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmd 
      Caption         =   "&Filtrar"
      Height          =   915
      Index           =   1
      Left            =   8550
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   4995
      Width           =   1335
   End
   Begin VB.CommandButton cmd 
      Cancel          =   -1  'True
      Caption         =   "&Cerrar"
      Height          =   915
      Index           =   0
      Left            =   8550
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   6120
      Width           =   1335
   End
   Begin VB.Frame Frame1 
      Caption         =   "Busqueda y Filtro:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2220
      Left            =   375
      TabIndex        =   1
      Top             =   4845
      Width           =   8070
      Begin VB.TextBox txt 
         Alignment       =   2  'Center
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
         Index           =   5
         Left            =   4635
         TabIndex        =   22
         Top             =   1260
         Width           =   1110
      End
      Begin VB.TextBox txt 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   4
         Left            =   4620
         TabIndex        =   8
         Text            =   "0,00"
         Top             =   795
         Visible         =   0   'False
         Width           =   1140
      End
      Begin VB.TextBox txt 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   3
         Left            =   2295
         TabIndex        =   6
         Text            =   "0,00"
         Top             =   795
         Width           =   1155
      End
      Begin VB.ComboBox cmb 
         Height          =   315
         Index           =   1
         ItemData        =   "frmBuscaCheque.frx":030A
         Left            =   1455
         List            =   "frmBuscaCheque.frx":031A
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   810
         Width           =   825
      End
      Begin VB.TextBox txt 
         Alignment       =   2  'Center
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
         Index           =   2
         Left            =   6690
         Locked          =   -1  'True
         TabIndex        =   20
         Top             =   1680
         Width           =   1110
      End
      Begin VB.TextBox txt 
         Height          =   315
         Index           =   1
         Left            =   1455
         TabIndex        =   12
         Top             =   1680
         Width           =   1995
      End
      Begin VB.TextBox txt 
         Height          =   315
         Index           =   0
         Left            =   1455
         TabIndex        =   3
         Top             =   375
         Width           =   1995
      End
      Begin VB.ComboBox cmb 
         Height          =   315
         Index           =   0
         ItemData        =   "frmBuscaCheque.frx":032A
         Left            =   1455
         List            =   "frmBuscaCheque.frx":032C
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   1245
         Width           =   1995
      End
      Begin MSMask.MaskEdBox msk 
         Height          =   315
         Index           =   0
         Left            =   4635
         TabIndex        =   14
         Top             =   345
         Width           =   1110
         _ExtentX        =   1958
         _ExtentY        =   556
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   10
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox msk 
         Height          =   315
         Index           =   1
         Left            =   6690
         TabIndex        =   16
         Top             =   345
         Width           =   1110
         _ExtentX        =   1958
         _ExtentY        =   556
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   10
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label lbl 
         Caption         =   "Cod.Inm.:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   8
         Left            =   3795
         TabIndex        =   21
         Top             =   1260
         Width           =   720
      End
      Begin VB.Label lbl 
         Caption         =   " y "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   7
         Left            =   3795
         TabIndex        =   7
         Top             =   825
         Visible         =   0   'False
         Width           =   300
      End
      Begin VB.Label lbl 
         Caption         =   "Monto:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   6
         Left            =   300
         TabIndex        =   4
         Top             =   825
         Width           =   570
      End
      Begin VB.Label lbl 
         Caption         =   "Registros:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   5
         Left            =   5700
         TabIndex        =   19
         Top             =   1695
         Width           =   945
      End
      Begin VB.Label lbl 
         Caption         =   "Forma de Pago:"
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
         Index           =   4
         Left            =   300
         TabIndex        =   9
         Top             =   1185
         Width           =   1140
      End
      Begin VB.Label lbl 
         Caption         =   "Hasta:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   3
         Left            =   5955
         TabIndex        =   15
         Top             =   360
         Width           =   720
      End
      Begin VB.Label lbl 
         Caption         =   "Desde:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   2
         Left            =   3795
         TabIndex        =   13
         Top             =   390
         Width           =   720
      End
      Begin VB.Label lbl 
         Caption         =   "Banco:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   1
         Left            =   300
         TabIndex        =   11
         Top             =   1695
         Width           =   1020
      End
      Begin VB.Label lbl 
         Caption         =   "Nº Cheque:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   0
         Left            =   300
         TabIndex        =   2
         Top             =   390
         Width           =   1065
      End
   End
   Begin MSDataGridLib.DataGrid Grid 
      Bindings        =   "frmBuscaCheque.frx":032E
      Height          =   3885
      Left            =   390
      TabIndex        =   0
      Top             =   630
      Width           =   9480
      _ExtentX        =   16722
      _ExtentY        =   6853
      _Version        =   393216
      AllowUpdate     =   0   'False
      HeadLines       =   1
      RowHeight       =   15
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
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "LISTADOS DE CHEQUES RECIBIDOS"
      ColumnCount     =   8
      BeginProperty Column00 
         DataField       =   "CodInmueble"
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
      BeginProperty Column01 
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
      BeginProperty Column02 
         DataField       =   "FechaMov"
         Caption         =   "Fecha Mov."
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
         DataField       =   "Fpago"
         Caption         =   "Forma Pago"
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
         DataField       =   "Ndoc"
         Caption         =   "Nº Cheque"
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
         DataField       =   "Banco"
         Caption         =   "Banco"
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
         DataField       =   "FechaDoc"
         Caption         =   "Fecha Cheque"
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
      BeginProperty Column07 
         DataField       =   "Monto"
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
            ColumnWidth     =   645,165
         EndProperty
         BeginProperty Column01 
            Alignment       =   2
            ColumnWidth     =   720
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   1049,953
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   1170,142
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   1200,189
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   1709,858
         EndProperty
         BeginProperty Column06 
            Alignment       =   2
            ColumnWidth     =   1005,165
         EndProperty
         BeginProperty Column07 
            Alignment       =   1
            ColumnWidth     =   1500,095
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc adoCheque 
      Height          =   330
      Left            =   45
      Top             =   6165
      Visible         =   0   'False
      Width           =   3450
      _ExtentX        =   6085
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
      Caption         =   "Cheques Recibidos"
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
Attribute VB_Name = "frmBuscaCheque"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub adoCheque_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
On Error Resume Next
Txt(2) = adoCheque.Recordset.RecordCount
If Err.Number <> 0 Then Txt(2) = "##"
End Sub

Private Sub cmb_Change(index As Integer)
If index = 1 Then
    
End If
End Sub

Private Sub cmb_Click(index As Integer)
'variables locales

If index = 1 Then

    If cmb(1).Text = cmb(1).List(3) Then
        Txt(4).Visible = True
        Lbl(7).Visible = True
    Else
        Txt(4).Visible = False
        Lbl(7).Visible = False
    End If
    '
End If
'
End Sub

Private Sub Cmd_Click(index As Integer)
If index = 0 Then 'cerrar formulario
    Unload Me
    Set frmBuscaCheque = Nothing
Else
    Call Filtra_Tabla
End If
End Sub

Private Sub Form_Load()
'variables locales
With adoCheque
    
    .CursorLocation = adUseClient
    .ConnectionString = cnnConexion
    .CommandType = adCmdText
    .RecordSource = "SELECT TDFCheques.*, MovimientoCaja.AptoMovimientoCaja FROM MovimientoCaja" _
    & " INNER JOIN TDFCheques ON MovimientoCaja.IDRecibo = TDFCheques.IDRecibo;"
    .Refresh
    Set Grid.DataSource = adoCheque
    
End With
'
cmb(0).AddItem "", 0
cmb(0).AddItem "CHEQUE", 1
cmb(0).AddItem "DEPOSITO", 2
cmb(0).AddItem "EFECTIVO", 3
cmb(0).AddItem "TRANSFERENCIA", 4
'co
cmd(0).Picture = LoadResPicture("SALIR", vbResIcon)
'cmd(1).Picture = LoadResCustom("FILTRO")
'
End Sub


Private Sub Grid_HeadClick(ByVal ColIndex As Integer)
'variables locales
Dim orden As String, O As String

Static Dire As Long
'
O = IIf(Dire = 0, "ASC", "DESC")

Select Case ColIndex
    Case 0
        orden = "CodInmueble " & O & ", AptoMovimientoCaja"
    Case 1
        orden = "AptoMovimientoCaja " & O & ", CodInmueble"
    Case 2
        orden = "Fpago " & O & ",FechaMov,CodInmueble, AptoMovimientoCaja"
    Case 3
        orden = "Ndoc " & O & ""
    Case 4
        orden = "Banco " & O & ",FechaMov,CodInmueble,AptoMovimientoCaja"
    Case 5
        orden = "FechaDoc " & O & ",CodInmueble,AptoMovimientoCaja"
    Case 6
        orden = "Monto " & O & ",FechaMov,CodInmueble,AptoMovimientoCaja"
End Select
Dire = IIf(Dire = 0, 1, 0)
adoCheque.Recordset.Sort = orden
End Sub

Private Sub msk_KeyPress(index As Integer, KeyAscii As Integer)
Call Validacion(KeyAscii, "0123456789")
End Sub


Private Sub txt_KeyPress(index As Integer, KeyAscii As Integer)

If index = 0 Then Call Validacion(KeyAscii, "0123456789")

If index = 3 Or index = 4 Or index = 5 Then
    If KeyAscii = Asc(".") Then KeyAscii = Asc(",")
    Call Validacion(KeyAscii, "0123456789,")
End If

If KeyAscii = 13 Then SendKeys vbTab
End Sub


Private Sub Filtra_Tabla()
'variables locales
Dim cadena_filtro As String
If Txt(0) <> "" Then cadena_filtro = "Ndoc ='" & Txt(0) & "'"
If Txt(1) <> "" Then
    cadena_filtro = IIf(cadena_filtro = "", "", cadena_filtro & " AND ") & " Banco ='" & Txt(1) & "'"
End If
If Msk(0) <> "" Then
    Msk(0).PromptInclude = True
    If IsDate(Msk(0)) Then
        cadena_filtro = IIf(cadena_filtro = "", "", cadena_filtro & " AND ") & " FechaMov>=#" & Msk(0) & "#"
        Msk(0).PromptInclude = False
    Else
        MsgBox "Introdujo una fecha no válida", vbInformation, App.ProductName
        Msk(0).SelStart = 0
        Msk(0).SelLength = Len(Msk(0))
        Msk(0).SetFocus
        Exit Sub
    End If
End If
'
If Msk(1) <> "" Then
    Msk(1).PromptInclude = True
    If IsDate(Msk(1)) Then
        cadena_filtro = IIf(cadena_filtro = "", "", cadena_filtro & " AND ") & " FechaMov<=#" & Msk(1) & "#"
        Msk(1).PromptInclude = False
    Else
        MsgBox "Introdujo una fecha no válida", vbInformation, App.ProductName
        Msk(1).SelStart = 0
        Msk(1).SelLength = Len(Msk(1))
        Msk(1).SetFocus
        Exit Sub
    End If
End If
'
If cmb(0) <> "" Then
    cadena_filtro = IIf(cadena_filtro = "", "", cadena_filtro & " AND ") & " Fpago='" & cmb(0).Text & "'"
End If
If cmb(1) <> "" Then
    If Txt(3) = "" Then Txt(3) = "0,00"
    
    If CCur(Txt(3)) > 0 Then
        If Txt(4).Visible = True Then
            If Txt(4) = "" Then Txt(4) = "0,00"
            If CCur(Txt(4)) > 0 Then
                cadena_filtro = IIf(cadena_filtro = "", "", cadena_filtro & " AND ") & " Monto " & IIf(cmb(1) = "/", ">=", cmb(1)) & _
                Replace(CCur(Txt(3)), ",", ".") & " AND Monto" & IIf(cmb(1) = "/", "<=", cmb(1)) & Replace(CCur(Txt(4)), ",", ".")
            Else
                MsgBox "Debe ingresar un cantidad mayor que cero", vbCritical, App.ProductName
                Txt(4).SelStart = 0
                Txt(4).SelLength = Len(Txt(4))
                Txt(4).SetFocus
                Exit Sub
            End If
        Else
            cadena_filtro = IIf(cadena_filtro = "", "", cadena_filtro & " AND ") & " Monto " & IIf(cmb(1) = "/", ">=", cmb(1)) & _
            Replace(CCur(Txt(3)), ",", ".")
        End If
    End If

End If
If Txt(5) <> "" Then
    cadena_filtro = IIf(cadena_filtro = "", "", cadena_filtro & " AND ") & "CodInmueble='" & Txt(5) & "'"
End If
        
adoCheque.Recordset.Filter = IIf(cadena_filtro = "", 0, cadena_filtro)

End Sub

Private Sub txt_LostFocus(index As Integer)
If index = 3 Or index = 4 Then
    If IsNumeric(Txt(index)) Then Txt(index) = Format(Txt(3), "#,##0.00")
End If
End Sub
