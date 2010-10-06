VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form FrmAvisoLegal 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Aviso Legal"
   ClientHeight    =   4080
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6630
   ClipControls    =   0   'False
   FillStyle       =   0  'Solid
   Icon            =   "FrmAvisoLegal.frx":0000
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4080
   ScaleWidth      =   6630
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FraAvisoCobro 
      Height          =   1245
      Index           =   2
      Left            =   135
      TabIndex        =   8
      Top             =   2625
      Width           =   6270
      Begin VB.CommandButton cmdAviso 
         Caption         =   "Im&primir"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   800
         Index           =   0
         Left            =   3855
         MouseIcon       =   "FrmAvisoLegal.frx":5C12
         MousePointer    =   99  'Custom
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   225
         Width           =   1095
      End
      Begin VB.CommandButton cmdAviso 
         Caption         =   "&Salir"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   800
         Index           =   1
         Left            =   5025
         MouseIcon       =   "FrmAvisoLegal.frx":5F1C
         MousePointer    =   99  'Custom
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   225
         Width           =   1095
      End
   End
   Begin VB.Frame FraAvisoCobro 
      Caption         =   "Selección Inmueble"
      Height          =   2310
      Index           =   0
      Left            =   200
      TabIndex        =   0
      Top             =   120
      Width           =   6180
      Begin VB.OptionButton optImp 
         Caption         =   "Ventana"
         Height          =   270
         Index           =   0
         Left            =   2250
         TabIndex        =   13
         Top             =   1770
         Value           =   -1  'True
         Width           =   1215
      End
      Begin VB.OptionButton optImp 
         Caption         =   "Impresora"
         Height          =   270
         Index           =   1
         Left            =   900
         TabIndex        =   12
         Top             =   1770
         Width           =   1215
      End
      Begin VB.OptionButton optImp 
         Caption         =   "Correo electrónico"
         Height          =   270
         Index           =   2
         Left            =   3525
         TabIndex        =   11
         Top             =   1770
         Width           =   1890
      End
      Begin MSDataListLib.DataCombo DtcAviso 
         Height          =   315
         Index           =   0
         Left            =   1170
         TabIndex        =   2
         Top             =   480
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   556
         _Version        =   393216
         ListField       =   "CodInm"
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo DtcAviso 
         Height          =   315
         Index           =   1
         Left            =   2265
         TabIndex        =   3
         Top             =   480
         Width           =   3765
         _ExtentX        =   6641
         _ExtentY        =   556
         _Version        =   393216
         ListField       =   "Nombre"
         Text            =   ""
      End
      Begin VB.Label Label1 
         Caption         =   "Salida:"
         Height          =   285
         Index           =   1
         Left            =   5160
         TabIndex        =   15
         Top             =   1080
         Width           =   900
      End
      Begin VB.Label Label1 
         Caption         =   "Salida:"
         Height          =   285
         Index           =   7
         Left            =   165
         TabIndex        =   14
         Top             =   1770
         Width           =   900
      End
      Begin VB.Label Label1 
         Height          =   285
         Index           =   5
         Left            =   1170
         TabIndex        =   10
         Top             =   1365
         Width           =   4155
      End
      Begin VB.Label Label1 
         Caption         =   "Ubicación:"
         Height          =   285
         Index           =   4
         Left            =   195
         TabIndex        =   9
         Top             =   1365
         Width           =   900
      End
      Begin VB.Label Label1 
         Height          =   285
         Index           =   3
         Left            =   1170
         TabIndex        =   5
         Top             =   945
         Width           =   4155
      End
      Begin VB.Label Label1 
         Caption         =   "Impresora:"
         Height          =   285
         Index           =   2
         Left            =   195
         TabIndex        =   4
         Top             =   945
         Width           =   900
      End
      Begin VB.Label Label1 
         Caption         =   "&Inmueble:"
         Height          =   285
         Index           =   0
         Left            =   200
         TabIndex        =   1
         Top             =   495
         Width           =   900
      End
   End
End
Attribute VB_Name = "FrmAvisoLegal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim objRst As ADODB.Recordset
Dim intMora As Integer
    
Private Sub cmdAviso_Click(Index As Integer)

Select Case Index
    Case 0  'PRESIONO IMPRIMIR
        Call imprimir_aviso_legal
        
    Case 1  'PRESIONO SALIR
            Unload Me
            Set FrmAvisoLegal = Nothing
End Select
'
End Sub

Private Sub DtcAviso_Click(Index As Integer, Area As Integer)
    If Area = 2 Then    'SI SELECCIONA UN ELEMENTO DE LA LISTA
        Select Case Index
    '
            Case 0  'BUSCA POR CODIGO DE INMUEBLE
                Call RtnBusqueda("CodInm = '" & DtcAviso(0).Text & "'")
            
            Case 1  'BUSCA POR NOMBRE DEL INMUEBLE
                Call RtnBusqueda("Nombre like '%" & DtcAviso(1).Text & "%'")
    '
        End Select
    '
    End If
End Sub

    '20/08/2002-----------------------------------------------------------------------------------
    Private Sub DtcAviso_KeyPress(Index As Integer, KeyAscii As Integer)
    '---------------------------------------------------------------------------------------------
    'variables locales
    KeyAscii = Asc(UCase(Chr(KeyAscii)))   'CONVIERTE A MAYUSCULAR
    If KeyAscii = 13 Then   'SI PRESIONA ENTER
    '
        Select Case Index
            'BUSQUEDA POR CODIGO DE INMUEBLE
            Case 0: Call RtnBusqueda("CodInm = '" & DtcAviso(0).Text & "'")
            'BUSQUEDA POR NOMBRE DEL INMUEBLE
            Case 1: Call RtnBusqueda("Nombre like '%" & DtcAviso(1).Text & "%'")
    '
        End Select
    '
    End If
    '
    End Sub

Private Sub Form_Load()
    cmdAviso(0).Picture = LoadResPicture("Print", vbResIcon)
    cmdAviso(1).Picture = LoadResPicture("Salir", vbResIcon)
    'ESTABLECE LA PROPIEDAD CAPTION DE LAS ETIQUETAS
    Label1(3).Caption = Printer.DeviceName 'IMPRESORA
    Label1(5) = Printer.Port    'UBICACION IMPRESORA
    Set objRst = New ADODB.Recordset
    objRst.Open "SELECT * FROM Inmueble ORDER BY CodInm", cnnConexion, adOpenStatic, _
    adLockReadOnly
    Set DtcAviso(0).RowSource = objRst  'ESTABLECE EL ORIGEN DE LA LISTA
    Set objRst = New ADODB.Recordset
    objRst.Open "SELECT * FROM Inmueble ORDER BY Nombre", cnnConexion, adOpenStatic, _
    adLockReadOnly
    Set DtcAviso(1).RowSource = objRst
End Sub

'20/08/2002-------Rutina que busca información sobre el inmueble seleccionado----------------
Private Sub RtnBusqueda(StrExpresion As String)
'--------------------------------------------------------------------------------------------
'variables locales
'
With objRst
'
    .MoveFirst
    .Find StrExpresion
    If .EOF Then
        MsgBox "Inmueble No Registrado", vbCritical, App.ProductName
        Exit Sub
    End If
    For I = 0 To 1
        DtcAviso(I) = .Fields(I)
    Next
    StrRutaInmueble = .Fields("Ubica")      'Carpeta del Inmueble
    intMM = .Fields("MesesMora")            'Constante Meses de Mora
    intMora = .Fields("HonoMorosidad")      'Procentaje de Honorarios
    '
End With
'
End Sub

Private Sub imprimir_aviso_legal()
Dim rpReporte As ctlReport
Dim rpGeneral As ctlReport
Dim gsTEMPDIR As String, strArchivo As String
Dim lchar As Long
'
Set rpReporte = New ctlReport
With rpReporte
'
    .Reporte = gcReport + "ges_cob_legal.rpt"
    .OrigenDatos(0) = gcPath + "\" + Me.DtcAviso(0) + "\inm.mdb"
    .OrigenDatos(1) = gcPath + "\sac.mdb"
    .FormuladeSeleccion = "{Inmueble.CodInm}='" & Me.DtcAviso(0) & "' and {Propietarios.Recibos} > 5"
    .TituloVentana = Me.DtcAviso(0) & " - Notificación Legal"
    .Salida = Salida
    If Salida = crEmail Then
        Dim rst As ADODB.Recordset
        Dim cnn As ADODB.Connection
        Dim sql As String
        Dim archivo As String
        
        archivo = gcPath & "/legal.txt"
        
        Set rst = New ADODB.Recordset
        Set cnn = New ADODB.Connection
        cnn.Open cnnOLEDB & gcPath & "\" & Me.DtcAviso(0) & "\inm.mdb"
        
        sql = "SELECT * FROM Propietarios WHERE email <> '' AND Demanda = False and recibos > 5"
        rst.Open sql, cnn, adOpenStatic, adLockReadOnly, adCmdText
        If Not (rst.EOF And rst.BOF) Then
            Do
                Err.Clear
                .FormuladeSeleccion = "{Inmueble.CodInm}='" & Me.DtcAviso(0) & "' and {Propietarios.Recibos} > 5 " & _
                    "AND {Propietarios.Codigo} = '" & rst("codigo") & "'"
                .Salida = crArchivoDisco
                'si la salida es enviar email, se guarda un reporte en el temporal de windows
                gsTEMPDIR = String$(255, 0)
                lchar = GetTempPath(255, gsTEMPDIR)
                gsTEMPDIR = Left(gsTEMPDIR, lchar)
                .Salida = crArchivoDisco
                .FormatoSalida = crEFTPortableDocFormat
                strArchivo = gsTEMPDIR & "NL-" & Me.DtcAviso(0) & rst("codigo") & ".pdf"
                .ArchivoSalida = strArchivo
                .Imprimir
                If Err <> 0 Then MsgBox Err.Description, vbCritical, Err
                'ahora enviamos el email
                If enviar_email(rst("email"), "legal@administradorasac.com", _
                "Notificación Legal", False, _
                Mensaje(archivo, rst("nombre"), rst("codigo"), _
                Trim(DtcAviso(1)), rst("recibos"), Format(rst("deuda"), "#,##0.00"), _
                Format(intMora / 100, "#%")), strArchivo) Then
                    '   email enviado con éxito
                Else
                    '   error al enviar el email
                End If
                rst.MoveNext
            Loop Until rst.EOF
        End If
    Else
        .Imprimir
        If Err <> 0 Then MsgBox Err.Description, vbCritical, Err
    End If
    
End With
'
Set rpGeneral = New ctlReport
With rpGeneral
    If Respuesta("Desea el reporte general notificaciones?") Then
        .Reporte = gcReport + "ges_cob_legal_total.rpt"
        .OrigenDatos(0) = gcPath + "\" + Me.DtcAviso(0) + "\inm.mdb"
        .FormuladeSeleccion = "{Propietarios.Recibos} > 5"
        .Salida = crPantalla
        .TituloVentana = Me.DtcAviso(0) & " - Reporte General Notificación Legal"
        .Imprimir
        If Err <> 0 Then MsgBox Err.Description, vbCritical, Err
'
    End If
End With
'
End Sub

'---------------------------------------------------------------------------------------------
'   Funcion:    Salida
'
'   Devuelve una constante de Destino
'---------------------------------------------------------------------------------------------
Private Function Salida() As crSalida
'
Select Case True
    Case optImp(0): Salida = crPantalla
    Case optImp(1): Salida = crImpresora
    Case optImp(2): Salida = crEmail
End Select
'
End Function

Private Function Mensaje(archivo$, ParamArray datos()) As String

    
    'variables locales
    Dim N As Long
    Dim Dato As Byte
    Dim O As Integer
    N = FreeFile
    Open archivo For Binary As #N
        Do While Not EOF(N)
            Get N, , Dato
            If Dato = 63 Then
                Mensaje = Mensaje & Space(1) & datos(I)
                I = I + 1
            Else
                Mensaje = Mensaje & Chr(Dato)
            End If
        Loop
        
    Close #N

End Function
