VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form FrmAvisos 
   Caption         =   "Cartas y Telegrmas"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdAviso 
      Caption         =   "&Recuperar Datos"
      Height          =   855
      Index           =   2
      Left            =   4185
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   4650
      Width           =   1215
   End
   Begin VB.CommandButton cmdAviso 
      Caption         =   "&Salir"
      Height          =   855
      Index           =   1
      Left            =   6615
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   4650
      Width           =   1215
   End
   Begin VB.CommandButton cmdAviso 
      Caption         =   "&Guardar"
      Height          =   855
      Index           =   0
      Left            =   5430
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   4650
      Width           =   1215
   End
   Begin MSFlexGridLib.MSFlexGrid FlexAvisos 
      Height          =   2955
      Left            =   400
      TabIndex        =   0
      Top             =   615
      Width           =   6915
      _ExtentX        =   12197
      _ExtentY        =   5212
      _Version        =   393216
      Cols            =   6
      FixedCols       =   0
      BackColorFixed  =   -2147483635
      ForeColorFixed  =   -2147483639
      BackColorSel    =   -2147483643
      ForeColorSel    =   -2147483640
      BackColorBkg    =   12632256
      WordWrap        =   -1  'True
      FocusRect       =   0
      GridLines       =   3
      GridLinesFixed  =   3
      FormatString    =   "Apto. |Propietario |^R.V. |Deuda |Carta |Telegrama"
   End
   Begin VB.Image ImgAceptar 
      Enabled         =   0   'False
      Height          =   150
      Index           =   1
      Left            =   1320
      Top             =   195
      Visible         =   0   'False
      Width           =   165
   End
   Begin VB.Image ImgAceptar 
      Enabled         =   0   'False
      Height          =   150
      Index           =   0
      Left            =   1005
      Top             =   210
      Visible         =   0   'False
      Width           =   165
   End
End
Attribute VB_Name = "FrmAvisos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
    'Seac.Modulo Facturacion [Avisos de Cobro]----Muestra todos los documentos enviados a los-----
    'propietarios por concepto de su deuda. Permite la edición de estos registros para generar los
    'cargos respectivos 03/09/2002. Solo cambia Si---->No [no] No--X-->SI-------------------------
    Dim cnnInmueble As ADODB.Connection     'Conexión pública a nivel de módulo al inmueble
    Dim adoAvisos As ADODB.Recordset         'Conjunto de Registos Prop. con Avisos enviados
    Dim I%, j%
    '---------------------------------------------------------------------------------------------
    
    '04/09/2002-----------------------------------------------------------------------------------
    Private Sub cmdAviso_Click(Index As Integer)    '
    '---------------------------------------------------------------------------------------------
    '
    Dim strCarta$, strTelegrama$
    
    Select Case Index
        Case 0  'Boton Guardar
    '   -------------------------------
            
            With FlexAvisos
                For I = 1 To .Rows - 1
                    strCarta = "false"
                    strTelegrama = "False"
                    For j = 4 To 5
                        .Col = j
                        .Row = I
                        If j = 4 And .CellPicture = ImgAceptar(0) Then strCarta = "True"
                        If j = 5 And .CellPicture = ImgAceptar(0) Then strTelegrama = "True"
                    Next
                cnnInmueble.Execute "UPDATE Propietarios SET Carta=" & strCarta & ", Telegrama=" _
                    & strTelegrama & " WHERE Codigo='" & .TextMatrix(I, 0) & "'"
                Next
                cmdAviso(2).Enabled = False
                MsgBox "Registros Actualizados...", vbInformation, App.ProductName
            End With
        Case 1 'Boton Cerrar
    '   -------------------------------
            Unload Me
            
        Case 2  'Boton Recuperar
    '   -------------------------------
            Call RtnGrid(adoAvisos)
    End Select
    '
    End Sub

    '03/09/2002-----------------------------------------------------------------------------------
    Private Sub FlexAvisos_Click()  '
    '---------------------------------------------------------------------------------------------
    Dim strPropietarios$
    'Permite la midificación de la condición SI a No, pero no a la inversa
    With FlexAvisos
    '
        If .ColSel = 4 Or .ColSel = 5 Then
            strPropietario = .TextMatrix(.RowSel, 0)
            adoAvisos.MoveFirst
            adoAvisos.Find "Codigo ='" & strPropietario & "'"
            If Not adoAvisos.EOF Then
                If adoAvisos.Fields(.ColSel) = True Then
                    .Col = .ColSel: .Row = .RowSel
                    Set .CellPicture = IIf(.CellPicture = ImgAceptar(0), ImgAceptar(1), ImgAceptar(0))
                    .CellPictureAlignment = flexAlignCenterCenter
                End If
            End If
    '
        End If
    '
    End With
    '
    End Sub

    '---------------------------------------------------------------------------------------------
    Private Sub Form_Load()
    '---------------------------------------------------------------------------------------------
    '
    Dim strSQL$
    
    ImgAceptar(0).Picture = LoadResPicture("Checked", vbResBitmap)
    ImgAceptar(1).Picture = LoadResPicture("Unchecked", vbResBitmap)
    cmdAviso(1).Picture = LoadResPicture("Salir", vbResIcon)
    cmdAviso(0).Picture = LoadResPicture("Guardar", vbResIcon)
    
    With FlexAvisos
        .RowHeight(0) = 300
        Call centra_titulo(FlexAvisos)
    End With
    '---------------------------------------------------------------------------------------------
    Set cnnInmueble = New ADODB.Connection
    Set adoAvisos = New ADODB.Recordset
    '---------------------------------------------------------------------------------------------
    cnnInmueble.CursorLocation = adUseClient
    cnnInmueble.Open cnnOLEDB + mcDatos
    '---------------------------------------------------------------------------------------------
    strSQL = "SELECT Codigo, Nombre,Recibos,Deuda, Carta, telegrama FROM propietarios " _
        & "WHERE Carta=True or Telegrama=True ORDER BY codigo;"
    
    adoAvisos.Open strSQL, cnnInmueble, adOpenStatic, adLockReadOnly
    If Not adoAvisos.EOF Then   'Si se recuperan registros
    '
        Call RtnGrid(adoAvisos)
        cmdAviso(2).Enabled = True
    '
    End If
    '
    End Sub

    '03/09/2002----------Configura la presentación en pantalla de la información dependiendo del
    Private Sub Form_Resize()   'estado visual del formulario
    '---------------------------------------------------------------------------------------------
    '
    If Me.WindowState <> vbMinimized Then
    '
        With FlexAvisos
    '
            .Top = 700
            .Left = 400
            .Width = (Me.ScaleWidth - (.Left * 2))
            .Height = (Me.ScaleHeight - 2000)
    '       Configura ancho de las columas
            .ColWidth(0) = 1000
            .ColWidth(2) = 700
            .ColWidth(3) = 1200
            .ColWidth(4) = 1000
            .ColWidth(5) = 1000
            .ColWidth(1) = _
                .Width - (.ColWidth(0) + .ColWidth(2) + .ColWidth(3) + .ColWidth(4) _
                + .ColWidth(5)) - 100
        End With
        For I = 0 To 2
            cmdAviso(I).Top = Me.ScaleHeight - 1000
        Next
        cmdAviso(2).Left = FlexAvisos.Left + FlexAvisos.Width - (cmdAviso(0).Width * 3)
        cmdAviso(0).Left = cmdAviso(2).Left + cmdAviso(2).Width
        cmdAviso(1).Left = cmdAviso(0).Left + cmdAviso(2).Width
    '
    End If
    '
    End Sub

    '---------------------------------------------------------------------------------------------
    Private Sub Form_Unload(Cancel As Integer)  'Descarga de memoria el formulario y la conexión
    '---------------------------------------------------------------------------------------------
    adoAvisos.Close
    Set adoAvisos = Nothing
    cnnInmueble.Close
    Set cnnInmueble = Nothing
    
    End Sub

    '04/09/2002-----------------------------------------------------------------------------------
    Private Sub RtnGrid(Ado As ADODB.Recordset) 'Muestra en el Grid los registros del ADODB.Recordset
    '---------------------------------------------------------------------------------------------
    If Not (Ado.EOF And Ado.BOF) Then
        Ado.MoveFirst
        I = 0
        MousePointer = vbHourglass
        Do Until Ado.EOF
        I = I + 1
        FlexAvisos.Rows = Ado.RecordCount + 1
        '
            For j = 0 To 5
                If j = 0 Or j <= 3 Then
                    FlexAvisos.TextMatrix(I, j) = _
                        IIf(IsNull(Ado.Fields(j)), "", Ado.Fields(j))
                    If j = 3 Then FlexAvisos.TextMatrix(I, j) = _
                        Format(FlexAvisos.TextMatrix(I, j), "#,##0.00")
                Else
                    FlexAvisos.Col = j
                    FlexAvisos.Row = I
                    If Ado.Fields(j) = True Then  'Marca los propietarios con
                        Set FlexAvisos.CellPicture = ImgAceptar(0)
                    Else
                        Set FlexAvisos.CellPicture = ImgAceptar(1)
                    End If
                    FlexAvisos.CellPictureAlignment = flexAlignCenterCenter
                End If
                
            Next
            Ado.MoveNext
        '
        Loop
        MousePointer = vbDefault
    End If
    '
    End Sub
