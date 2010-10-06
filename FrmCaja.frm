VERSION 5.00
Begin VB.Form FrmCaja 
   AutoRedraw      =   -1  'True
   Caption         =   "Form1"
   ClientHeight    =   3045
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4350
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "FrmCaja.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3045
   ScaleWidth      =   4350
   Begin VB.CommandButton CmdTaquilla 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      Height          =   855
      Index           =   1
      Left            =   2295
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1875
      Width           =   1215
   End
   Begin VB.CommandButton CmdTaquilla 
      Caption         =   "&Aceptar"
      Height          =   855
      Index           =   0
      Left            =   840
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1875
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   1440
      Left            =   165
      TabIndex        =   4
      Top             =   870
      Width           =   4020
      Begin VB.TextBox TxtMonto 
         Height          =   315
         Left            =   1500
         TabIndex        =   5
         Top             =   765
         Visible         =   0   'False
         Width           =   1320
      End
      Begin VB.Label Label1 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Index           =   3
         Left            =   1500
         TabIndex        =   8
         Top             =   353
         Width           =   2295
      End
      Begin VB.Label Label1 
         Caption         =   "Usuario:"
         Height          =   210
         Index           =   2
         Left            =   135
         TabIndex        =   7
         Top             =   405
         Width           =   1200
      End
      Begin VB.Label Label1 
         Caption         =   "Monto de Apertura"
         Height          =   450
         Index           =   4
         Left            =   135
         TabIndex        =   6
         Top             =   690
         Visible         =   0   'False
         Width           =   1200
      End
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Height          =   735
      Index           =   5
      Left            =   855
      TabIndex        =   9
      Top             =   2220
      Width           =   270
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      Index           =   1
      X1              =   4015
      X2              =   315
      Y1              =   465
      Y2              =   465
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderWidth     =   2
      Index           =   0
      X1              =   4000
      X2              =   330
      Y1              =   465
      Y2              =   465
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Label1"
      Height          =   195
      Index           =   1
      Left            =   300
      TabIndex        =   1
      Top             =   600
      Width           =   480
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000002&
      Height          =   300
      Index           =   0
      Left            =   315
      TabIndex        =   0
      Top             =   105
      Width           =   825
   End
End
Attribute VB_Name = "FrmCaja"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
    Public Titulo As String
    Public boton As String
    '
    'Rev.11/09/2002-------------------------------------------------------------------------------
    Private Sub CmdTaquilla_Click(Index As Integer) '
    '---------------------------------------------------------------------------------------------
    '
    Dim n&
    Select Case Index
    '
        Case 0  'ACEPTAR
    '   -----------------------------------------
            Select Case Trim(Me.Caption)
                
                Case "Abrir Caja"
    '           -------------
    '           Abre la taquilla y actualiza sus datos
                cnnConexion.Execute "UPDATE Taquillas SET Estado = True, Fecha= Date(), Hora=Da" _
                & "teAdd('n',-5,Time()) ,Usuario ='" & gcUsuario & "' WHERE IDTaquilla =" & _
                IntTaquilla, n
                If n = 0 Then
                    cnnConexion.Execute "INSERT INTO Taquillas(IDTaquilla,OpenSaldo,Estado,Cuadre," & _
                    "Fecha,Hora,Usuario) VALUES(" & IntTaquilla & ",0,-1,0,Date(),Time(),'" & _
                    gcUsuario & "')"
                End If
                Call rtnEstadoCaja("True")
                strFCaja = Format(Date, "mm/dd/yyyy")
                Unload Me
                Set FrmCaja = Nothing
                MsgBox "Caja  '" & Format(IntTaquilla, "00") & "' Abierta", vbInformation, App.ProductName
                Call rtnBitacora("Apertura Caja " & IntTaquilla)
                '
                
                
            Case "Cerrar Caja"
    '       ---------------------
                'cnnConexion.Mode = adModeShareDenyNone
                cnnConexion.Execute "UPDATE Taquillas SET Estado=FALSE,Fecha=Date(), Hora=" _
                & "Time(), Usuario='" & gcUsuario & "',Cuadre=False WHERE IDTaquilla=" _
                & IntTaquilla
                rtnEstadoCaja ("False")
                cnnConexion.Execute "UPDATE Inmueble SET DeudaAct=Deuda,Fondo=FondoAct;"
                Unload Me
                Set FrmCaja = Nothing
                MsgBox "Caja " & IntTaquilla & " Cerrada", vbInformation, App.ProductName
                Call rtnBitacora("Cierre Caja " & IntTaquilla)
                '
        End Select
    
    Case 1  'CANCELAR / CIERRA EL FORMULARIO
'   ---------------------------------------------
        Unload Me
        Set FrmCaja = Nothing
        
End Select
'
End Sub

    
    Private Sub Form_Load()
    Me.Visible = False
    DoEvents
    CenterForm Me
    Me.Caption = Titulo
    Label1(0) = "Taquilla Nº " & Format(IntTaquilla, "00")
    Label1(1) = Date & " / " & Time
    Label1(3) = " " & gcUsuario
    CmdTaquilla(0).Caption = boton
    If boton = "&Totales" Then CmdTaquilla(0).Enabled = True
    CmdTaquilla(0).Picture = LoadResPicture("Llave", vbResIcon)
    CmdTaquilla(1).Picture = LoadResPicture("SALIR", vbResIcon)
    Me.Show
    Me.Visible = True
    End Sub

    Private Sub TxtMonto_Change(): If Len(TxtMonto) = 0 Then CmdTaquilla(0).Enabled = False
    End Sub

    Private Sub TxtMonto_KeyPress(KeyAscii As Integer)
    '
    If KeyAscii = 46 Then KeyAscii = 44 'CONVIERTE COMA A PUNTO
    'VALIDA LA ENTRADA DE DATOS
    Call Validacion(KeyAscii, "1234567890.")
    If KeyAscii = 13 Then   'PRESIONO ENTER
        TxtMonto = Format(TxtMonto, "#,##0.00")
        CmdTaquilla(0).Enabled = True
        CmdTaquilla(0).SetFocus
    End If
    '
    End Sub

