VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmMsg 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Mensajería SAC 1.1"
   ClientHeight    =   5655
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5250
   Icon            =   "frmMsg.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmMsg.frx":27A2
   ScaleHeight     =   377
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   350
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Index           =   1
      Left            =   4650
      Top             =   30
   End
   Begin VB.Timer Timer1 
      Index           =   0
      Left            =   4215
      Top             =   30
   End
   Begin VB.CommandButton cmd 
      Caption         =   "&Limpiar"
      Height          =   390
      Index           =   3
      Left            =   2625
      TabIndex        =   7
      Top             =   4815
      Width           =   1050
   End
   Begin VB.CommandButton cmd 
      Caption         =   "&Salir"
      Height          =   390
      Index           =   2
      Left            =   540
      TabIndex        =   6
      Top             =   4815
      Width           =   1050
   End
   Begin VB.CommandButton cmd 
      Caption         =   "&Conectar"
      Enabled         =   0   'False
      Height          =   390
      Index           =   1
      Left            =   1590
      TabIndex        =   5
      Tag             =   "0"
      Top             =   4815
      Width           =   1050
   End
   Begin MSDataListLib.DataCombo dtc 
      Bindings        =   "frmMsg.frx":4EC8
      Height          =   315
      Left            =   1365
      TabIndex        =   4
      Top             =   4170
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   556
      _Version        =   393216
      ListField       =   "NombreUsuario"
      Text            =   "--Selecciona un usuario--"
   End
   Begin VB.CommandButton cmd 
      Caption         =   "&Enviar"
      Enabled         =   0   'False
      Height          =   390
      Index           =   0
      Left            =   3660
      TabIndex        =   2
      Top             =   4815
      Width           =   1050
   End
   Begin VB.TextBox txt 
      Height          =   675
      Index           =   1
      Left            =   195
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   3300
      Width           =   4890
   End
   Begin VB.TextBox txt 
      Height          =   1935
      Index           =   0
      Left            =   195
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   1080
      Width           =   4890
   End
   Begin MSWinsockLib.Winsock wsServidor 
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      LocalPort       =   888
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   465
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      RemotePort      =   888
   End
   Begin MSAdodcLib.Adodc adoUser 
      Height          =   360
      Left            =   210
      Top             =   2370
      Visible         =   0   'False
      Width           =   2235
      _ExtentX        =   3942
      _ExtentY        =   635
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
      Caption         =   "Usuarios en Linea"
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
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      Index           =   1
      X1              =   1
      X2              =   348
      Y1              =   311
      Y2              =   311
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00C0C0C0&
      BorderWidth     =   2
      Index           =   0
      X1              =   1
      X2              =   348
      Y1              =   311
      Y2              =   311
   End
   Begin VB.Label lbl 
      Height          =   255
      Index           =   1
      Left            =   90
      TabIndex        =   8
      Top             =   5400
      Width           =   5085
   End
   Begin VB.Label lbl 
      BackStyle       =   0  'Transparent
      Caption         =   "Conectar con:"
      ForeColor       =   &H00004080&
      Height          =   270
      Index           =   0
      Left            =   240
      TabIndex        =   3
      Top             =   4215
      Width           =   1215
   End
End
Attribute VB_Name = "frmMsg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmd_Click(Index As Integer)
Dim remoto As String
'
If Index = 1 Then 'conectar/desconectar
    If cmd(1).Tag = 0 Then  'conectar
        If dtc.Tag = "" Or IsNull(dtc.Tag) Then
            MsgBox "Debe seleccionar un usuario", vbInformation, App.ProductName
        Else
            remoto = dtc.Tag
            Winsock1.Close
            Winsock1.Connect remoto, 888
            txt(0) = txt(0) + "Conectado con " & dtc & vbNewLine
            cmd(1).Caption = "&Desconect."
            dtc.Locked = True
            cmd(1).Tag = 1
            cmd(0).Enabled = True
        End If
    Else    'desconectar
        Winsock1.Close
        cmd(1).Caption = "&Conectar"
        dtc.Locked = False
        cmd(1).Tag = 0
        txt(0) = txt(0) + "Conexión Finalizada...." & vbNewLine
        cmd(0).Enabled = False
    End If
    
ElseIf Index = 0 Then  'enviar datos
    Call enviar_msg
ElseIf Index = 2 Then
    Unload Me
    Set frmMsg = Nothing
Else
    txt(0) = ""
End If
'
End Sub

Private Sub dtc_Change()
With adoUser.Recordset
    .MoveFirst
    .Find "NombreUsuario='" & dtc & "'"
    If Not .EOF Then dtc.Tag = !Maquina
End With
Winsock1.Close
End Sub

Private Sub dtc_Click(Area As Integer)
If Area = 0 Then adoUser.Recordset.Requery
End Sub

Private Sub dtc_KeyPress(KeyAscii As Integer): KeyAscii = 0
End Sub

Private Sub Form_Load()

With adoUser

    .ConnectionString = cnnOLEDB + gcPath & "\tablas.mdb"
    .CursorLocation = adUseClient
    .CursorType = adOpenKeyset
    .CommandType = adCmdText
    .RecordSource = "SELECT * FROM Usuarios WHERE LogIn = True;"
    .Refresh

End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
'variables locales
If Me.Tag <> "" Then
    FrmAutorizacion.Inm = Left(Me.Tag, InStr(Me.Tag, "|") - 1)
    FrmAutorizacion.apto = Right(Me.Tag, Len(Me.Tag) - InStr(Me.Tag, "|"))
    Me.Tag = ""
    Me.Hide
    If FrmAdmin.WindowState = vbMinimized Then FrmAdmin.WindowState = vbMaximized
    Load FrmAutorizacion
    Unload Me
End If
'
End Sub

Private Sub Timer1_Timer(Index As Integer)
If Index = 0 Then
    lbl(1) = VerEstado(wsServidor.State)
Else
    lbl(1) = VerEstado(Winsock1.State)
End If
End Sub

Private Sub txt_Change(Index As Integer)
If Index = 0 Then
    txt(0).SelStart = Len(txt(0))
    'txt(1).SetFocus
End If
End Sub

Private Sub txt_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii = 13 Then
    If Winsock1.State = 7 Then Call enviar_msg
    KeyAscii = 0
End If
End Sub

Private Sub Winsock1_Close()
txt(0) = txt(0) + "Conexión cerrada" + vbNewLine
Winsock1.Close
cmd(1).Caption = "&Conectar"
dtc.Locked = False
End Sub

Private Sub Winsock1_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, _
ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, _
CancelDisplay As Boolean)
txt(0) = txt(0) + "Error " & Number & " -- " & Description & vbNewLine
cmd(1).Caption = "&Conectar"
dtc.Locked = False
cmd(1).Tag = 0
cmd(0).Enabled = False
End Sub

Private Sub Winsock1_SendComplete()
lbl(1) = "Mensaje enviado a " & Winsock1.RemoteHost
End Sub

Private Sub wsServidor_Close()
txt(0) = txt(0) + "La conexion ha sido cerrada desde el equipo remoto" & vbNewLine
wsServidor.Close
End Sub

Public Sub wsServidor_ConnectionRequest(ByVal requestID As Long)
wsServidor.Close
wsServidor.Accept requestID
txt(0) = "Conexión Aceptada desde ---" & wsServidor.RemoteHostIP & vbCrLf
End Sub

Private Sub wsServidor_DataArrival(ByVal bytesTotal As Long)
'
Dim datos As String
Dim Fin As String
wsServidor.GetData datos
Fin = UCase(Mid(datos, InStr(datos, vbNewLine) - 6, 6))

If datos Like "*AUTORIZACION DEDUCCION*" Then
    
    'MsgBox datos, vbInformation, App.ProductName
    txt(0) = datos & vbCrLf & "Presione Salir para continuar..."
    txt(1).Visible = False
    lbl(0).Visible = False
    dtc.Visible = False
    cmd(1).Enabled = False
    cmd(0).Enabled = False
    cmd(3).Enabled = False
    Me.Tag = deduccion(0, datos) & "|" & deduccion(1, datos)
    Me.Show vbModeless, FrmAdmin
    
'    FrmAutorizacion.Inm = deduccion(0, datos)
'    FrmAutorizacion.apto = deduccion(1, datos)
'    Load FrmAutorizacion
'    Unload Me
    
ElseIf Fin = "ENDSAC" Then
    'MsgBox "fin", vbInformation
    
    SetTimer hWnd, NV_CLOSEMSGBOX, 4000&, AddressOf TimerProc
    Call MessageBox(hWnd, "SAC fué cerrado por el administrador de sistema", _
    App.ProductName, vbInformation)

    Call rtnBitacora("Cierre Remoto " & Left(datos, InStr(datos, " ")) & " -- " & _
    wsServidor.RemoteHostIP)
    Unload FrmAdmin
    End
    
Else
    Me.Show vbModeless, FrmAdmin
    txt(0) = txt(0) + datos
End If
'
End Sub

Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)
Dim datos As String
Winsock1.GetData datos
txt(0) = txt(0) + datos
End Sub

Private Sub enviar_msg()
Dim enviar As String
If txt(1) <> "" Then
    enviar = gcUsuario & ": " & txt(1)
    If Winsock1.State <> sckClosed Then
        Winsock1.SendData enviar & vbNewLine
    Else
        wsServidor.SendData enviar & vbNewLine
    End If
    txt(0) = txt(0) + gcUsuario & ": " + txt(1) + vbCrLf
    txt(1) = ""
    txt(1).SetFocus
End If
End Sub

Private Function VerEstado(Estado As Byte) As String

    Select Case Estado
        Case 0
            VerEstado = "Sin Conexiones"
        Case 1
            VerEstado = "Abierto"
        Case 2
            VerEstado = "Esperando Conexion"
        Case 3
            VerEstado = "Conexion Pendiente"
        Case 4
            VerEstado = "Resolviendo Host"
        Case 5
            VerEstado = "Host Resuelto"
        Case 6
            VerEstado = "Conectando"
        Case 7
            VerEstado = "Conectado con -- " & IIf(wsServidor.RemoteHostIP = "", Winsock1.RemoteHost, wsServidor.RemoteHostIP)
        Case 8
            VerEstado = "Cerrando Conexion"
            wsServidor.Close
            wsServidor.Listen
            
        Case 9: VerEstado = "Error"
        End Select
End Function


Private Function deduccion(InmuebleApartamento As Byte, Mensaje As String) As String
Dim intPos As Integer, intSpace As Integer
Dim strNL As String
intPos = InStr(Mensaje, "/")

If InmuebleApartamento = 0 Then 'Inmueble
    intSpace = InStrRev(Mensaje, " ")
    deduccion = Trim(Mid(Mensaje, intSpace, intPos - intSpace))
Else 'Apartamento
    strNL = vbNewLine
    intSpace = InStrRev(Mensaje, strNL)
    deduccion = Trim(Mid(Mensaje, intPos + 1, intSpace - intPos - 1))
End If
'
End Function
