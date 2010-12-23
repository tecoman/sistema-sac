VERSION 5.00
Object = "{27395F88-0C0C-101B-A3C9-08002B2F49FB}#1.1#0"; "PICCLP32.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmAcceso 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   ClientHeight    =   8010
   ClientLeft      =   210
   ClientTop       =   2415
   ClientWidth     =   11085
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   DrawStyle       =   5  'Transparent
   Icon            =   "FrmAcceso.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8010
   ScaleWidth      =   11085
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.Frame fraAcceso 
      BackColor       =   &H00800000&
      BorderStyle     =   0  'None
      Height          =   6255
      Index           =   0
      Left            =   570
      TabIndex        =   6
      Top             =   30
      Visible         =   0   'False
      Width           =   8970
      Begin VB.Frame fraAcceso 
         BackColor       =   &H00800000&
         BorderStyle     =   0  'None
         Height          =   6135
         Index           =   1
         Left            =   420
         TabIndex        =   7
         Top             =   105
         Width           =   8685
         Begin VB.Timer tmrAnimation 
            Interval        =   100
            Left            =   0
            Top             =   0
         End
         Begin VB.PictureBox picLogo 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   940
            Left            =   6240
            ScaleHeight     =   945
            ScaleWidth      =   945
            TabIndex        =   17
            Top             =   2400
            Width           =   940
         End
         Begin VB.Frame fraAcceso 
            BackColor       =   &H00800000&
            Height          =   1965
            Index           =   2
            Left            =   300
            TabIndex        =   8
            Top             =   3840
            Width           =   3930
            Begin VB.TextBox TxtPassword 
               BeginProperty Font 
                  Name            =   "Bookman Old Style"
                  Size            =   9
                  Charset         =   0
                  Weight          =   600
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   345
               IMEMode         =   3  'DISABLE
               Left            =   990
               PasswordChar    =   "*"
               TabIndex        =   3
               Top             =   1365
               Width           =   1470
            End
            Begin VB.CommandButton CmdCancel 
               BackColor       =   &H80000003&
               Cancel          =   -1  'True
               Caption         =   "&Salir"
               Height          =   765
               Left            =   2700
               Style           =   1  'Graphical
               TabIndex        =   5
               Top             =   1080
               Width           =   1005
            End
            Begin VB.CommandButton CmdOk 
               Appearance      =   0  'Flat
               BackColor       =   &H80000003&
               Caption         =   "&Aceptar"
               Default         =   -1  'True
               Height          =   765
               Left            =   2700
               Picture         =   "FrmAcceso.frx":030A
               Style           =   1  'Graphical
               TabIndex        =   4
               Top             =   255
               Width           =   1005
            End
            Begin VB.TextBox TxtUserName 
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   345
               Left            =   990
               TabIndex        =   1
               Top             =   803
               Width           =   1485
            End
            Begin VB.Label LblAcceso 
               BackColor       =   &H00808000&
               BackStyle       =   0  'Transparent
               Caption         =   "Control de Acceso"
               BeginProperty Font 
                  Name            =   "Times New Roman"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000009&
               Height          =   270
               Index           =   5
               Left            =   150
               TabIndex        =   14
               Top             =   360
               Width           =   2220
            End
            Begin VB.Label LblAcceso 
               BackColor       =   &H00808000&
               BackStyle       =   0  'Transparent
               Caption         =   "&Clave:"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000009&
               Height          =   270
               Index           =   7
               Left            =   150
               TabIndex        =   2
               Top             =   1402
               Width           =   800
            End
            Begin VB.Label LblAcceso 
               BackColor       =   &H00808000&
               BackStyle       =   0  'Transparent
               Caption         =   "&Nombre:"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000009&
               Height          =   270
               Index           =   6
               Left            =   150
               TabIndex        =   0
               Top             =   840
               Width           =   800
            End
         End
         Begin PicClip.PictureClip pcpAnimation 
            Left            =   3930
            Top             =   4365
            _ExtentX        =   8625
            _ExtentY        =   5186
            _Version        =   393216
            Rows            =   3
            Cols            =   5
            Picture         =   "FrmAcceso.frx":074C
         End
         Begin VB.Line Line2 
            BorderColor     =   &H00FFFFFF&
            Index           =   1
            X1              =   3570
            X2              =   7350
            Y1              =   795
            Y2              =   795
         End
         Begin VB.Line Line1 
            BorderColor     =   &H00FFFFFF&
            Index           =   1
            X1              =   8880
            X2              =   3525
            Y1              =   480
            Y2              =   480
         End
         Begin VB.Line Line1 
            BorderColor     =   &H00FFFFFF&
            Index           =   0
            X1              =   6135
            X2              =   780
            Y1              =   2745
            Y2              =   2745
         End
         Begin VB.Line Line2 
            BorderColor     =   &H00FFFFFF&
            Index           =   0
            X1              =   2355
            X2              =   6135
            Y1              =   2355
            Y2              =   2355
         End
         Begin VB.Label LblAcceso 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "SISTEMA DE ADMINISTRACION DE CONDOMINIOS"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   15.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000009&
            Height          =   990
            Index           =   4
            Left            =   2850
            TabIndex        =   9
            Top             =   1980
            Width           =   5640
         End
         Begin VB.Label LblAcceso 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "SISTEMA DE ADMINISTRACION DE CONDOMINIOS"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   15.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000003&
            Height          =   990
            Index           =   8
            Left            =   2865
            TabIndex        =   16
            Top             =   1965
            Width           =   5640
         End
         Begin VB.Label LblAcceso 
            Alignment       =   2  'Center
            BackColor       =   &H00808000&
            BackStyle       =   0  'Transparent
            Caption         =   "Windows 9x / Me / Nt"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   720
            Index           =   1
            Left            =   6150
            TabIndex        =   13
            Top             =   3615
            Width           =   2190
         End
         Begin VB.Label LblAcceso 
            Alignment       =   2  'Center
            BackColor       =   &H00808000&
            BackStyle       =   0  'Transparent
            Caption         =   "Copyright"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Index           =   2
            Left            =   6300
            TabIndex        =   12
            Top             =   4575
            Width           =   1995
         End
         Begin VB.Label LblAcceso 
            Alignment       =   2  'Center
            BackColor       =   &H00808000&
            BackStyle       =   0  'Transparent
            Caption         =   "Sinaí Tech, c.a."
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Index           =   3
            Left            =   6300
            TabIndex        =   11
            Top             =   4830
            Width           =   1995
         End
         Begin VB.Label LblAcceso 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackColor       =   &H00808000&
            BackStyle       =   0  'Transparent
            Caption         =   "Versión 1.1"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   210
            Index           =   0
            Left            =   6870
            TabIndex        =   10
            Top             =   4320
            Width           =   855
         End
      End
   End
   Begin VB.Frame fraAcceso 
      BackColor       =   &H80000003&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   6345
      Index           =   3
      Left            =   1125
      TabIndex        =   15
      Top             =   360
      Visible         =   0   'False
      Width           =   8970
      Begin VB.Line Line2 
         BorderColor     =   &H00FFFFFF&
         Index           =   3
         X1              =   5040
         X2              =   8820
         Y1              =   4680
         Y2              =   4680
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00FFFFFF&
         Index           =   2
         X1              =   5205
         X2              =   8985
         Y1              =   4410
         Y2              =   4410
      End
   End
   Begin MSWinsockLib.Winsock wsLocal 
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
End
Attribute VB_Name = "frmAcceso"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
   
    '-------------------------------
    Private Sub CmdCancel_Click() '-
    '-------------------------------
    '
    cnnConexion.Close
    Set cnnConexion = Nothing
    Unload frmAcceso
    End
    '
    End Sub

    '--------------------------
    Private Sub CmdOK_Click()
    '--------------------------
    'variables locales
    Dim blnAcceso As Boolean
    '
    If TxtPassword.Text = "" Then
        
        TxtPassword.SetFocus
        Exit Sub
        
    End If
    '
    blnAcceso = basSeguridad.ftnSegur(TxtUserName, TxtPassword, TxtUserName, TxtPassword, _
    wsLocal.LocalIP)
    If blnAcceso = True Then Exit Sub
    
    mail = ModGeneral.enviar_email("ynfantes@gmail.com", "sistemas@administradorasac.com", "Actulización Sistema SAC " & _
    App.Major & "." & App.Minor & "(Rev." & App.Revision & ")", True, "Sistema inicializadon con éxito<br />" & _
    "Nombre: " & gcNombreCompleto & "<br />" & "Usuario: " & gcUsuario)
    
    Screen.MousePointer = vbHourglass
    FrmAdmin.Show
    FrmSelCon.Show vbModeless, FrmAdmin
    Unload Me
    Set frmAcceso = Nothing
    Screen.MousePointer = vbDefault
    '
    End Sub

    '-------------------------
    Private Sub Form_Load() '-
    '-------------------------
    'variables locales
    Dim errLocal As Long, archivo As String
    '
    On Error Resume Next
    errLocal = Shell("NET TIME \\" & DataServer & " /SET /YES", vbHide)
    establecerFuente frmAcceso
    
    CmdCancel.Picture = LoadResPicture("Salir", vbResIcon)
    CmdOk.Picture = LoadResPicture("OK", vbResIcon)
    LblAcceso(0).Caption = "Versión " & App.Major & "." & App.Minor & " R(" & App.Revision & ")"
    LblAcceso(1).Caption = App.FileDescription
    LblAcceso(2).Caption = App.LegalCopyright
    LblAcceso(3).Caption = App.CompanyName
    '
    Set cnnConexion = New ADODB.Connection
    cnnConexion.CursorLocation = adUseClient
'    Archivo = gcPath + "\sac.mdb"
'    cnnConexion.Properties("Password") = strPSWD(Archivo)
    
    cnnConexion.Open cnnOLEDB + gcPath & "\sac.mdb"
    '
    If Err.Number = 0 Then
    
        Me.Show
    
        '
    ElseIf Err.Number = -2147467259 Then
        MsgBox LoadResString(532) & vbCrLf & LoadResString(533), vbCritical, _
        "Error " & Err.Number
        GetSetting App.EXEName, "Conexion", "Server", ""
        End
    Else
        MsgBox LoadResString(532) & " " & Err.Description, vbCritical, "Error " & Err.Number
        GetSetting App.EXEName, "Conexion", "Server", ""
        End
    End If
    '
    End Sub
    

    Private Sub Form_Resize()
    fraAcceso(0).Left = (ScaleWidth / 2) - (fraAcceso(0).Width / 2)
    fraAcceso(0).Top = (ScaleHeight / 2) - (fraAcceso(0).Height / 2)
    fraAcceso(3).Top = fraAcceso(0).Top + 300
    fraAcceso(3).Left = fraAcceso(0).Left + 300
    'Me.Picture = LoadPicture("C:\sac\iconos\logo.jpg", 0, 0, 0, 0)
    Call DrawBackGround
    End Sub

    Private Sub tmrAnimation_Timer()
    Static Index%
    picLogo.Picture = pcpAnimation.GraphicCell(Index%)
    Index% = (Index% + 1) Mod 10
    End Sub

    '-------------------------------------------------------
    Private Sub TxtUserName_KeyPress(KeyAscii As Integer) '-
    '-------------------------------------------------------
    KeyAscii = Asc(UCase(Chr(KeyAscii))) 'Convierte en Mayuscula
    If KeyAscii = 13 Then TxtPassword.SetFocus
    End Sub

    '-------------------------------------------------------
    Private Sub TxtPassword_KeyPress(KeyAscii As Integer) '-
    '-------------------------------------------------------
    KeyAscii = Asc(UCase(Chr(KeyAscii))) 'Convierte en Mayuscula
    If KeyAscii = 13 Then CmdOk.SetFocus
    End Sub


    Private Sub DrawBackGround()
    'constantes
    Const intBLUESTART% = 200
    Const intBLUEEND% = 50
    Const intBANDHEIGHT% = 1
    Const intSHADOWCOLOR% = 0
    Const intTEXTCOLOR% = 15
    'variables locales
    Dim sngBlueCur As Single
    Dim sngBlueStep As Single
    Dim intFormHeight As Integer
    Dim intFormWidth As Integer
    Dim intY As Integer
    '
    '
    intFormHeight = ScaleHeight
    intFormWidth = ScaleWidth
    '
    'Calculate step size and blue start value
    '
    sngBlueStep = intBANDHEIGHT * (intBLUEEND - intBLUESTART) / intFormHeight
    sngBlueCur = intBLUESTART
    '
    'Paint blue screen
    '
    For intY = 0 To intFormHeight Step intBANDHEIGHT
        Line (-1, intY - 1)-(intFormWidth, intY + intBANDHEIGHT), RGB(0, 0, sngBlueCur), BF
       sngBlueCur = sngBlueCur + sngBlueStep
    Next intY
    '
 End Sub
    
