VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmEmpresa 
   Caption         =   "Datos de la Empresa"
   ClientHeight    =   4650
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8535
   Icon            =   "frmEmpresa.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4650
   ScaleWidth      =   8535
   StartUpPosition =   3  'Windows Default
   Begin MSAdodcLib.Adodc adoEmpresa 
      Height          =   435
      Left            =   1065
      Top             =   3990
      Visible         =   0   'False
      Width           =   2715
      _ExtentX        =   4789
      _ExtentY        =   767
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   2
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
      RecordSource    =   "Empresa"
      Caption         =   "Datos Empresa"
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
   Begin MSComDlg.CommonDialog Dialogo 
      Left            =   525
      Top             =   3975
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmd 
      Cancel          =   -1  'True
      Caption         =   "&Cerrar"
      Height          =   495
      Index           =   1
      Left            =   6645
      TabIndex        =   10
      Top             =   3975
      Width           =   1215
   End
   Begin VB.CommandButton cmd 
      Caption         =   "Aceptar"
      Height          =   495
      Index           =   0
      Left            =   5250
      TabIndex        =   9
      Top             =   3990
      Width           =   1215
   End
   Begin VB.TextBox txt 
      DataField       =   "codProv"
      DataSource      =   "adoEmpresa"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   3
      Left            =   2955
      TabIndex        =   7
      Top             =   3345
      Width           =   1515
   End
   Begin VB.TextBox txt 
      DataField       =   "codCaja"
      DataSource      =   "adoEmpresa"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   2
      Left            =   2955
      TabIndex        =   6
      Top             =   2595
      Width           =   1515
   End
   Begin VB.TextBox txt 
      DataField       =   "codInm"
      DataSource      =   "adoEmpresa"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   1
      Left            =   2955
      TabIndex        =   5
      Top             =   1890
      Width           =   1515
   End
   Begin VB.TextBox txt 
      DataField       =   "Nombre"
      DataSource      =   "adoEmpresa"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   0
      Left            =   2955
      TabIndex        =   1
      Top             =   1245
      Width           =   4950
   End
   Begin VB.Label lbl 
      BackStyle       =   0  'Transparent
      Caption         =   "Click para cambiar logo"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   285
      Index           =   5
      Left            =   4755
      MouseIcon       =   "frmEmpresa.frx":2AFA
      MousePointer    =   99  'Custom
      TabIndex        =   11
      Top             =   1815
      Width           =   2295
   End
   Begin VB.Image imgLogo 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      DataSource      =   "adoEmpresa"
      Height          =   1920
      Left            =   4695
      Stretch         =   -1  'True
      Top             =   1800
      Width           =   3195
   End
   Begin VB.Label lbl 
      Caption         =   $"frmEmpresa.frx":2E04
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Index           =   4
      Left            =   510
      TabIndex        =   8
      Top             =   390
      Width           =   7530
   End
   Begin VB.Label lbl 
      Caption         =   "Código de Caja:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   3
      Left            =   495
      TabIndex        =   4
      Top             =   2670
      Width           =   2295
   End
   Begin VB.Label lbl 
      Caption         =   "Código de Proveedor:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   2
      Left            =   525
      TabIndex        =   3
      Top             =   3435
      Width           =   2295
   End
   Begin VB.Label lbl 
      Caption         =   "Código Inmueble:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   1
      Left            =   525
      TabIndex        =   2
      Top             =   1965
      Width           =   2295
   End
   Begin VB.Label lbl 
      Caption         =   "Nombre ó Razón Social:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   0
      Left            =   525
      TabIndex        =   0
      Top             =   1290
      Width           =   2295
   End
End
Attribute VB_Name = "frmEmpresa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim strLogo As String
Dim strLogo1 As String

Private Sub cmd_Click(Index As Integer)
On Error Resume Next
If Index = 0 Then
    If strLogo1 = "" Then
        strLogo1 = strLogo
    Else
        FileCopy strLogo, strLogo1
    End If
    adoEmpresa.Recordset("logo") = strLogo1
    adoEmpresa.Recordset("Usuario") = gcUsuario
    adoEmpresa.Recordset("Fecha") = Date
    adoEmpresa.Recordset.Update
Else
    Unload Me
End If
End Sub

Private Sub Form_Load()
adoEmpresa.ConnectionString = cnnOLEDB & gcPath & "\sac.mdb"
adoEmpresa.Refresh
strLogo = IIf(IsNull(adoEmpresa.Recordset("Logo")), "", adoEmpresa.Recordset("Logo"))
If Dir(strLogo) <> "" Then imgLogo = LoadPicture(strLogo)
End Sub

Private Sub lbl_Click(Index As Integer)
On Error Resume Next
If Index = 5 Then

    Dialogo.CancelError = True
    Dialogo.Filter = "Imágenes (*.jpg;*.gif)|*.jpg;*.gif"
    Dialogo.FilterIndex = 1
    Dialogo.DialogTitle = "Cargar Logo"
    Dialogo.ShowOpen
    
    If Err.Number = cdlCancel Then Exit Sub
    strLogo = Dialogo.FileName
    imgLogo.Picture = LoadPicture(strLogo)
    imgLogo.Refresh
    strLogo1 = gcPath & "\" & Dialogo.FileTitle
    
End If

End Sub
