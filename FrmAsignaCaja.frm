VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Begin VB.Form FrmAsignaCaja 
   Caption         =   "Asignación de Caja"
   ClientHeight    =   2790
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5175
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2790
   ScaleWidth      =   5175
   Visible         =   0   'False
   Begin ComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   315
      Left            =   0
      TabIndex        =   10
      Top             =   2475
      Width           =   5175
      _ExtentX        =   9128
      _ExtentY        =   556
      Style           =   1
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   1
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Salir"
      Height          =   495
      Index           =   1
      Left            =   3000
      TabIndex        =   9
      Top             =   1920
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Aceptar"
      Height          =   495
      Index           =   0
      Left            =   1320
      TabIndex        =   8
      Top             =   1920
      Width           =   1215
   End
   Begin VB.TextBox txtDescripCaja 
      DataField       =   "DescripCaja"
      Height          =   315
      Left            =   1605
      TabIndex        =   7
      Top             =   1425
      Width           =   3375
   End
   Begin VB.TextBox txtNombre 
      Height          =   315
      Left            =   1605
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   585
      Width           =   3375
   End
   Begin MSDataListLib.DataCombo ctlCodInm 
      Bindings        =   "FrmAsignaCaja.frx":0000
      DataField       =   "CodInm"
      Height          =   315
      Left            =   1605
      TabIndex        =   1
      Top             =   195
      Width           =   1140
      _ExtentX        =   2011
      _ExtentY        =   556
      _Version        =   393216
      ListField       =   "CodInm"
      BoundColumn     =   "Nombre"
      Text            =   "DataCombo1"
      Object.DataMember      =   ""
   End
   Begin MSDataListLib.DataCombo ctlCaja 
      Bindings        =   "FrmAsignaCaja.frx":000B
      DataField       =   "Caja"
      Height          =   315
      Left            =   1605
      TabIndex        =   5
      Top             =   1005
      Width           =   1140
      _ExtentX        =   2011
      _ExtentY        =   556
      _Version        =   393216
      ListField       =   "CodigoCaja"
      BoundColumn     =   "DescripCaja"
      Text            =   "DataCombo1"
      Object.DataMember      =   ""
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Descripción"
      Height          =   195
      Index           =   3
      Left            =   645
      TabIndex        =   6
      Top             =   1470
      Width           =   840
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&Caja N°:"
      Height          =   195
      Index           =   2
      Left            =   900
      TabIndex        =   4
      Top             =   1020
      Width           =   585
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Nombre:"
      Height          =   195
      Index           =   1
      Left            =   885
      TabIndex        =   2
      Top             =   645
      Width           =   600
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Código de &Inmueble"
      Height          =   195
      Index           =   0
      Left            =   75
      TabIndex        =   0
      Top             =   240
      Width           =   1410
   End
End
Attribute VB_Name = "FrmAsignaCaja"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
    Dim rstlocal As New ADODB.Recordset
    
    Private Sub Command1_Click(Index As Integer)
    '
    Select Case Index
        '
        Case 0 'Guardar
            On Error Resume Next
            If ctlCodInm = "" Or ctlCaja = "" Then
                MsgBox "Falta parametro...", vbExclamation, App.ProductName
                Exit Sub
            End If
            If Not ctlCaja.MatchedWithList Then
                cnnConexion.Execute "INSERT INTO Caja(CodigoCaja,DescripCaja,MonedaCaja) VALUES" _
                & "('" & ctlCaja & "','" & txtDescripCaja & "','BS');"
            End If
            cnnConexion.Execute "UPDATE Inmueble SET Caja='" & ctlCaja & "' WHERE CodInm ='" _
            & ctlCodInm & "';"
            FrmInmueble.Caja = ctlCaja
            StatusBar1.SimpleText = "Cambios Realizados con éxito..."
            Unload Me
            Set FrmAsignaCaja = Nothing
            '
        Case 1 'cancela la operacion
        Unload Me
        Set FrmAsignaCaja = Nothing
        '
    End Select
    '
    End Sub

    Private Sub Command1_GotFocus(Index As Integer)
    '
    Select Case Index
        '
        Case 0
        StatusBar1.SimpleText = "Presione Aceptar para procesar la información..."
        '
        Case 1
        StatusBar1.SimpleText = "Presione Salir para cerrar esta ventana..."
        '
    End Select
    '
    End Sub

    Private Sub ctlCaja_Change()
    '
    If ctlCaja.MatchedWithList Then
        txtDescripCaja = ctlCaja.BoundText
    Else
        txtDescripCaja = ""
    End If
    '
    End Sub


    Private Sub ctlCaja_GotFocus()
    StatusBar1.SimpleText = "Ingrese Codigo de Caja a Asignar..."
    End Sub

    Private Sub ctlCaja_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then Command1(0).SetFocus
    End Sub

    Private Sub ctlCodInm_Change()
    '
    If ctlCodInm.MatchedWithList Then
        txtNombre = ctlCodInm.BoundText
        Dim rstCaja As New ADODB.Recordset
        rstCaja.Open "SELECT Caja FROM Inmueble WHERE CodInm='" & ctlCodInm.Text & "';", _
        cnnConexion, adOpenStatic, adLockReadOnly
        If rstCaja.RecordCount > 0 Then
            ctlCaja = IIf(IsNull(rstCaja.Fields(0)), "", rstCaja.Fields(0))
        End If
        rstCaja.Close
        Set rstCaja = Nothing
    Else
        txtNombre = ""
    End If
    '
    End Sub

    Private Sub ctlCodInm_GotFocus()
    StatusBar1.SimpleText = "Ingrese el Codigo de Condomino...."
    End Sub

    Private Sub ctlCodInm_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then ctlCaja.SetFocus
    KeyAscii = 0
    End Sub

    Private Sub Form_Load()
    Me.Left = 3675
    Me.Top = 4275
    rstlocal.Open "Caja", cnnConexion, adOpenKeyset, adLockOptimistic, adCmdTable
    Set ctlCodInm.DataSource = FrmAdmin.objRst
    Set ctlCaja.DataSource = FrmAdmin.objRst
    Set txtDescripCaja.DataSource = rstlocal
    Set ctlCodInm.RowSource = FrmAdmin.objRst
    Set ctlCaja.RowSource = rstlocal
    
    End Sub


    Private Sub txtDescripCaja_KeyPress(KeyAscii%): KeyAscii = Asc(UCase(Chr(KeyAscii)))
    End Sub
