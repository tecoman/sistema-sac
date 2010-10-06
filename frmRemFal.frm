VERSION 5.00
Begin VB.Form frmRemFal 
   Caption         =   "Facturas Pendientes"
   ClientHeight    =   4080
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   4080
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame fra 
      Caption         =   "Seleccione el Servicio"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3540
      Index           =   1
      Left            =   195
      TabIndex        =   5
      Top             =   270
      Width           =   4260
      Begin VB.CommandButton cmd 
         Cancel          =   -1  'True
         Caption         =   "Salir"
         Height          =   765
         Index           =   1
         Left            =   2295
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   2565
         Width           =   1005
      End
      Begin VB.CommandButton cmd 
         Caption         =   "Aceptar"
         Height          =   765
         Index           =   0
         Left            =   960
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   2565
         Width           =   1005
      End
      Begin VB.Frame fra 
         Caption         =   "Periodo ( Mes - Año )"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1110
         Index           =   0
         Left            =   255
         TabIndex        =   6
         Top             =   1230
         Width           =   3750
         Begin VB.ComboBox cmb 
            DataField       =   "TipoMovimientoCaja"
            Height          =   315
            Index           =   0
            ItemData        =   "frmRemFal.frx":0000
            Left            =   465
            List            =   "frmRemFal.frx":002B
            Style           =   2  'Dropdown List
            TabIndex        =   1
            Top             =   435
            Width           =   1320
         End
         Begin VB.ComboBox cmb 
            DataField       =   "TipoMovimientoCaja"
            Height          =   315
            Index           =   1
            ItemData        =   "frmRemFal.frx":0094
            Left            =   1875
            List            =   "frmRemFal.frx":0096
            TabIndex        =   2
            Top             =   435
            Width           =   1290
         End
      End
      Begin VB.ComboBox cmb 
         DataField       =   "TipoMovimientoCaja"
         Height          =   315
         Index           =   2
         ItemData        =   "frmRemFal.frx":0098
         Left            =   300
         List            =   "frmRemFal.frx":009A
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   570
         Width           =   3690
      End
   End
End
Attribute VB_Name = "frmRemFal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Cmd_Click(Index As Integer)

Select Case Index
    Case 1
        Unload Me
        Set frmRemFal = Nothing
        
    Case 0
        Call Procesar_Informacion
        
End Select

End Sub

Private Sub Form_Load()
'variable locales
Dim rstlocal As New ADODB.Recordset

Cmd(0).Picture = LoadResPicture("OK", vbResIcon)
Cmd(1).Picture = LoadResPicture("SALIR", vbResIcon)
'
rstlocal.Open "ServiciosTipo", cnnConexion, adOpenKeyset, adLockOptimistic, adCmdTable
If Not rstlocal.EOF And Not rstlocal.BOF Then

    rstlocal.MoveFirst
    Do
        cmb(2).AddItem rstlocal!Descripcion
        rstlocal.MoveNext
    Loop Until rstlocal.EOF
End If
rstlocal.Close
For I = Year(Date) To Year(Date) - 3 Step -1
    cmb(1).AddItem I
Next
Call CenterForm(frmRemFal)
Set rstlocal = Nothing
End Sub

Private Sub Procesar_Informacion()
'variables locales
Dim rpReporte As ctlReport
Dim rstlocal As New ADODB.Recordset
'valida los datos mínimos requeriso
If cmb(2).ListIndex = -1 Then
    MsgBox "Seleccione el tipo de servicio", vbCritical, App.ProductName
    Exit Sub
End If
If Not IsDate("01/" & cmb(0) & "/" & cmb(1)) Then
    MsgBox "Complete los datos del período", vbCritical, App.ProductName
    Exit Sub
End If
'
cnnConexion.Execute "DELETE * FROM tempRemesa"
With ObjCmd
    .ActiveConnection = cnnConexion
    .CommandType = adCmdText
    .CommandText = "INSERT INTO tempRemesa SELECT * FROM qdfRemesa1"
    .Parameters(0) = cmb(2).ListIndex
    .Parameters(1) = Format(CDate("01/" & cmb(0) & "/" & cmb(1)), "mm/yyyy")
    .Execute N
End With
'
For I = 0 To 1000000
Next

'imprimir reporte
Set rpReporte = New ctlReport
With rpReporte
    '.Reset
    '.ProgressDialog = False
    .Reporte = gcReport & "cxp_remesa_pen.rpt"
    .OrigenDatos(0) = gcPath & "\sac.mdb"
    .Formulas(0) = "Periodo='" & ObjCmd.Parameters(1) & "'"
    .Formulas(1) = "subtitulo='" & cmb(2) & "'"
    .Salida = crPantalla
    .TituloVentana = "Facturas Pendientes Remesa"
    Call rtnBitacora("Imprimiendo Fact.Fal.Rem: " & cmb(2) & "/" & _
    Format(CDate("01/" & cmb(0) & "/" & cmb(1)), "mm/yyyy"))
    .Imprimir
    Unload Me
    Set frmRemFal = Nothing
End With
Set rpReporte = Nothing
'
End Sub
