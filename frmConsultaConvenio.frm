VERSION 5.00
Begin VB.Form frmConCon 
   Caption         =   "!!ATENCION"
   ClientHeight    =   4440
   ClientLeft      =   4815
   ClientTop       =   2460
   ClientWidth     =   4680
   ControlBox      =   0   'False
   Icon            =   "frmConsultaConvenio.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4440
   ScaleWidth      =   4680
   Begin VB.CommandButton cmd 
      Caption         =   "&Cerrar"
      Height          =   495
      Left            =   3225
      TabIndex        =   2
      Top             =   3795
      Width           =   1215
   End
   Begin VB.PictureBox PIC 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   480
      Left            =   3855
      Picture         =   "frmConsultaConvenio.frx":5D52
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   1
      Top             =   375
      Width           =   480
   End
   Begin VB.Image img 
      Enabled         =   0   'False
      Height          =   150
      Index           =   0
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   165
   End
   Begin VB.Image img 
      Enabled         =   0   'False
      Height          =   150
      Index           =   1
      Left            =   195
      Top             =   0
      Visible         =   0   'False
      Width           =   165
   End
   Begin VB.Label lbl 
      Caption         =   "Este propietario tine convenimiento de pago. Con el siguiente detalle:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   690
      Index           =   0
      Left            =   315
      TabIndex        =   0
      Top             =   315
      Width           =   3630
   End
End
Attribute VB_Name = "frmConCon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Public rstConvenio As New ADODB.Recordset
Public Inmueble As String, apto As String

Private Sub cmd_Click()
Unload Me
Set frmConsultaConvenio = Nothing
End Sub
'
Private Sub Form_Load()
Dim rstConvenio As New ADODB.Recordset
Dim strSql As String
strSql = "SELECT Convenio.*, Convenio_Detalle.* FROM Convenio INNER JOIN Convenio_Detalle O" _
& "N Convenio.IDConvenio = Convenio_Detalle.IDConvenio WHERE Convenio.CodInm='" & Inmueble & _
"' AND Convenio.CodProp ='" & apto & "' AND IDStatus = 1"
rstConvenio.Open strSql, cnnConexion, adOpenKeyset, adLockOptimistic, adCmdText

'variables locales
If Not (rstConvenio.EOF And rstConvenio.BOF) Then
    'presenta la informacion en el grid
'    Call centra_titulo(Grid, True)
'    Img(0).Picture = LoadResPicture("Unchecked", vbResBitmap)
'    Img(1).Picture = LoadResPicture("Checked", vbResBitmap)
'    Grid.Rows = 2
'    Call rtnLimpiar_Grid(Grid)
'    rstConvenio.MoveFirst
'    Grid.Rows = rstConvenio.RecordCount + 3
'    I = I + 1
'    Grid.TextMatrix(I, 1) = "GASTOS COB."
'    Grid.TextMatrix(I, 2) = Format(rstConvenio("Gastos") - rstConvenio("DedGasto"), "#,##0.00")
'    Call Marca(rstConvenio("CanceladoG"), I)
'    I = I + 1
'    Grid.TextMatrix(I, 1) = "HONORARIOS EXT."
'    Grid.TextMatrix(I, 2) = Format(rstConvenio("Honorarios") - rstConvenio("DedHono"), "#,##0.00")
'    Call Marca(rstConvenio("CanceladoH"), I)
'    Do
'        I = I + 1
'        Grid.TextMatrix(I, 0) = Format(I - 2, "00")
'        Grid.TextMatrix(I, 1) = rstConvenio("Convenio_Detalle.Fecha")
'        Grid.TextMatrix(I, 2) = Format(rstConvenio("Monto"), "#,##0.00")
'        Call Marca(rstConvenio("Cancelada"), I)
'        rstConvenio.MoveNext
'    Loop Until rstConvenio.EOF
'    Me.Show vbModal, FrmAdmin
'
End If
rstConvenio.Close
Set rstConvenio = Nothing
'
End Sub

Private Sub Marca(V As Boolean, ByVal Fila As Long)
'variables locales
With Grid
    .Col = 3
    .Row = Fila
    Set .CellPicture = IIf(V, img(1), img(0))
    .CellPictureAlignment = flexAlignCenterCenter
End With
End Sub

