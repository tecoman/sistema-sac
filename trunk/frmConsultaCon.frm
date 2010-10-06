VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmConsultaCon 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Convenimiento de Pago"
   ClientHeight    =   4635
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   4680
   Icon            =   "frmConsultaCon.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4635
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmd 
      Caption         =   "&Cerrar"
      Height          =   495
      Left            =   3285
      TabIndex        =   0
      Top             =   3960
      Width           =   1215
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid Grid 
      Height          =   2580
      Left            =   210
      TabIndex        =   2
      Tag             =   "400|1200|1200|800"
      Top             =   1110
      Width           =   4245
      _ExtentX        =   7488
      _ExtentY        =   4551
      _Version        =   393216
      Cols            =   4
      FixedCols       =   0
      BackColorFixed  =   -2147483646
      ForeColorFixed  =   -2147483639
      BackColorBkg    =   -2147483636
      SelectionMode   =   1
      FormatString    =   "Nº |Fecha |>Monto |Pagó"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   4
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   3855
      Picture         =   "frmConsultaCon.frx":000C
      Top             =   375
      Width           =   480
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
   Begin VB.Image img 
      Enabled         =   0   'False
      Height          =   150
      Index           =   0
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   165
   End
   Begin VB.Label lbl 
      Caption         =   "Este propietario tiene convenimiento de pago. Con el siguiente detalle:"
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
      TabIndex        =   1
      Top             =   300
      Width           =   3630
   End
End
Attribute VB_Name = "frmConsultaCon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Cmd_Click()
Unload Me
Set frmConsultaCon = Nothing
End Sub

Private Sub Form_Load()
'variables locals
Dim rstConvenio As New ADODB.Recordset
Dim strSql$, StrInmueble$, strApartamento$
'centra el formulario
Call CenterForm(frmConsultaCon)

StrInmueble = FrmConsultaCxC.Dat(0)
strApartamento = FrmConsultaCxC.Dat(2)
strSql = "SELECT Convenio.*, Convenio_Detalle.* FROM Convenio INNER JOIN Convenio_Detalle O" _
& "N Convenio.IDConvenio = Convenio_Detalle.IDConvenio WHERE Convenio.CodInm='" & StrInmueble & _
"' AND Convenio.CodProp ='" & strApartamento & "' AND IDStatus = 1"

rstConvenio.Open strSql, cnnConexion, adOpenKeyset, adLockOptimistic, adCmdText

'variables locales
If Not (rstConvenio.EOF And rstConvenio.BOF) Then
    'presenta la informacion en el grid
    Call centra_titulo(Grid, True)
    img(0).Picture = LoadResPicture("Unchecked", vbResBitmap)
    img(1).Picture = LoadResPicture("Checked", vbResBitmap)
    Grid.Rows = 2
    Call rtnLimpiar_Grid(Grid)
    rstConvenio.MoveFirst
    Grid.Rows = rstConvenio.RecordCount + 3
    i = i + 1
    Grid.TextMatrix(i, 1) = "GASTOS COB."
    Grid.TextMatrix(i, 2) = Format(rstConvenio("Gastos") - rstConvenio("DedGasto"), "#,##0.00")
    Call Marca(rstConvenio("CanceladoG"), i)
    i = i + 1
    Grid.TextMatrix(i, 1) = "HONORARIOS EXT."
    Grid.TextMatrix(i, 2) = Format(rstConvenio("Honorarios") - rstConvenio("DedHono"), "#,##0.00")
    Call Marca(rstConvenio("CanceladoH"), i)
    Do
        i = i + 1
        Grid.TextMatrix(i, 0) = Format(i - 2, "00")
        Grid.TextMatrix(i, 1) = rstConvenio("Convenio_Detalle.Fecha")
        Grid.TextMatrix(i, 2) = Format(rstConvenio("Monto"), "#,##0.00")
        Call Marca(rstConvenio("Cancelada"), i)
        rstConvenio.MoveNext
    Loop Until rstConvenio.EOF

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

