VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form frmGFMen 
   Caption         =   "Gastos Fijos Menores"
   ClientHeight    =   3765
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5175
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   3765
   ScaleWidth      =   5175
   Tag             =   "1"
   WindowState     =   2  'Maximized
   Begin ComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   480
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5175
      _ExtentX        =   9128
      _ExtentY        =   847
      ButtonWidth     =   714
      ButtonHeight    =   688
      AllowCustomize  =   0   'False
      Appearance      =   1
      ImageList       =   "ImageList1"
      _Version        =   327682
      BeginProperty Buttons {0713E452-850A-101B-AFC0-4210102A8DA7} 
         NumButtons      =   8
         BeginProperty Button1 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.ToolTipText     =   "Último Registro"
            Object.Tag             =   ""
            Style           =   4
            Object.Width           =   1e-4
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button2 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "NEW"
            Object.ToolTipText     =   "Nuevo Registro"
            Object.Tag             =   ""
            ImageIndex      =   5
         EndProperty
         BeginProperty Button3 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Enabled         =   0   'False
            Key             =   "SAVE"
            Object.ToolTipText     =   "Guardar (Ctl + G)"
            Object.Tag             =   ""
            ImageIndex      =   6
         EndProperty
         BeginProperty Button4 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Enabled         =   0   'False
            Key             =   "CANCEL"
            Object.ToolTipText     =   "Cancelar"
            Object.Tag             =   ""
            ImageIndex      =   8
         EndProperty
         BeginProperty Button5 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Enabled         =   0   'False
            Object.Visible         =   0   'False
            Key             =   "EDIT"
            Object.ToolTipText     =   "Editar Registro"
            Object.Tag             =   ""
            ImageIndex      =   10
         EndProperty
         BeginProperty Button6 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Enabled         =   0   'False
            Key             =   "DELETE"
            Object.ToolTipText     =   "Eliminar"
            Object.Tag             =   ""
            ImageIndex      =   9
         EndProperty
         BeginProperty Button7 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Enabled         =   0   'False
            Key             =   "PRINT"
            Object.ToolTipText     =   "Imprimir"
            Object.Tag             =   ""
            ImageIndex      =   11
         EndProperty
         BeginProperty Button8 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "CLOSE"
            Object.ToolTipText     =   "Cerrar"
            Object.Tag             =   ""
            ImageIndex      =   12
         EndProperty
      EndProperty
      BorderStyle     =   1
      MousePointer    =   99
      MouseIcon       =   "frmGFMen.frx":0000
   End
   Begin VB.Frame fra 
      Enabled         =   0   'False
      Height          =   9375
      Left            =   135
      TabIndex        =   1
      Top             =   630
      Width           =   13560
      Begin VB.CommandButton cmd 
         Caption         =   "&Agregar"
         Height          =   315
         Left            =   8775
         TabIndex        =   11
         Top             =   1575
         Width           =   1215
      End
      Begin MSDataListLib.DataCombo Dat 
         DataField       =   "I"
         DataSource      =   "ADOcontrol(0)"
         Height          =   315
         Index           =   0
         Left            =   615
         TabIndex        =   2
         ToolTipText     =   "Codigo del Inmueble"
         Top             =   735
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   -2147483643
         ListField       =   "CodInm"
         BoundColumn     =   ""
         Text            =   ""
         Object.DataMember      =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDataListLib.DataCombo Dat 
         DataSource      =   "ADOcontrol(0)"
         Height          =   315
         Index           =   1
         Left            =   2040
         TabIndex        =   3
         ToolTipText     =   "Codigo del Inmueble"
         Top             =   735
         Width           =   6225
         _ExtentX        =   10980
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   -2147483643
         ListField       =   "CodInm"
         BoundColumn     =   ""
         Text            =   ""
         Object.DataMember      =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDataListLib.DataCombo Dat 
         DataSource      =   "ADOcontrol(0)"
         Height          =   315
         Index           =   2
         Left            =   615
         TabIndex        =   4
         ToolTipText     =   "Codigo del Inmueble"
         Top             =   1575
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   -2147483643
         ListField       =   "CodInm"
         BoundColumn     =   ""
         Text            =   ""
         Object.DataMember      =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDataListLib.DataCombo Dat 
         DataField       =   "InmuebleMovimientoCaja"
         DataSource      =   "ADOcontrol(0)"
         Height          =   315
         Index           =   3
         Left            =   2040
         TabIndex        =   5
         ToolTipText     =   "Codigo del Inmueble"
         Top             =   1575
         Width           =   6225
         _ExtentX        =   10980
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   -2147483643
         ListField       =   "CodInm"
         BoundColumn     =   "CodInm"
         Text            =   ""
         Object.DataMember      =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSFlexGridLib.MSFlexGrid Flex 
         Height          =   3315
         Left            =   615
         TabIndex        =   10
         Tag             =   "Caja"
         Top             =   2355
         Width           =   7860
         _ExtentX        =   13864
         _ExtentY        =   5847
         _Version        =   393216
         Cols            =   5
         FixedCols       =   0
         RowHeightMin    =   285
         BackColorFixed  =   -2147483646
         ForeColorFixed  =   -2147483639
         BackColorSel    =   65280
         ForeColorSel    =   0
         GridColor       =   -2147483645
         WordWrap        =   -1  'True
         AllowBigSelection=   0   'False
         FocusRect       =   2
         HighLight       =   2
         ScrollBars      =   2
         AllowUserResizing=   1
         BorderStyle     =   0
         MousePointer    =   99
         FormatString    =   "Codigo Gasto |Detalle|Código|Proveedor |Sel "
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MouseIcon       =   "frmGFMen.frx":031A
      End
      Begin VB.Label Lbl 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         Caption         =   "No existen gastos fijos registrados en el catálogo de gastos de este condominio!!!"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   720
         Index           =   4
         Left            =   8910
         TabIndex        =   12
         Top             =   705
         Visible         =   0   'False
         Width           =   3585
      End
      Begin VB.Image imgRemesa 
         Enabled         =   0   'False
         Height          =   480
         Index           =   0
         Left            =   100
         Top             =   100
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Image imgRemesa 
         Enabled         =   0   'False
         Height          =   480
         Index           =   1
         Left            =   0
         Top             =   0
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Label Lbl 
         BackStyle       =   0  'Transparent
         Caption         =   "Cod.Gasto"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   0
         Left            =   615
         TabIndex        =   9
         Top             =   435
         Width           =   990
      End
      Begin VB.Label Lbl 
         BackStyle       =   0  'Transparent
         Caption         =   "Descripción del Gasto"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   1
         Left            =   2085
         TabIndex        =   8
         Top             =   435
         Width           =   2775
      End
      Begin VB.Label Lbl 
         BackStyle       =   0  'Transparent
         Caption         =   "Cod.Proveedor"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   2
         Left            =   615
         TabIndex        =   7
         Top             =   1260
         Width           =   1380
      End
      Begin VB.Label Lbl 
         BackStyle       =   0  'Transparent
         Caption         =   "Nombre / Razón Social"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   3
         Left            =   2085
         TabIndex        =   6
         Top             =   1260
         Width           =   1380
      End
   End
   Begin ComctlLib.ImageList ImageList1 
      Left            =   11340
      Top             =   675
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   20
      ImageHeight     =   20
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   12
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmGFMen.frx":047C
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmGFMen.frx":05FE
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmGFMen.frx":0780
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmGFMen.frx":0902
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmGFMen.frx":0A84
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmGFMen.frx":0C06
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmGFMen.frx":0D88
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmGFMen.frx":0F0A
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmGFMen.frx":108C
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmGFMen.frx":120E
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmGFMen.frx":1390
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmGFMen.frx":1512
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmGFMen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rstCatalogo As ADODB.Recordset
Attribute rstCatalogo.VB_VarHelpID = -1
Dim rstProveedor As ADODB.Recordset
Attribute rstProveedor.VB_VarHelpID = -1
Dim rstGastosM As ADODB.Recordset
Dim cnnLocal As ADODB.Connection
Dim WithEvents mView As ctlReport
Attribute mView.VB_VarHelpID = -1

Private Sub cmd_Click()
agregar_gastomenor
End Sub

Private Sub Dat_Click(Index As Integer, Area As Integer)
If Area = 2 Then
    Select Case Index
        Case 0, 1   'catálogo de gastos
            Dat(IIf(Index = 0, 1, 0)) = Dat(Index).BoundText
        Case 2, 3   'catálogo de proveedores
            Dat(IIf(Index = 2, 3, 2)) = Dat(Index).BoundText
    End Select
End If
End Sub

Private Sub Dat_KeyPress(Index As Integer, KeyAscii As Integer)
'convierte el caracter en mayúscula
KeyAscii = Asc(UCase(Chr(KeyAscii)))
If KeyAscii = 13 Then
    Select Case Index
        Case 0
            Dat(1) = Buscar(rstCatalogo, "CodGasto ='" & Dat(0) & "'", "Titulo", "CodGasto")
        Case 1
            Dat(0) = Buscar(rstCatalogo, "Titulo LIKE '*" & Dat(1) & "*'", "CodGasto", "Titulo", Dat(Index))
        Case 2
            Dat(3) = Buscar(rstProveedor, "Codigo='" & Dat(2) & "'", "NombProv", "Codigo")
        Case 3
            Dat(2) = Buscar(rstProveedor, "NombProv LIKE '*" & Dat(3) & "*'", "Codigo", "NombProv", Dat(Index))
    End Select
    'SendKeys vbTab
End If
End Sub

Private Sub Flex_Click()
If Flex.Col = Flex.Cols - 1 And Flex.Row > 0 Then
    If Flex.TextMatrix(Flex.RowSel, 0) <> "" Then
        Set Flex.CellPicture = IIf(Flex.CellPicture = imgRemesa(0), imgRemesa(1), imgRemesa(0))
        Flex.CellPictureAlignment = flexAlignCenterCenter
    End If
End If
End Sub

Private Sub Form_Load()
Dim i&
'creamos unas instancias de los ADODB.Recordsets
Set rstCatalogo = New ADODB.Recordset
Set rstProveedor = New ADODB.Recordset
Set rstGastosM = New ADODB.Recordset
Set cnnLocal = New ADODB.Connection
'
'configura el origen de los datos
rstCatalogo.CursorLocation = adUseClient
rstProveedor.CursorLocation = adUseClient
rstGastosM.CursorLocation = adUseClient
cnnLocal.Open cnnOLEDB + gcPath + gcUbica + "inm.mdb"
'
rstCatalogo.Open "SELECT * FROM TGastos WHERE Fijo=True ORDER BY CodGasto", cnnLocal, _
adOpenKeyset, adLockOptimistic, adCmdText
rstProveedor.Open "Proveedores", cnnConexion, adOpenKeyset, adLockOptimistic, adCmdTable
rstGastosM.Open "SELECT *,CodGasto & CodProv as Clave FROM GastosMenores ORDER BY CodGasto,CodProv", cnnLocal, _
adOpenKeyset, adLockOptimistic, adCmdText
rstProveedor.Sort = "Codigo"

Toolbar1.Buttons("PRINT").Enabled = Not (rstGastosM.EOF And rstGastosM.BOF)
'
'enlaza los datos a los controles
Dat(0).ListField = "Codgasto"
Dat(0).BoundColumn = "Titulo"
Dat(1).ListField = "Titulo"
Dat(1).BoundColumn = "CodGasto"
Dat(2).ListField = "Codigo"
Dat(2).BoundColumn = "NombProv"
Dat(3).ListField = "NombProv"
Dat(3).BoundColumn = "Codigo"
lbl(4).Visible = rstCatalogo.EOF And rstCatalogo.BOF
'
Set Dat(0).RowSource = rstCatalogo
Set Dat(1).RowSource = rstCatalogo
Set Dat(2).RowSource = rstProveedor
Set Dat(3).RowSource = rstProveedor
'
imgRemesa(0).Picture = LoadResPicture("Unchecked", vbResBitmap)
imgRemesa(1).Picture = LoadResPicture("Checked", vbResBitmap)
'
'configura la presentación de la rejilla
Flex.Redraw = False
Flex.TextArray(0) = "Código" & vbCrLf & "Gasto"
Flex.TextArray(1) = "Descripción" & vbCrLf & "Gasto"
Flex.TextArray(2) = "Código" & vbCrLf & "Proveedor"
Flex.TextArray(3) = "Nombre / Razón Social" & vbCrLf & "Proveedor"
Flex.RowHeight(0) = 1.2 * TextHeight(Flex.TextMatrix(0, 0))
Flex.ColAlignment(0) = flexAlignCenterCenter
Flex.ColAlignment(1) = flexAlignLeftCenter
Flex.ColAlignment(2) = flexAlignCenterCenter
Flex.ColAlignment(3) = flexAlignLeftCenter
Flex.Col = Flex.Cols - 1
Flex.Row = 1
Set Flex.CellPicture = imgRemesa(0)
Flex.CellPictureAlignment = flexAlignCenterCenter
Flex.Row = 0
Flex.Col = 0
Flex.RowSel = 0
Flex.ColSel = Flex.Cols - 1
Flex.FillStyle = flexFillRepeat
Flex.CellAlignment = flexAlignCenterCenter
Flex.FillStyle = flexFillSingle
muestra_gastosmenores
'
End Sub

Private Sub Form_Resize()
'configura la presentacion de los controles
If FrmAdmin.WindowState <> vbMinimized Then
    '
    fra.Top = 600
    fra.Left = 200
    fra.Height = ScaleHeight - 200 - fra.Top
    fra.Width = ScaleWidth - 400
    Flex.Width = fra.Width - (Flex.Left * 2)
    Flex.Height = fra.Height - Flex.Top - 400
    Flex.ColWidth(0) = 0.1 * Flex.Width
    Flex.ColWidth(1) = 0.4 * Flex.Width
    Flex.ColWidth(2) = 0.08 * Flex.Width
    Flex.ColWidth(3) = 0.32 * Flex.Width
    Flex.ColWidth(4) = 0.05 * Flex.Width
    '
End If
'
End Sub

Private Sub Form_Unload(Cancel As Integer)
'cierra y descarga de memoria los objetos
rstCatalogo.Close
Set rstCatalogo = Nothing
rstProveedor.Close
Set rstProveedor = Nothing
cnnLocal.Close
Set cnnLocal = Nothing
End Sub

Private Sub mView_error(ID As Long, Descripcion As String)
MsgBox Descripcion, vbCritical, "Error " & ID
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As ComctlLib.Button)

Select Case Button.Key
    Case "CLOSE"
        Unload Me
        Set frmGFMen = Nothing
    
    Case "SAVE"
        Toolbar1.Buttons("NEW").Enabled = True
        Toolbar1.Buttons("SAVE").Enabled = False
        Toolbar1.Buttons("EDIT").Enabled = False
        Toolbar1.Buttons("CLOSE").Enabled = True
        Toolbar1.Buttons("CANCEL").Enabled = False
        Toolbar1.Buttons("PRINT").Enabled = Not (rstGastosM.EOF And rstGastosM.BOF)
                
                
    Case "NEW"
        fra.Enabled = True
        Toolbar1.Buttons("NEW").Enabled = False
        Toolbar1.Buttons("SAVE").Enabled = True
        Toolbar1.Buttons("EDIT").Enabled = True
        Toolbar1.Buttons("CLOSE").Enabled = False
        Toolbar1.Buttons("PRINT").Enabled = False
        Toolbar1.Buttons("CANCEL").Enabled = True
        Toolbar1.Buttons("DELETE").Enabled = True
        'cnnLocal.BeginTrans
            
    Case "CANCEL"
        fra.Enabled = False
        'cnnLocal.RollbackTrans
        Dat(0) = ""
        Dat(1) = ""
        Dat(2) = ""
        Dat(3) = ""
        Toolbar1.Buttons("NEW").Enabled = True
        Toolbar1.Buttons("SAVE").Enabled = False
        Toolbar1.Buttons("EDIT").Enabled = False
        Toolbar1.Buttons("CLOSE").Enabled = True
        Toolbar1.Buttons("DELETE").Enabled = False
        Toolbar1.Buttons("CANCEL").Enabled = False
        Toolbar1.Buttons("PRINT").Enabled = Not (rstGastosM.EOF And rstGastosM.BOF)
        muestra_gastosmenores
       
    Case "DELETE"
        eliminar_gastosmenores
        Toolbar1.Buttons("PRINT").Enabled = Not (rstGastosM.EOF And rstGastosM.BOF)

    
    Case "PRINT"
        Printer_Report
        
End Select

End Sub

Private Sub agregar_gastomenor()
Dim strCadena As String
Dim intFila As Single

'valida los campos mínimos requerios
If Trim(Dat(0)) = "" Then
    strCadena = "- Introduzca el código del gasto." & vbCrLf
ElseIf Not Dat(0).MatchedWithList Then
    strCadena = "- El Código del gasto no corresponde a un marcado " _
    & vbCr & "como fijo en el catálogo de gasto." & vbCrLf
End If
If Trim(Dat(1)) = "" Then
    strCadena = strCadena & "- Falta la descripción del gasto." & _
    vbCrLf
End If
If Trim(Dat(2)) = "" Then
    strCadena = strCadena & "- Falta el código del proveedor." & vbCrLf
ElseIf Not Dat(2).MatchedWithList Then
    strCadena = strCadena & "- El código del proveedor que introdujo no" & _
    vbCrLf & "corresponde a algún proveedor registrado." & vbCrLf
End If
If Not Dat(3).MatchedWithList Then
    strCadena = strCadena & "- Verifique el nombre o razón social del proveedor." _
    & vbCrLf & "Si desea modificar el nombre del proveedor hágalo" & vbCrLf & _
    "en la ficha del proveedor"
End If
'verifica la correspondencia del código del proveedor con
'el nombre o razón social
With rstProveedor
    If Not (.EOF And .BOF) Then
        .MoveFirst
        .Find "Codigo ='" & Dat(2) & "'"
        If Not .EOF Then
            If Not Dat(3) = !NombProv Then
                Dat(2) = !Codigo
            End If
        End If
    End If
    .MoveFirst
End With
'veirifica que no exista duplicidad de la informacion
With rstGastosM
    If Not (.EOF And .BOF) Then
        .Filter = "CodGasto='" & Dat(0) & "'"
        If Not .EOF Then
            .MoveFirst
            Do
            
                If !codProv = Dat(2) Then
                    strCadena = strCadena & "- Imposible guardar este registro. Ya este proveedor" & _
                    "tiene una provisión por el mismo código." & vbCrLf
                    Exit Do
                End If
                .MoveNext
            Loop Until .EOF
        End If
        .Filter = 0
    End If
End With
If strCadena = "" Then
    'agrega el gasto a la rejilla
    Flex.Redraw = False
    If Flex.TextMatrix(Flex.Rows - 1, 0) <> "" Then Flex.AddItem ""
    intFila = Flex.Rows - 1
    'guarda la información en la tabla
    rstGastosM.AddNew
    rstGastosM("CodGasto") = Dat(0)
    rstGastosM("Descripcion") = Dat(1)
    rstGastosM("CodProv") = Dat(2)
    rstGastosM("Nombre") = Dat(3)
    rstGastosM("Usuario") = gcUsuario
    rstGastosM("Fecha") = Date
    rstGastosM.Update
    Flex.TextMatrix(intFila, 0) = Dat(0)
    Flex.TextMatrix(intFila, 1) = Dat(1)
    Flex.TextMatrix(intFila, 2) = Dat(2)
    Flex.TextMatrix(intFila, 3) = Dat(3)
    Flex.Row = intFila
    Flex.Col = Flex.Cols - 1
    Set Flex.CellPicture = imgRemesa(0)
    Flex.CellPictureAlignment = flexAlignCenterCenter
    Dat(0) = ""
    Dat(1) = ""
    Dat(2) = ""
    Dat(3) = ""
    Flex.Redraw = True
Else
    MsgBox "No se puede procesar la transacción." & vbCrLf & _
    "Tome nota de (los) siguiente(s) error(es):" & vbCrLf & _
    vbCrLf & strCadena, vbInformation, App.ProductName
End If

End Sub


Private Function Buscar(rst As ADODB.Recordset, Criterio As String, campo As String, _
Campo1 As String, Optional ByRef ctl) As String
'variables locales
'
If Not (rst.EOF And rst.BOF) Then
    rst.MoveFirst
    rst.Find Criterio
    If Not rst.EOF Then
        Buscar = rst.Fields(campo)
        If InStr(Criterio, "LIKE") > 0 Then ctl.Text = rst.Fields(Campo1)
    End If
End If
End Function


Private Sub muestra_gastosmenores()
'distribuye la información en la rejilla
Flex.Redraw = False
Flex.Rows = 2
Flex.Row = 1
Flex.Col = 0
Flex.ColSel = Flex.Cols - 1
Flex.FillStyle = flexFillRepeat
Flex.Text = ""
Flex.FillStyle = flexFillSingle
Flex.Col = Flex.Cols - 1
Flex.Row = 1
Set Flex.CellPicture = imgRemesa(0)

With rstGastosM
    '
    If Not (.EOF And .BOF) Then
        .MoveFirst
        Do
            DoEvents
            i = i + 1
            If i >= Flex.Rows Then Flex.AddItem ""
            Flex.TextMatrix(i, 0) = !codGasto
            Flex.TextMatrix(i, 1) = !Descripcion
            Flex.TextMatrix(i, 2) = !codProv
            Flex.TextMatrix(i, 3) = !Nombre
            Flex.Row = i
            Flex.Col = Flex.Cols - 1
            Set Flex.CellPicture = imgRemesa(0)
            Flex.CellPictureAlignment = flexAlignCenterCenter
            .MoveNext
        Loop Until .EOF
    End If
    '
End With
Flex.Redraw = True
'
End Sub

Private Sub eliminar_gastosmenores()
'variables locales
Dim i%

With Flex
    .Col = .Cols - 1
    
    For i = .Rows - 1 To 1 Step -1
        .Row = i
        If .CellPicture = imgRemesa(1) Then
            If Not (rstGastosM.EOF And rstGastosM.BOF) Then
                rstGastosM.MoveFirst
                rstGastosM.Find "Clave='" & .TextMatrix(i, 0) & .TextMatrix(i, 2) & "'"
                If Not rstGastosM.EOF Then rstGastosM.Delete
            End If
        End If
    Next
    muestra_gastosmenores
End With
End Sub

Private Sub Printer_Report()
'variables locales
Set mView = New ctlReport
mView.Reporte = gcReport & "fact_gfm.rpt"
mView.Salida = crPantalla
mView.OrigenDatos(0) = mcDatos
mView.TituloVentana = "Gastos Fijos Menores"
mView.Imprimir
Set mView = Nothing
End Sub
