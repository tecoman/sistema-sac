VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmList 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Facturas Recibidas"
   ClientHeight    =   3135
   ClientLeft      =   2355
   ClientTop       =   2910
   ClientWidth     =   9135
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3135
   ScaleWidth      =   9135
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txt 
      Height          =   345
      Left            =   4095
      TabIndex        =   3
      Top             =   2460
      Width           =   1215
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid gridList 
      Height          =   1965
      Left            =   180
      TabIndex        =   2
      Tag             =   "0|0|1000|5000|1200|1000"
      Top             =   180
      Width           =   8685
      _ExtentX        =   15319
      _ExtentY        =   3466
      _Version        =   393216
      Cols            =   4
      FixedCols       =   0
      BackColorFixed  =   -2147483646
      ForeColorFixed  =   -2147483639
      FormatString    =   "NDoc| Descripción|Total|Estatus"
      _NumberOfBands  =   1
      _Band(0).Cols   =   4
   End
   Begin VB.CommandButton cmdLista 
      Cancel          =   -1  'True
      Caption         =   "Cancelar"
      Height          =   495
      Index           =   1
      Left            =   7650
      TabIndex        =   1
      Top             =   2415
      Width           =   1215
   End
   Begin VB.CommandButton cmdLista 
      Caption         =   "Aceptar"
      Height          =   495
      Index           =   0
      Left            =   6300
      TabIndex        =   0
      Top             =   2430
      Width           =   1215
   End
   Begin VB.Label lbl 
      Caption         =   "Ingrese el mes al que corresponde la factura: (MM/YYYY) ejm: 05/2003"
      Height          =   495
      Left            =   240
      TabIndex        =   4
      Top             =   2475
      Width           =   3795
   End
End
Attribute VB_Name = "frmList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Proveedor$, Inm$  'código del proveedor
Public Mantenimiento As Boolean, MontoFactura@
Dim rstCpp As New ADODB.Recordset

Private Sub cmdLista_Click(Index As Integer)
FrmFactura.Tag = Index
Select Case Index
    Case 0  'aceptar
        If Not Mantenimiento Then Unload Me: Exit Sub
        If Txt <> "" Then
            FrmFactura.Descrip = Txt
            Unload Me
        Else
            MsgBox "Presione cancelar para cerrar esta ventana.", vbInformation, App.ProductName
        End If
        
    Case 1: Unload Me
    '
End Select
End Sub

Private Sub Form_Load()
'Variables locales
Dim strEncabezado$
Dim strSQL$
'
If Mantenimiento Then
    strSQL = "SELECT CodProv,Frecep,Ndoc,Detalle,Format(Total,'#,##0.00') as T,Estatus FROM Cpp" _
    & " WHERE Codinm='" & Inm & "' AND (Detalle LIKE '%MANT%' or Detalle LIKE '%MTT%');"
    strEncabezado = "||^NDoc|Descripción|>Total|^Estatus"
Else
    lbl.Visible = False
    Txt.Visible = False
    strSQL = "SELECT CodProv,Frecep,Ndoc,Detalle,Format(Total,'#,##0.00') as T,Estatus FROM Cpp WHERE Cod" _
    & "Inm='" & Inm & "' AND Monto=" & MontoFactura & ";"
    gridList.Tag = "0|0|1000|5000|1200|1000"
    strEncabezado = "||^NDoc|Descripción|>Total|^Estatus"
    gridList.Refresh
End If
'
With rstCpp
    '
    .Open strSQL, cnnConexion, adOpenStatic, adLockOptimistic, adCmdText
    .Filter = "CodProv='" & Proveedor & "'"
    .Sort = "Frecep"
    Set gridList.DataSource = rstCpp
    gridList.FormatString = strEncabezado
    Call centra_titulo(gridList, True)
'
End With
'
End Sub


Private Sub Form_Unload(Cancel As Integer)
'cierra y descarga los objetos de memoria
rstCpp.Close
Set rstCpp = Nothing
Set frmList = Nothing
End Sub


Private Sub txt_KeyPress(KeyAscii As Integer)
Call Validacion(KeyAscii, "0123456789/")
If KeyAscii = 13 And Txt <> "" Then cmdLista_Click (0)
End Sub
