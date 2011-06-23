VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmOrdPag 
   Caption         =   "Emisión Orden de Pago"
   ClientHeight    =   6495
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5385
   Icon            =   "frmOrdPag.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6495
   ScaleWidth      =   5385
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "Salida:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   990
      Left            =   240
      TabIndex        =   16
      Top             =   5340
      Width           =   3165
      Begin VB.OptionButton Option1 
         Caption         =   "Impresora"
         Height          =   315
         Index           =   1
         Left            =   1215
         TabIndex        =   18
         Top             =   615
         Value           =   -1  'True
         Width           =   1215
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Ventana"
         Height          =   315
         Index           =   0
         Left            =   825
         TabIndex        =   17
         Top             =   255
         Width           =   1215
      End
   End
   Begin VB.CommandButton cmd 
      Caption         =   "&Imprimir"
      Height          =   495
      Index           =   1
      Left            =   3930
      TabIndex        =   14
      Top             =   5910
      Width           =   1215
   End
   Begin VB.CommandButton cmd 
      Cancel          =   -1  'True
      Caption         =   "&Cerrar"
      Height          =   495
      Index           =   0
      Left            =   3930
      TabIndex        =   15
      Top             =   5340
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   1215
      Index           =   2
      Left            =   240
      MaxLength       =   250
      MultiLine       =   -1  'True
      TabIndex        =   12
      Top             =   3930
      Width           =   4905
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   1
      Left            =   1650
      TabIndex        =   10
      Text            =   "0,00"
      Top             =   3165
      Width           =   1605
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Index           =   0
      Left            =   1665
      TabIndex        =   8
      Top             =   2685
      Width           =   3495
   End
   Begin MSDataListLib.DataCombo dtc 
      Height          =   315
      Index           =   0
      Left            =   1665
      TabIndex        =   4
      Top             =   1365
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   556
      _Version        =   393216
      Text            =   ""
   End
   Begin VB.ComboBox cmbDpto 
      Height          =   315
      ItemData        =   "frmOrdPag.frx":1782
      Left            =   1695
      List            =   "frmOrdPag.frx":17A4
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   270
      Width           =   2865
   End
   Begin MSDataListLib.DataCombo dtc 
      Height          =   315
      Index           =   1
      Left            =   1665
      TabIndex        =   6
      Top             =   1815
      Width           =   3525
      _ExtentX        =   6218
      _ExtentY        =   556
      _Version        =   393216
      Text            =   ""
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
   Begin VB.Label lbl 
      Caption         =   "(Máx 250 car.)"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   7
      Left            =   1215
      TabIndex        =   13
      Top             =   3660
      Width           =   1215
   End
   Begin VB.Label lbl 
      Caption         =   "Información del Inmueble"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   6
      Left            =   225
      TabIndex        =   2
      Top             =   945
      Width           =   2625
   End
   Begin VB.Label lbl 
      Caption         =   "Concepto:"
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
      Index           =   5
      Left            =   270
      TabIndex        =   11
      Top             =   3645
      Width           =   1215
   End
   Begin VB.Label lbl 
      Caption         =   "Monto:"
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
      Index           =   4
      Left            =   270
      TabIndex        =   9
      Top             =   3195
      Width           =   1215
   End
   Begin VB.Label lbl 
      Caption         =   "Beneficiario:"
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
      Left            =   285
      TabIndex        =   7
      Top             =   2715
      Width           =   1215
   End
   Begin VB.Label lbl 
      Caption         =   "Nombre:"
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
      Left            =   570
      TabIndex        =   5
      Top             =   1830
      Width           =   1005
   End
   Begin VB.Label lbl 
      Caption         =   "Codigo:"
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
      Left            =   570
      TabIndex        =   3
      Top             =   1395
      Width           =   1005
   End
   Begin VB.Label lbl 
      Caption         =   "Departamento:"
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
      Left            =   270
      TabIndex        =   0
      Top             =   285
      Width           =   1305
   End
End
Attribute VB_Name = "frmOrdPag"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmbDpto_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys vbTab
End Sub

Private Sub cmd_Click(Index As Integer)
'VAIRABLES LOCLES
Dim rpReporte As ctlReport, strSQL As String, NDoc As String
Dim toLetras As New clsNum2Let

If Index = 0 Then
    Unload Me
    Set frmOrdPag = Nothing
Else

    If Not faltan_datos Then
        
        NDoc = Registro_Proveedor(Text1(0), True, "OP S/N", Text1(2), Text1(1), dtc(0))
        If NDoc <> "" Then
            toLetras.Moneda = "Bs."
            toLetras.Numero = CCur(Text1(1))
            Set rpReporte = New ctlReport
            With rpReporte
                    .Reporte = gcReport + "cxp_OrdPag.rpt"
                    .Formulas(0) = "sDepartamento='" & cmbDpto & "'"
                    .Formulas(1) = "sUsuario='" & gcUsuario & "'"
                    .Formulas(2) = "sCodInm='" & dtc(0) & "'"
                    .Formulas(3) = "sNombreInm='" & dtc(1) & "'"
                    .Formulas(4) = "sBeneficiario='" & Text1(0) & "'"
                    .Formulas(5) = "nMonto=" & Replace(CStr(CCur(Text1(1))), ",", ".")
                    .Formulas(6) = "sConcepto1='" & Replace(Text1(2), Chr(13), Chr(0)) & "'"
                    .Formulas(7) = "sConcepto2=''"
                    .Formulas(8) = "sMonto='" & toLetras.ALetra & "'"
                    If Option1(0) Then
                    
                        .Salida = crPantalla
                        .TituloVentana = "Orden de Pago"
                        
                    Else
                        .Salida = crImpresora
                    End If
                    .Imprimir
            End With
            Set rpReporte = Nothing
        Else
            MsgBox "Ocurrio un error durante el proceso. Consulte al administrador del sistema", _
            vbInformation, App.ProductName
        End If
    End If
End If
End Sub

Private Sub dtc_Click(Index As Integer, Area As Integer)
If Area = 2 Then
    dtc(IIf(Index = 0, 1, 0)) = dtc(Index).BoundText
    Text1(0).SetFocus
End If
End Sub

Private Sub dtc_KeyPress(Index As Integer, KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
If KeyAscii = 13 Then
    If Index = 0 Then
        Call dtc_Click(0, 2)
        Text1(0).SetFocus
    Else
        With FrmAdmin.ObjRstNom
        
            If Not .EOF And Not .BOF Then
                .MoveFirst
                .Find "Nombre LIKE '*" & dtc(1) & "*'"
                If Not .EOF Then
                    dtc(1) = !Nombre
                    Call dtc_Click(1, 2)
                    Text1(0).SetFocus
                Else
                    .MoveFirst
                End If
                '
            End If
        End With
    End If
    
End If
'
End Sub

Private Sub Form_Load()
'variables locales
Set dtc(0).RowSource = FrmAdmin.objRst
Set dtc(1).RowSource = FrmAdmin.ObjRstNom
'
dtc(0).ListField = "CodInm"
dtc(0).BoundColumn = "Nombre"
dtc(1).ListField = "Nombre"
dtc(1).BoundColumn = "CodInm"
'
End Sub

Private Sub Text1_GotFocus(Index As Integer)
'
If Index = 1 Then
    If IsNumeric(Text1(1)) Then Text1(1) = CCur(Text1(1))
    Text1(1).SelStart = 0
    Text1(1).SelLength = Len(Text1(1))
End If
'
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
If Index = 1 Then
    If KeyAscii = Asc(".") Then KeyAscii = Asc(",")
    Call Validacion(KeyAscii, "0123456789,")
End If
If KeyAscii = 13 Then SendKeys vbTab
If Index = 2 And KeyAscii = 13 Then KeyAscii = 0
End Sub

Private Function faltan_datos() As Boolean
'variables locales
Dim Cadena As String
'
If cmbDpto.Text = "" Then Cadena = "- Indique el Dpto. solicitante. " & vbCrLf
If Not dtc(0).MatchedWithList Then Cadena = Cadena & _
"- El Código del Inmueble no corresponde con un elemento de la lista." & vbCrLf
If Not dtc(1).MatchedWithList Then Cadena = Cadena & _
"- El Nombre del Inmueble no corresponde con un elemento de la lista." & vbCrLf
If Text1(0) = "" Then Cadena = Cadena & "- Falta datos del proveedor." & vbCrLf
If Text1(1) = "" Then Text1(1) = "0,00"
If Not Text1(1) > 0 Then Cadena = Cadena & "- Revise el monto de la orden de pago." & vbCrLf
If Text1(2) = "" Then Cadena = Cadena & "- Indique el concepto para emitir la orden de pago."
If Cadena <> "" Then
    Cadena = "Imposible emitir la orden de pago." & vbCrLf & "Revise el(los) siguiente(s) error(es): " _
    & vbCrLf & vbCrLf & Cadena
    faltan_datos = MsgBox(Cadena, vbCritical, App.ProductName)
End If
'
End Function

Private Sub Text1_LostFocus(Index As Integer)
If Index = 1 Then If Text1(1) <> "" Then Text1(1) = Format(Text1(1), "#,##0.00")
End Sub
