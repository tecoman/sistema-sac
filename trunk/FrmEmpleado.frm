VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Begin VB.Form frmGestion 
   Caption         =   "Gestiones de Cobro"
   ClientHeight    =   9645
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11955
   ControlBox      =   0   'False
   Icon            =   "FrmEmpleado.frx":0000
   LinkTopic       =   "Form3"
   MDIChild        =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   9645
   ScaleWidth      =   11955
   Tag             =   "1"
   WindowState     =   2  'Maximized
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid gridGestion 
      Height          =   3750
      Left            =   495
      TabIndex        =   30
      Top             =   4110
      Width           =   10845
      _ExtentX        =   19129
      _ExtentY        =   6615
      _Version        =   393216
      Cols            =   6
      FixedCols       =   3
      BackColorFixed  =   -2147483646
      ForeColorFixed  =   -2147483639
      BackColorSel    =   65280
      BackColorBkg    =   -2147483636
      AllowBigSelection=   0   'False
      AllowUserResizing=   1
      FormatString    =   "Fecha |Hora |Apto. |Teléfono |Contacto |Resultado"
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
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   6
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   195
      Top             =   4365
   End
   Begin VB.Frame fraGestion 
      Caption         =   "Gestión de Cobro"
      Height          =   2580
      Index           =   1
      Left            =   4305
      TabIndex        =   22
      Top             =   1110
      Width           =   6990
      Begin VB.TextBox txtGestion 
         Height          =   315
         Index           =   11
         Left            =   720
         Locked          =   -1  'True
         TabIndex        =   28
         Top             =   2100
         Width           =   6180
      End
      Begin VB.TextBox txtGestion 
         Height          =   730
         Index           =   10
         Left            =   1050
         MaxLength       =   240
         MultiLine       =   -1  'True
         TabIndex        =   5
         Top             =   1245
         Width           =   5850
      End
      Begin VB.TextBox txtGestion 
         Height          =   315
         Index           =   9
         Left            =   1590
         TabIndex        =   3
         Top             =   810
         Width           =   5325
      End
      Begin VB.TextBox txtGestion 
         Height          =   315
         Index           =   7
         Left            =   2670
         Locked          =   -1  'True
         TabIndex        =   26
         Top             =   375
         Width           =   1215
      End
      Begin VB.TextBox txtGestion 
         Height          =   315
         Index           =   6
         Left            =   720
         Locked          =   -1  'True
         TabIndex        =   24
         Top             =   375
         Width           =   1215
      End
      Begin MSMask.MaskEdBox MskTelefono 
         Height          =   315
         Left            =   5325
         TabIndex        =   1
         Top             =   375
         Width           =   1545
         _ExtentX        =   2725
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   16
         Format          =   "(####)-###-##-##"
         Mask            =   "(####)-###-##-##"
         PromptChar      =   "_"
      End
      Begin VB.Label lblGestion 
         Caption         =   "&Resultado:"
         Height          =   285
         Index           =   11
         Left            =   165
         TabIndex        =   4
         Top             =   1245
         Width           =   885
      End
      Begin VB.Label lblGestion 
         Caption         =   "Nota:"
         Height          =   240
         Index           =   10
         Left            =   165
         TabIndex        =   27
         Top             =   2137
         Width           =   885
      End
      Begin VB.Label lblGestion 
         Caption         =   "&Persona Contacto:"
         Height          =   285
         Index           =   9
         Left            =   165
         TabIndex        =   2
         Top             =   825
         Width           =   1350
      End
      Begin VB.Label lblGestion 
         Caption         =   "&Número Discado:"
         Height          =   285
         Index           =   8
         Left            =   3975
         TabIndex        =   0
         Top             =   390
         Width           =   1350
      End
      Begin VB.Label lblGestion 
         Caption         =   "Hora:"
         Height          =   285
         Index           =   7
         Left            =   2190
         TabIndex        =   25
         Top             =   390
         Width           =   840
      End
      Begin VB.Label lblGestion 
         Caption         =   "Fecha:"
         Height          =   285
         Index           =   6
         Left            =   165
         TabIndex        =   23
         Top             =   390
         Width           =   840
      End
   End
   Begin VB.Frame fraGestion 
      Caption         =   "Datos Propietario:"
      Height          =   2580
      Index           =   0
      Left            =   495
      TabIndex        =   9
      Top             =   1110
      Width           =   3270
      Begin VB.TextBox txtGestion 
         Height          =   315
         Index           =   5
         Left            =   1020
         Locked          =   -1  'True
         TabIndex        =   21
         Top             =   1950
         Width           =   1980
      End
      Begin VB.TextBox txtGestion 
         Height          =   315
         Index           =   4
         Left            =   1020
         Locked          =   -1  'True
         TabIndex        =   19
         Top             =   1635
         Width           =   1980
      End
      Begin VB.TextBox txtGestion 
         Height          =   315
         Index           =   3
         Left            =   1020
         Locked          =   -1  'True
         TabIndex        =   17
         Top             =   1320
         Width           =   1980
      End
      Begin VB.TextBox txtGestion 
         Height          =   315
         Index           =   2
         Left            =   1020
         Locked          =   -1  'True
         TabIndex        =   15
         Top             =   1005
         Width           =   1980
      End
      Begin VB.TextBox txtGestion 
         Height          =   315
         Index           =   1
         Left            =   1020
         Locked          =   -1  'True
         TabIndex        =   13
         Top             =   690
         Width           =   750
      End
      Begin VB.TextBox txtGestion 
         Height          =   315
         Index           =   0
         Left            =   1020
         Locked          =   -1  'True
         TabIndex        =   11
         Top             =   375
         Width           =   1980
      End
      Begin VB.Label lblGestion 
         Caption         =   "Fax:"
         Height          =   285
         Index           =   5
         Left            =   210
         TabIndex        =   20
         Top             =   1650
         Width           =   840
      End
      Begin VB.Label lblGestion 
         Caption         =   "e-mail:"
         Height          =   285
         Index           =   4
         Left            =   210
         TabIndex        =   18
         Top             =   1965
         Width           =   840
      End
      Begin VB.Label lblGestion 
         Caption         =   "Celular:"
         Height          =   285
         Index           =   3
         Left            =   210
         TabIndex        =   16
         Top             =   1335
         Width           =   840
      End
      Begin VB.Label lblGestion 
         Caption         =   "Telf. Hab.:"
         Height          =   285
         Index           =   2
         Left            =   210
         TabIndex        =   14
         Top             =   1020
         Width           =   840
      End
      Begin VB.Label lblGestion 
         Caption         =   "Ext.:"
         Height          =   285
         Index           =   1
         Left            =   210
         TabIndex        =   12
         Top             =   705
         Width           =   840
      End
      Begin VB.Label lblGestion 
         Caption         =   "Telf. Ofc.:"
         Height          =   285
         Index           =   0
         Left            =   210
         TabIndex        =   10
         Top             =   390
         Width           =   840
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   300
      Top             =   975
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   20
      ImageHeight     =   20
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   11
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmEmpleado.frx":0442
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmEmpleado.frx":0556
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmEmpleado.frx":0676
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmEmpleado.frx":078E
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmEmpleado.frx":08A6
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmEmpleado.frx":09BE
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmEmpleado.frx":0B42
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmEmpleado.frx":0C5A
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmEmpleado.frx":0DDE
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmEmpleado.frx":18AA
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmEmpleado.frx":1CFE
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin ComCtl3.CoolBar clbGestion 
      Align           =   1  'Align Top
      Height          =   870
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   11955
      _ExtentX        =   21087
      _ExtentY        =   1535
      BandCount       =   5
      VariantHeight   =   0   'False
      _CBWidth        =   11955
      _CBHeight       =   870
      _Version        =   "6.0.8169"
      BandBackColor1  =   -2147483639
      Child1          =   "Toolbar1"
      MinHeight1      =   390
      Width1          =   8490
      NewRow1         =   0   'False
      BandBackColor2  =   -2147483643
      Caption2        =   "Rec. Pend.:"
      Child2          =   "Text1"
      MinWidth2       =   405
      MinHeight2      =   390
      Width2          =   8880
      NewRow2         =   0   'False
      BandEmbossPicture2=   -1  'True
      Caption3        =   "Mes:"
      Child3          =   "Combo1 (0)"
      MinWidth3       =   1395
      MinHeight3      =   315
      Width3          =   1395
      NewRow3         =   0   'False
      Caption4        =   "Año:"
      Child4          =   "Combo1 (1)"
      MinWidth4       =   1200
      MinHeight4      =   315
      Width4          =   1200
      NewRow4         =   0   'False
      Caption5        =   "Propietarios:"
      Child5          =   "ImageCombo1"
      MinHeight5      =   330
      Width5          =   1230
      NewRow5         =   -1  'True
      Begin VB.ComboBox Combo1 
         Height          =   315
         Index           =   1
         ItemData        =   "FrmEmpleado.frx":2152
         Left            =   10665
         List            =   "FrmEmpleado.frx":2154
         Style           =   2  'Dropdown List
         TabIndex        =   32
         Top             =   60
         Width           =   1200
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Index           =   0
         ItemData        =   "FrmEmpleado.frx":2156
         Left            =   8610
         List            =   "FrmEmpleado.frx":2184
         Style           =   2  'Dropdown List
         TabIndex        =   31
         Top             =   60
         Width           =   1395
      End
      Begin MSComctlLib.ImageCombo ImageCombo1 
         Height          =   330
         Left            =   1215
         TabIndex        =   29
         Top             =   480
         Width           =   10650
         _ExtentX        =   18785
         _ExtentY        =   582
         _Version        =   393216
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ImageList       =   "ImageList1"
      End
      Begin VB.TextBox Text1 
         Height          =   390
         Left            =   7545
         TabIndex        =   8
         Text            =   "5"
         Top             =   30
         Width           =   405
      End
      Begin MSComctlLib.Toolbar Toolbar1 
         Height          =   570
         Left            =   165
         TabIndex        =   7
         Top             =   30
         Width           =   6225
         _ExtentX        =   10980
         _ExtentY        =   1005
         ButtonWidth     =   714
         ButtonHeight    =   688
         Style           =   1
         ImageList       =   "ImageList1"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   8
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Enabled         =   0   'False
               Key             =   "New"
               ImageIndex      =   1
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Save"
               Object.ToolTipText     =   "Guardar"
               ImageIndex      =   2
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Find"
               Object.ToolTipText     =   "Diario"
               ImageIndex      =   3
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Undo"
               ImageIndex      =   4
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Enabled         =   0   'False
               Key             =   "Delete"
               ImageIndex      =   5
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Edit"
               ImageIndex      =   6
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Print"
               Object.ToolTipText     =   "Imprimir"
               ImageIndex      =   7
               Style           =   5
               BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
                  NumButtonMenus  =   3
                  BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "Todo"
                     Text            =   "General"
                  EndProperty
                  BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "Propietario"
                     Text            =   "Propietario"
                  EndProperty
                  BeginProperty ButtonMenu3 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "Diario"
                     Text            =   "Diario"
                  EndProperty
               EndProperty
            EndProperty
            BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Close"
               Object.ToolTipText     =   "Cerrar ventana"
               ImageIndex      =   8
            EndProperty
         EndProperty
      End
   End
End
Attribute VB_Name = "frmGestion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Const Mascara$ = "(____)-___-__-__"
Dim strQDF As String, strFecha As String, timHora As String
Dim apto As String, Dpropietario() As String
Dim Modo As Mode
Enum Mode
    moAdd
    moEdit
End Enum


Private Sub Combo1_Click(Index%)
'rem
If Combo1(0) <> "" And Combo1(1) <> "" Then Call Vista
End Sub

Private Sub Form_Load()
For i = 0 To 2
    Combo1(1).AddItem Year(Date) - i, i
Next
Combo1(1) = Combo1(1).List(0)
Combo1(0) = Combo1(0).List(Month(Date) - 1)
'Call config_grid
Call Listar
With clbGestion
'
    For i = 1 To .Bands.Count - 1: .Bands(i).MinHeight = 315
    Next
End With
'
End Sub

Private Sub Form_Resize()
On Error Resume Next

gridGestion.Height = Me.Height - gridGestion.Top - 200 - Me.clbGestion.Height
gridGestion.Width = Me.Width - (gridGestion.Left * 2)
Call Config_Grid

End Sub


Private Sub gridGestion_Click()
'variables locales
'
With gridGestion
    '
    If .Rows > 0 Then
        '
        Fila = .RowSel
        If .TextMatrix(Fila, 3) <> "" Then
            MskTelefono = .TextMatrix(Fila, 3)
        End If
        txtGestion(9) = .TextMatrix(Fila, 4)
        txtGestion(10) = .TextMatrix(Fila, 5)
        strFecha = Format(.TextMatrix(Fila, 0), "MM/DD/YYYY")
        timHora = .TextMatrix(Fila, 1)
        apto = .TextMatrix(Fila, 2)
        Toolbar1.Buttons("Delete").Enabled = True
        '
    End If
    '
End With
'
End Sub

Private Sub ImageCombo1_Click()
'variables locales
Dim Pos As Integer
Dim Prop As String
'
If ImageCombo1.SelectedItem.Indentation = 0 Then

10    txtGestion(0) = ""
    txtGestion(1) = ""
    txtGestion(2) = ""
    txtGestion(3) = ""
    txtGestion(4) = ""
    txtGestion(5) = ""
    txtGestion(11) = ""
    Call Vista
    Exit Sub
End If
Pos = InStr(ImageCombo1.Text, " ")
If Pos = 0 Then GoTo 10
Prop = Left(ImageCombo1.Text, Pos - 1)
apto = Prop
For Pos = 0 To UBound(Dpropietario)

    If Prop = Dpropietario(Pos, 0) Then
    
        txtGestion(0) = Dpropietario(Pos, 1)
        txtGestion(1) = Dpropietario(Pos, 2)
        txtGestion(2) = Dpropietario(Pos, 3)
        txtGestion(3) = Dpropietario(Pos, 4)
        txtGestion(4) = Dpropietario(Pos, 5)
        txtGestion(5) = Dpropietario(Pos, 6)
        txtGestion(11) = Dpropietario(Pos, 7)
        Exit For
        
    End If
Next
Call Vista
End Sub

Private Sub ImageCombo1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 46 Then KeyCode = 0
End Sub

Private Sub ImageCombo1_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub

Private Sub MskTelefono_GotFocus()
With MskTelefono
    .SelStart = 0
    .SelLength = Len(.Text)
End With
End Sub

Private Sub MskTelefono_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys vbTab
Call Validacion(KeyAscii, "0123456789")
End Sub

Private Sub Text1_Change()
Call Listar
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
Call Validacion(KeyAscii, "0123456789")
End Sub

Private Sub Timer1_Timer()
'txtGestion(6) = Date
'txtGestion(7) = Time
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
'variables locales
Dim Fila As Integer
'
Select Case Button.Key

    Case "New": Modo = moAdd
        
    Case "Save"
        Call guardar_Registro
        
    Case "Find"
        ImageCombo1.Text = ""
        Call Cancelar
        Call Vista("SELECT * FROM Gestion WHERE Fecha=Date();")
        
    Case "Undo"
        Call Cancelar
        
    Case "Delete"
        Call eliminar_registro
        Toolbar1.Buttons("Delete").Enabled = False
    Case "Edit"
        With gridGestion
            Fila = .RowSel
            MskTelefono = .TextMatrix(Fila, 3)
            txtGestion(9) = .TextMatrix(Fila, 4)
            txtGestion(10) = .TextMatrix(Fila, 5)
            strFecha = Format(.TextMatrix(Fila, 0), "MM/DD/YYYY")
            timHora = Left(.TextMatrix(Fila, 1), 8)
            apto = .TextMatrix(Fila, 2)
        End With
        Modo = moEdit
        
    Case "Print"
        Call rtnGenerator(mcDatos, strQDF, "QdfGestion")
        mcTitulo = "Gestión de Cobro"
        mcReport = "Gestion.Rpt"
        FrmReport.Show
        'Call Muestra_Formulario(FrmReport, "Impresióm " & mcTitulo)
        
    Case "Close"
        Unload Me
End Select
'
End Sub

    
    Private Sub Listar()
    'variables locales
    Dim strSQL As String
    Dim Elemento As ComboItem
    Dim rstPropietario As New ADODB.Recordset
    Dim cnnPropietario As New ADODB.Connection
    Dim i As Integer
    '
    If Text1 = "" Or IsNull(Text1) Then Exit Sub
    cnnPropietario.Open cnnOLEDB & mcDatos
    strSQL = "SELECT * FROM PROPIETARIOS WHERE Recibos>=" & Text1
    rstPropietario.Open strSQL, cnnPropietario, adOpenStatic, adLockReadOnly, adCmdText
    
    With rstPropietario
        If Not .EOF Or Not .BOF Then
            ReDim Dpropietario(.RecordCount, 7)
            .MoveFirst: i = 0
            ImageCombo1.ComboItems.Clear
            Set Elemento = ImageCombo1.ComboItems.Add(, , gcCodInm & " " & gcNomInm, 9, 9, 0)
            Do
                'Llena una matriz con la información de c/propietario
                Dpropietario(i, 0) = IIf(IsNull(!Codigo), "", !Codigo)
                Dpropietario(i, 1) = IIf(IsNull(!telefonos), "", !telefonos)
                Dpropietario(i, 2) = IIf(IsNull(!ExtOfc), "", !ExtOfc)
                Dpropietario(i, 3) = IIf(IsNull(!TelfHab), "", !TelfHab)
                Dpropietario(i, 4) = IIf(IsNull(!Celular), "", !Celular)
                Dpropietario(i, 5) = IIf(IsNull(!Fax), "", !Fax)
                Dpropietario(i, 6) = IIf(IsNull(!email), "", !email)
                Dpropietario(i, 7) = IIf(IsNull(!Notas), "", !Notas)
                strSQL = !Codigo & " " & !Nombre & " Rec.Pen.:(" & !Recibos & ")" 'Deuda: " & _
                Format(!Deuda, "#,##0.00")
                
                Set Elemento = ImageCombo1.ComboItems.Add(, , strSQL, 10, 11, 1)
                .MoveNext: i = i + 1
            Loop Until .EOF
            ImageCombo1.SelectedItem = ImageCombo1.ComboItems(1)
           .Close
        End If
    End With
    Call Vista
    '
    Set rstPropietario = Nothing
    cnnPropietario.Close
    Set cnnPropietario = Nothing
    
    End Sub

    Private Sub Toolbar1_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
    Select Case ButtonMenu.Key
        Case "Todo"
        Case "Propietario"
    End Select
    End Sub

    Private Sub txtGestion_DblClick(Index As Integer)
    '
    Select Case Index
    
        Case 0, 2, 3, 4
        
            MskTelefono.PromptInclude = False
            MskTelefono.Text = txtGestion(Index).Text
            MskTelefono.PromptInclude = True
            txtGestion(9).SetFocus
            
    End Select
    '
    End Sub

    Private Sub txtGestion_KeyPress(Index%, KeyAscii%)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    End Sub

    Private Sub guardar_Registro()
    Dim ObjCnn As New ADODB.Connection  'Variables locales
    Dim strSQL As String
    Dim Fila As Integer
    
    ObjCnn.Open cnnOLEDB + mcDatos
    
    If Modo = moAdd Then
        If apto = "" Then
            strSQL = "Falta Codigo del apartamento.."
        ElseIf MskTelefono = Mascara Then
           strSQL = "Falta Número de Teléfono"
        ElseIf txtGestion(9) = "" Then
            strSQL = "Falta Nombre de la persona contacto"
        ElseIf txtGestion(10) = "" Then
            strSQL = "Escriba el resultado de la gestión"
        End If
        If Not strSQL = "" Then
            MsgBox strSQL, vbInformation, App.ProductName
            Exit Sub
        End If
        
        ObjCnn.BeginTrans
        strSQL = "INSERT INTO Gestion (Apto,Telf,Contacto,Resultado,Fecha,Hora,Usuario) VALUES " _
        & "('" & apto & "','" & MskTelefono & "','" & txtGestion(9) & "','" & txtGestion(10) & _
        "',Date(),Time(),'" & gcUsuario & "');"
        '
        ObjCnn.Execute strSQL
        Call rtnBitacora("Guardar Gestión de Cobro " & gcCodInm & "/" & apto)
        If Err.Number = 0 Then
            ObjCnn.CommitTrans
            With gridGestion
                Fila = .Rows - 1
                If .TextMatrix(Fila, 0) <> "" And .TextMatrix(Fila, 1) <> "" And .TextMatrix(Fila, 2) <> "" Then
                    .AddItem (Date)
                    Fila = Fila + 1
                Else
                    .TextMatrix(Fila, 0) = Date
                End If
                .TextMatrix(Fila, 1) = Time
                .TextMatrix(Fila, 2) = apto
                .TextMatrix(Fila, 3) = MskTelefono
                .TextMatrix(Fila, 4) = txtGestion(9)
                .TextMatrix(Fila, 5) = txtGestion(10)
            End With
            Call Cancelar
            MsgBox "Regristro Guardado", vbInformation, App.ProductName
        Else
            ObjCnn.RollbackTrans
            MsgBox Err.Description, vbCritical, Err.Number
        End If
    Else
        ObjCnn.Execute "UPDATE Gestion SET Resultado='" & txtGestion(10) & "',Contacto='" & _
        txtGestion(9) & "',Telf='" & MskTelefono & "' WHERE Fecha=#" & strFecha & "# AND Hora=#" _
        & timHora & "# AND Apto='" & apto & "';"
        Call rtnBitacora("Editar Gestión de Cobro " & gcCodInm & "/" & apto)
        MsgBox "Registro Actualizado..", vbInformation, App.ProductName
        Call Cancelar
        Call Vista
    End If
    ObjCnn.Close
    Set ObjCnn = Nothing
    
    End Sub

'----------------------------------------------------------------------------------------
'   Rutina: Vista
'
'   Muestra las en el grid las gestiones de cobro efectuadas en determinado
'   inmueble y/o determinado propietario
'----------------------------------------------------------------------------------------
Private Sub Vista(Optional Diario$)
'variables locales
Dim objRst As New ADODB.Recordset
Dim ObjCnn As New ADODB.Connection
Dim strSQL As String
Dim Periodo1, Periodo2
'
ObjCnn.Open cnnOLEDB + mcDatos
If Diario = "" Then
    If Combo1(0) = "TODO" Then
        Periodo1 = "01/01/1975"
        Periodo2 = Date
    Else
        Periodo1 = "01/" & Combo1(0) & "/" & Combo1(1)
        Periodo2 = DateAdd("m", 1, Periodo1)
        Periodo2 = DateAdd("d", -1, Periodo2)
    End If
    With FrmReport
        .apto = IIf(apto <> "", ImageCombo1.Text, "")
        .Desde = Format(Periodo1, "YYYY,MM,DD")
        .Hasta = Format(Periodo2, "YYYY,MM,DD")
    End With
    Periodo1 = Format(Periodo1, "mm/dd/yyyy")
    Periodo2 = Format(Periodo2, "mm/dd/yyyy")
    If ImageCombo1.Text = "" Then
        strSQL = "SELECT * FROM Gestion WHERE Fecha BETWEEN #" & Periodo1 & "# AND #" & Periodo2 & _
        "# ORDER BY FEcha,Hora,Apto"
    Else
        If ImageCombo1.SelectedItem.Indentation = 0 Then
            strSQL = "SELECT * FROM Gestion WHERE Fecha BETWEEN #" & Periodo1 & "# AND #" & Periodo2 _
            & "# ORDER BY FEcha,Hora,Apto;"
        Else
            strSQL = "SELECT * FROM Gestion WHERE Apto='" & apto & "' AND Fecha BETWEEN #" & _
            Periodo1 & "# AND #" & Periodo2 & "#ORDER By Fecha,Hora;"
        End If
    End If
Else
    strSQL = Diario
End If
'
objRst.Open strSQL, ObjCnn, adOpenKeyset, adLockOptimistic
Call rtnLimpiar_Grid(gridGestion)
'
With objRst
    If Not .EOF Or Not .BOF Then
        strQDF = strSQL
        .MoveFirst: i = 0
        gridGestion.Rows = .RecordCount + 1
        Do
            i = i + 1
            gridGestion.TextMatrix(i, 0) = !fecha
            gridGestion.TextMatrix(i, 1) = Format(!Hora, "hh:mm:ss ampm")
            gridGestion.TextMatrix(i, 2) = !apto
            gridGestion.TextMatrix(i, 3) = !Telf
            gridGestion.TextMatrix(i, 4) = !Contacto
            gridGestion.TextMatrix(i, 5) = !Resultado
            .MoveNext
        Loop Until .EOF
        gridGestion.Col = 3
        '
    End If
    .Close
End With
'
End Sub

Private Sub Config_Grid()
'variables locales
Dim ancho As Long
'
With gridGestion
    .ColWidth(0) = 1100
    .ColWidth(1) = 1200
    .ColWidth(2) = 700
    .ColWidth(3) = 1400
    .ColWidth(4) = 2300
    For i = 0 To 4: ancho = ancho + .ColWidth(i)
    Next
    .ColWidth(5) = .Width - ancho - 200
    .ColAlignmentFixed = flexAlignCenterCenter
End With
End Sub

Private Sub Cancelar()
MskTelefono.Text = Mascara
txtGestion(9) = ""
txtGestion(10) = ""
apto = ""
End Sub

Private Sub eliminar_registro()
'elimina el registr seleccionado
Dim ObjCnn As ADODB.Connection
Dim strSQL As String, n%

Set ObjCnn = New ADODB.Connection

ObjCnn.Open cnnOLEDB + mcDatos

strSQL = "DELETE * FROM Gestion WHERE Fecha=#" & strFecha & "# AND " & _
" HORA=#12/30/1899 " & Format(timHora, "hh:mm:ss ampm") & "# AND telf='" & MskTelefono _
& "' AND contacto='" & txtGestion(9) & "' AND resultado='" & txtGestion(10) & _
"'" & IIf(gcNivel > nuAdministrador, " AND Usuario='" & gcUsuario & "'", "")
ObjCnn.Execute strSQL, n
Call Vista
End Sub
