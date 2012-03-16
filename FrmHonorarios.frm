VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Begin VB.Form FrmHonorarios 
   Caption         =   "Honorarios"
   ClientHeight    =   6765
   ClientLeft      =   255
   ClientTop       =   600
   ClientWidth     =   10110
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   6765
   ScaleWidth      =   10110
   WindowState     =   2  'Maximized
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   5655
      Top             =   3975
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   20
      ImageHeight     =   20
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmHonorarios.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmHonorarios.frx":031C
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin ComCtl3.CoolBar clb 
      Height          =   705
      Left            =   5190
      TabIndex        =   30
      Top             =   7050
      Width           =   6540
      _ExtentX        =   11536
      _ExtentY        =   1244
      BandCount       =   2
      BandBorders     =   0   'False
      VariantHeight   =   0   'False
      EmbossPicture   =   -1  'True
      _CBWidth        =   6540
      _CBHeight       =   705
      _Version        =   "6.0.8169"
      Child1          =   "tlb"
      MinHeight1      =   645
      Width1          =   2595
      NewRow1         =   0   'False
      Child2          =   "pgbHono"
      MinHeight2      =   645
      Width2          =   3705
      NewRow2         =   0   'False
      Begin MSComctlLib.Toolbar tlb 
         Height          =   645
         Left            =   165
         TabIndex        =   32
         Top             =   30
         Width           =   2400
         _ExtentX        =   4233
         _ExtentY        =   1138
         ButtonWidth     =   1931
         ButtonHeight    =   1138
         Style           =   1
         ImageList       =   "ImageList1"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   2
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Imprimir"
               Key             =   "Print"
               ImageIndex      =   1
               Style           =   5
               Object.Width           =   1e-4
               BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
                  NumButtonMenus  =   2
                  BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "Window"
                     Text            =   "Pantalla"
                  EndProperty
                  BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "Printer"
                     Text            =   "Impresora"
                  EndProperty
               EndProperty
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "      Salir       "
               Key             =   "Close"
               ImageIndex      =   2
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.ProgressBar pgbHono 
         Height          =   645
         Left            =   2760
         TabIndex        =   31
         Top             =   30
         Visible         =   0   'False
         Width           =   3690
         _ExtentX        =   6509
         _ExtentY        =   1138
         _Version        =   393216
         Appearance      =   0
         Min             =   1e-4
         Max             =   200
         Scrolling       =   1
      End
   End
   Begin VB.Frame fraHono 
      Caption         =   "Totales:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2000
      Index           =   4
      Left            =   180
      TabIndex        =   21
      Top             =   5775
      Width           =   4800
      Begin VB.TextBox TxtHono 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   0
         Left            =   2910
         Locked          =   -1  'True
         TabIndex        =   29
         Top             =   390
         Width           =   1785
      End
      Begin VB.TextBox TxtHono 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   3
         Left            =   2910
         Locked          =   -1  'True
         TabIndex        =   27
         Top             =   1560
         Width           =   1785
      End
      Begin VB.TextBox TxtHono 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   2
         Left            =   2910
         Locked          =   -1  'True
         TabIndex        =   26
         Top             =   1170
         Width           =   1785
      End
      Begin VB.TextBox TxtHono 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   1
         Left            =   2910
         Locked          =   -1  'True
         TabIndex        =   25
         Top             =   780
         Width           =   1785
      End
      Begin VB.Label lblHono 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Pago Condominio:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   5
         Left            =   150
         TabIndex        =   28
         Top             =   405
         Width           =   2700
      End
      Begin VB.Label lblHono 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Honorarios Netos:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   8
         Left            =   150
         TabIndex        =   24
         Top             =   1575
         Width           =   2700
      End
      Begin VB.Label lblHono 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Deducciones Hono. Abogado"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   7
         Left            =   150
         TabIndex        =   23
         Top             =   1185
         Width           =   2700
      End
      Begin VB.Label lblHono 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Honorarios de Abogado"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   6
         Left            =   150
         TabIndex        =   22
         Top             =   795
         Width           =   2700
      End
   End
   Begin MSFlexGridLib.MSFlexGrid FlexHono 
      Height          =   6600
      Left            =   5205
      TabIndex        =   20
      Top             =   345
      Width           =   6525
      _ExtentX        =   11509
      _ExtentY        =   11642
      _Version        =   393216
      Cols            =   6
      FixedCols       =   0
      BackColorFixed  =   -2147483646
      ForeColorFixed  =   -2147483639
      FormatString    =   "Cod.Inm |Apto. |Pag. Cond. |Hono. Abg. |Deducciones |Neto Hono."
   End
   Begin VB.Frame fraHono 
      Caption         =   "Opciones:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5325
      Index           =   0
      Left            =   180
      TabIndex        =   0
      Top             =   225
      Width           =   4800
      Begin VB.Frame fraHono 
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   1335
         Index           =   2
         Left            =   50
         TabIndex        =   6
         Top             =   1950
         Width           =   4600
         Begin VB.OptionButton optHono 
            Caption         =   "Selección:"
            CausesValidation=   0   'False
            Height          =   285
            Index           =   3
            Left            =   250
            TabIndex        =   8
            Top             =   930
            Width           =   1215
         End
         Begin VB.OptionButton optHono 
            Caption         =   "Todos"
            CausesValidation=   0   'False
            Height          =   285
            Index           =   2
            Left            =   250
            TabIndex        =   7
            Top             =   510
            Value           =   -1  'True
            Width           =   1215
         End
         Begin MSDataListLib.DataCombo dtcHono 
            Height          =   315
            Index           =   0
            Left            =   1485
            TabIndex        =   9
            Top             =   915
            Width           =   960
            _ExtentX        =   1693
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
            Style           =   2
            ListField       =   "CodInm"
            Text            =   ""
            Object.DataMember      =   ""
         End
         Begin VB.Label lblHono 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Inmuble:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   270
            Index           =   3
            Left            =   105
            TabIndex        =   10
            Top             =   180
            Width           =   1215
         End
      End
      Begin VB.Frame fraHono 
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   1200
         Index           =   3
         Left            =   50
         TabIndex        =   1
         Top             =   3825
         Width           =   4600
         Begin VB.OptionButton optHono 
            Caption         =   "Selección:"
            CausesValidation=   0   'False
            Height          =   285
            Index           =   5
            Left            =   250
            TabIndex        =   3
            Top             =   690
            Width           =   1215
         End
         Begin VB.OptionButton optHono 
            Caption         =   "Todas"
            CausesValidation=   0   'False
            Height          =   285
            Index           =   4
            Left            =   250
            TabIndex        =   2
            Top             =   285
            Value           =   -1  'True
            Width           =   1215
         End
         Begin MSDataListLib.DataCombo dtcHono 
            Height          =   315
            Index           =   1
            Left            =   1515
            TabIndex        =   4
            Top             =   675
            Width           =   930
            _ExtentX        =   1640
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
            Style           =   2
            ListField       =   "IDTaquilla"
            Text            =   ""
         End
         Begin VB.Label lblHono 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Taquilla:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   270
            Index           =   4
            Left            =   100
            TabIndex        =   5
            Top             =   0
            Width           =   1215
         End
      End
      Begin VB.Frame fraHono 
         BorderStyle     =   0  'None
         Caption         =   "Periodo"
         Height          =   1200
         Index           =   1
         Left            =   50
         TabIndex        =   11
         Top             =   360
         Width           =   4600
         Begin VB.OptionButton optHono 
            Caption         =   "Rango Fecha:"
            Height          =   435
            Index           =   1
            Left            =   250
            TabIndex        =   19
            Top             =   810
            Width           =   975
         End
         Begin VB.OptionButton optHono 
            Caption         =   "Mes"
            Height          =   285
            Index           =   0
            Left            =   250
            TabIndex        =   18
            Top             =   405
            Value           =   -1  'True
            Width           =   1215
         End
         Begin VB.ComboBox cmbHono 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            ItemData        =   "FrmHonorarios.frx":0638
            Left            =   1530
            List            =   "FrmHonorarios.frx":0662
            Style           =   2  'Dropdown List
            TabIndex        =   12
            Top             =   375
            Width           =   1290
         End
         Begin MSComCtl2.DTPicker DTPicker1 
            Height          =   315
            Index           =   0
            Left            =   1545
            TabIndex        =   13
            Top             =   855
            Width           =   1245
            _ExtentX        =   2196
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            CustomFormat    =   "abcdefghijklmnñopqrstuvwxyz"
            Format          =   85262337
            CurrentDate     =   37505
         End
         Begin MSComCtl2.DTPicker DTPicker1 
            Height          =   315
            Index           =   1
            Left            =   3270
            TabIndex        =   14
            Top             =   855
            Width           =   1245
            _ExtentX        =   2196
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            CustomFormat    =   "abcdefghijklmnñopqrstuvwxyz"
            Format          =   85262337
            CurrentDate     =   37505
         End
         Begin VB.Label lblHono 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Del:"
            Enabled         =   0   'False
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   1
            Left            =   1215
            TabIndex        =   16
            Top             =   915
            Width           =   285
         End
         Begin VB.Label lblHono 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Perídodo:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   270
            Index           =   0
            Left            =   100
            TabIndex        =   17
            Top             =   0
            Width           =   1215
         End
         Begin VB.Label lblHono 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Al:"
            Enabled         =   0   'False
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   2
            Left            =   3015
            TabIndex        =   15
            Top             =   915
            Width           =   180
         End
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         Index           =   2
         X1              =   270
         X2              =   4515
         Y1              =   3525
         Y2              =   3525
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         Index           =   1
         X1              =   300
         X2              =   4545
         Y1              =   1800
         Y2              =   1800
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00808080&
         BorderWidth     =   2
         Index           =   0
         X1              =   330
         X2              =   4530
         Y1              =   1800
         Y2              =   1800
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00808080&
         BorderWidth     =   2
         Index           =   3
         X1              =   300
         X2              =   4500
         Y1              =   3525
         Y2              =   3525
      End
   End
End
Attribute VB_Name = "FrmHonorarios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'SAC------------------------------------------------------------------------------------------
    Dim RstCajas As ADODB.Recordset
    Dim codDedHonoA$, codPagCond$, codHonoA$, codAbono$
    

    '10/09/2002-----------------------------------------------------------------------------------
    Private Sub cmbHono_Click()
    '---------------------------------------------------------------------------------------------
    '
    cmbHono.Refresh
    If cmbHono.Text <> "" Then Call rtnFlexHono
    '
    End Sub

    
    '10/09/2002-----------------------------------------------------------------------------------
    Private Sub dtcHono_Click(Index As Integer, Area As Integer)
    '---------------------------------------------------------------------------------------------
    '
    dtcHono(Index).Refresh
    If Area = 2 Then Call rtnFlexHono
    End Sub

    '10/09/2002-----------------------------------------------------------------------------------
    Private Sub DTPicker1_CloseUp(Index As Integer)
    '---------------------------------------------------------------------------------------------
    '
    DTPicker1(0).Refresh: DTPicker1(1).Refresh
    If DTPicker1(0).Value <= DTPicker1(1).Value Then Call rtnFlexHono
    End Sub

    '---------------------------------------------------------------------------------------------
    Private Sub Form_Load() '
    '---------------------------------------------------------------------------------------------
    Dim strSQL$    'Cadena SQL
    Dim rstCod As New ADODB.Recordset
    '
    Set RstCajas = New ADODB.Recordset
    
    Set dtcHono(0).RowSource = FrmAdmin.objRst
    dtcHono(0).ListField = "CodInm"
    '
    strSQL = "SELECT * FROM Taquillas ORDER BY IDTaquilla"
    RstCajas.Open strSQL, cnnConexion, adOpenStatic, adLockReadOnly, adCmdText
    strSQL = "SELECT * FROM Inmueble WHERE CodInm='" & sysCodInm & "';"
    Set dtcHono(1).RowSource = RstCajas
    rstCod.Open strSQL, cnnConexion, adOpenStatic, adLockReadOnly, adCmdText
    codPagCond = rstCod!CodPagoCondominio
    codDedHonoA = rstCod!CodRebHA
    codHonoA = rstCod!CodHA
    codAbono = rstCod!CodAbonoCta
    rstCod.Close
    Set rstCod = Nothing
    '
    With FlexHono
        .Cols = 6
        .Row = 0
        .ColWidth(0) = 700
        .ColWidth(1) = 550
        For I = 2 To 5
            .Col = I
            .ColWidth(I) = 1200
            .CellAlignment = flexAlignCenterCenter
        Next
        .ColAlignment(0) = flexAlignCenterCenter
        .ColAlignment(1) = flexAlignCenterCenter
    End With
    cmbHono = cmbHono.List(Month(Date) - 1)
    
    End Sub

    '---------------------------------------------------------------------------------------------
    Private Sub optHono_Click(Index As Integer) '
    '---------------------------------------------------------------------------------------------
    '
    Select Case Index
        Case 0  'Periodo 'Todos'
    '   -------------------------------
            Call rtnPeriodo("False")
            Call rtnFlexHono
            
        Case 1  'Periodo 'Rango'
    '   -------------------------------
            Call rtnPeriodo("True")
            If DTPicker1(0).Value <= DTPicker1(1).Value Then Call rtnFlexHono
            
        Case 2  'Inmuble 'Todos'
    '   -------------------------------
            dtcHono(0).Enabled = False
            Call rtnFlexHono
            
        Case 3  'Inmueble 'Seleccion'
    '   -------------------------------
            dtcHono(0).Enabled = True
            If dtcHono(0) <> "" And Len(dtcHono(0)) = 4 Then Call rtnFlexHono
        
        Case 4  'Caja   'Todas'
    '   -------------------------------
            dtcHono(1).Enabled = False
            Call rtnFlexHono
        
        Case 5  'Caja 'Seleccion'
    '   -------------------------------
            dtcHono(1).Enabled = True
            If dtcHono(1) <> "" Then rtnFlexHono
            
    End Select
    '
    End Sub
    
    '------------------------------------------------------
    Private Sub rtnPeriodo(StrEstado As String)  '
    '------------------------------------------------------
    '
        For I = 0 To 1
            lblHono(I + 1).Enabled = StrEstado
            DTPicker1(I).Enabled = StrEstado
        Next
        cmbHono.Enabled = Not cmbHono.Enabled
    '
    End Sub
    
    '10/09/2002-----------------------------------------------------------------------------------
    Private Sub rtnFlexHono()
    '---------------------------------------------------------------------------------------------
    '
    Dim ADOcontrol As New ADODB.Recordset
    Dim strSQL$, Date1$, Date2$, Codigo$, I%
    Dim vecCHR(2 To 4) As Currency
    '
    

    pgbHono.Visible = True
    pgbHono.Value = 0.01
    '---------------------------------------------------------------------------------------------
    If optHono(0) Then
        Date1 = CDate("01/" & cmbHono & "/" & Year(Date))
        Date2 = DateAdd("m", 1, Date1)
        Date2 = DateAdd("d", -1, Date2)
    Else
       Date1 = DTPicker1(0).Value
       Date2 = DTPicker1(1).Value
    End If
    strSQL = "WHERE Mc.FechaMovimientoCaja Between #" & CStr(Format(Date1, "mm/dd/yy")) _
    & "# AND #" & CStr(Format(Date2, "mm/dd/yy")) & "#"
    If optHono(3).Value = True Then
        strSQL = strSQL & IIf(strSQL = "", "", " AND ") & "MC.InmuebleMovimientoCaja='" & _
        dtcHono(0) & "'"
    End If
    If optHono(5).Value = True Then
        strSQL = strSQL & IIf(strSQL = "", "", " AND ") & "MC.IDTaquilla=" & CInt(dtcHono(1))
    End If
    '---------------------------------------------------------------------------------------------
    
    For I = 1 To 3
        TxtHono(I) = "0,00"
        If I = 3 Then
            Codigo = "AND (DE.CodGasto='" & codDedHonoA & "' or DE.CodGasto='9007')"
        Else
            Codigo = " AND (P.CodGasto='" & IIf(I = 1, codPagCond & "' or P.CodGasto='" & codAbono & "')" _
            , codHonoA & "' or P.CodGasto='8101')")
        End If
        pgbHono.Value = 50 * I
        Call CreateQDF(strSQL & Codigo, I)
    Next
    TxtHono(0) = "0,00"
    '
    'strSql = "SELECT Hono1.InmuebleMovimientoCaja, Hono1.AptoMovimientoCaja,Hono1.PC, Hono2.PC," _
    & "Hono3.RHA FROM (Hono1 LEFT JOIN Hono2 ON Hono1.IDRecibo = Hono2.IDRecibo) LEFT JOIN Hono" _
    & "3 ON Hono2.IDRecibo = Hono3.IDRecibo;"
    'Call rtnGenerator(gcPath & "\sac.mdb", strSql, "qdfHonorarios")
    pgbHono.Value = 175
'    For I = 0 To 30000
'        'retardo para que reconozca las consultas
'    Next
    '
    On Error Resume Next
    With ADOcontrol
1      .Open "Hono4", cnnConexion, adOpenKeyset, adLockOptimistic, adCmdTable
        If Err.Number = -2147217865 Then Err.Clear: GoTo 1
        FlexHono.Rows = .RecordCount + 1
        If Not .EOF Then
        .MoveFirst
        I = 0
        pgbHono.Value = 200
        Do Until .EOF
            I = I + 1
            For K = 0 To 4
                FlexHono.TextMatrix(I, K) = IIf(IsNull(.Fields(K)), _
                IIf(K >= 2 And K <= 4, 0, ""), IIf(K >= 2 And K <= 4, _
                Format(.Fields(K), "#,##0.00"), .Fields(K)))
            Next
            FlexHono.TextMatrix(I, 5) = Format(CCur(FlexHono.TextMatrix(I, 3) - _
            FlexHono.TextMatrix(I, 4)), "#,##0.00")
            For j = 2 To 4
                vecCHR(j) = vecCHR(j) + IIf(IsNull(.Fields(j)), 0, .Fields(j))
            Next
            .MoveNext
        Loop
        For j = 0 To 2
            TxtHono(j) = Format(vecCHR(j + 2), "#,##0.00")
        Next
        TxtHono(3) = Format(vecCHR(3) - vecCHR(4), "#,##0.00")
        End If
    End With
    pgbHono.Visible = False
    '
    End Sub
    
    '---------------------------------------------------------------------------------------------
    Private Sub CreateQDF(strCriterio$, j%)  '
    '---------------------------------------------------------------------------------------------
    '
    Dim strSQL1$, strSQL2$, strSQL3$
    Dim qdfTemp As QueryDef
    '
    strSQL1 = "SELECT P.IDRecibo, P.CodGasto, Sum(P.Monto) AS PC, MC.InmuebleMovimientoCaja, " _
    & "MC.AptoMovimientoCaja FROM MovimientoCaja As MC INNER JOIN Periodos  as P ON MC.IDRecibo" _
    & "= P.IDRecibo " & strCriterio & " GROUP BY P.IDRecibo, P.CodGasto, MC.InmuebleMovimientoC" _
    & "aja, MC.AptoMovimientoCaja"
    '
      strSQL2 = "SELECT P.IDRecibo, De.CodGasto, Sum(De.Monto) AS RHA  FROM (MovimientoCaja as MC" _
    & " INNER JOIN Periodos as P ON MC.IDRecibo = P.IDRecibo) LEFT JOIN Deducciones as De ON P." _
    & "IDPeriodos = De.IDPeriodos " & strCriterio & " GROUP BY P.IDRecibo, De.CodGasto, MC.Inmu" _
    & "ebleMovimientoCaja, MC.AptoMovimientoCaja"
    strSQL3 = IIf(j = 3, strSQL2, strSQL1)
    '
    Call rtnGenerator(gcPath & "\sac.mdb", strSQL3, "Hono" & j)
    '
    End Sub

    Private Sub tlb_ButtonClick(ByVal Button As MSComctlLib.Button)
    '
    Select Case UCase(Button.Key)
        Case "PRINT"
           Call Print_report(crPantalla)
        Case "CLOSE"
            Unload Me
            Set FrmHonorarios = Nothing
    End Select
    '
    End Sub

    Private Sub Print_report(Optional Salida As crSalida)
    'variables locales
    Dim strDesde$, strHasta$, strInm$, strTaquilla$, errLocal&
    Dim rpReporte As ctlReport
    If optHono(0).Value = True Then
        strDesde = Format(CDate("01/" & cmbHono & "/" & Year(Date)), "dd/mm/yy")
        strHasta = Format(CDate("01/" & cmbHono & "/" & Year(Date)), "dd/mm/yy")
        strHasta = DateAdd("m", 1, CDate(strHasta))
        strHasta = DateAdd("d", -1, CDate(strHasta))
    Else
        strDesde = DTPicker1(0)
        strHasta = DTPicker1(1)
    End If
    '
    If optHono(2).Value = True Then
        strInm = "Todos"
    Else
        strInm = dtcHono(0)
    End If
    '
    If optHono(4).Value = True Then
        strTaquilla = "Todas"
    Else
        strTaquilla = dtcHono(1)
    End If
    'Call clear_Crystal(FrmAdmin.rptReporte)
    '
    Set rpReporte = New ctlReport
    With rpReporte
    
        .OrigenDatos(0) = gcPath & "\sac.mdb"
        .OrigenDatos(1) = gcPath & "\sac.mdb"
        .Reporte = gcReport + "cxc_pagos.rpt"
        .Formulas(0) = "Desde='" & strDesde & "'"
        .Formulas(1) = "Hasta='" & strHasta & "'"
        .Formulas(2) = "CodInm='" & strInm & "'"
        .Formulas(3) = "Taquilla='" & strTaquilla & "'"
        .TituloVentana = "Consulta de Pagos"
        'errlocal = .PrintReport
        .Imprimir
        Call rtnBitacora("Imprimir consulta de pagos..Inm:" & strInm & " De:" & strDesde _
        & " Al:" & strHasta)
'        If errlocal <> 0 Then
'            MsgBox .LastErrorString, vbCritical, "Error " & .LastErrorNumber
'            Call rtnBitacora("Error " & .LastErrorNumber & " al imprimir el reporte")
'        End If
        
    End With
    Set rpReporte = Nothing
    '
    End Sub


