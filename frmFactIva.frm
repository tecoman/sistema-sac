VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmFactIva 
   Caption         =   "Emisión Facturas I.V.A"
   ClientHeight    =   6360
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7080
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   6360
   ScaleWidth      =   7080
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Fra 
      Caption         =   "Salida: "
      Height          =   1155
      Left            =   5475
      TabIndex        =   9
      Top             =   180
      Width           =   1410
      Begin VB.OptionButton Opt 
         Caption         =   "Ventana"
         Height          =   285
         Index           =   1
         Left            =   120
         TabIndex        =   11
         Top             =   735
         Width           =   900
      End
      Begin VB.OptionButton Opt 
         Caption         =   "Impresora"
         Height          =   270
         Index           =   0
         Left            =   120
         TabIndex        =   10
         Top             =   345
         Value           =   -1  'True
         Width           =   1005
      End
   End
   Begin VB.CheckBox Chk 
      Height          =   285
      Left            =   5850
      TabIndex        =   7
      Top             =   1935
      Width           =   300
   End
   Begin VB.CommandButton Cmd 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      Height          =   495
      Index           =   2
      Left            =   5895
      TabIndex        =   6
      Top             =   5505
      Width           =   1005
   End
   Begin VB.CommandButton Cmd 
      Caption         =   "Imprimir"
      Height          =   495
      Index           =   1
      Left            =   5895
      TabIndex        =   5
      Top             =   4815
      Width           =   1005
   End
   Begin VB.CommandButton Cmd 
      Caption         =   "Ver Facturas"
      Height          =   495
      Index           =   0
      Left            =   3540
      TabIndex        =   3
      Tag             =   "0"
      Top             =   720
      Width           =   1635
   End
   Begin MSMask.MaskEdBox mskPer 
      Bindings        =   "frmFactIva.frx":0000
      DataField       =   "FRecep"
      Height          =   315
      Left            =   2115
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   780
      Width           =   1065
      _ExtentX        =   1879
      _ExtentY        =   556
      _Version        =   393216
      MaxLength       =   7
      Format          =   "MM/yyyy"
      Mask            =   "##/####"
      PromptChar      =   "_"
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid FlexFact 
      Bindings        =   "frmFactIva.frx":0022
      CausesValidation=   0   'False
      Height          =   4110
      Left            =   345
      TabIndex        =   4
      Tag             =   "800|1000|1300|1100|500"
      Top             =   1875
      Width           =   5325
      _ExtentX        =   9393
      _ExtentY        =   7250
      _Version        =   393216
      Cols            =   5
      FixedCols       =   0
      BackColorFixed  =   -2147483646
      ForeColorFixed  =   -2147483639
      BackColorSel    =   65280
      ForeColorSel    =   -2147483646
      BackColorBkg    =   -2147483636
      GridColor       =   -2147483633
      AllowBigSelection=   0   'False
      FocusRect       =   2
      HighLight       =   0
      GridLinesFixed  =   0
      GridLinesUnpopulated=   4
      MergeCells      =   1
      AllowUserResizing=   1
      FormatString    =   "^Cod. Inm |^Nº Doc. |>Monto|Fecha |^Sel"
      BandDisplay     =   1
      BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   5
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   0
   End
   Begin MSMask.MaskEdBox fecha 
      Bindings        =   "frmFactIva.frx":0037
      DataField       =   "FRecep"
      Height          =   315
      Left            =   2100
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   1200
      Width           =   1065
      _ExtentX        =   1879
      _ExtentY        =   556
      _Version        =   393216
      MaxLength       =   10
      Format          =   "dd/MM/yyyy"
      Mask            =   "##/##/####"
      PromptChar      =   "_"
   End
   Begin VB.Image Img 
      Enabled         =   0   'False
      Height          =   480
      Index           =   1
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Img 
      Enabled         =   0   'False
      Height          =   480
      Index           =   0
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Label lbl 
      Caption         =   "Solo facturas por imprimir"
      Height          =   990
      Index           =   2
      Left            =   6165
      TabIndex        =   8
      Top             =   1965
      Width           =   645
   End
   Begin VB.Line Lin 
      BorderColor     =   &H00FFFFFF&
      Index           =   1
      X1              =   360
      X2              =   6800
      Y1              =   1575
      Y2              =   1575
   End
   Begin VB.Line Lin 
      BorderColor     =   &H00C0C0C0&
      BorderWidth     =   2
      Index           =   0
      X1              =   360
      X2              =   6800
      Y1              =   1575
      Y2              =   1575
   End
   Begin VB.Label lbl 
      Caption         =   "Ingrese el período de facturación mm-yyyy:"
      Height          =   525
      Index           =   1
      Left            =   375
      TabIndex        =   1
      Top             =   750
      Width           =   2010
   End
   Begin VB.Label lbl 
      Caption         =   "Impresíon de facturas Gastos Administrativos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   0
      Left            =   360
      TabIndex        =   0
      Top             =   375
      Width           =   5370
   End
End
Attribute VB_Name = "frmFactIva"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim blnSalir As Boolean, blnProceso As Boolean, blnDetener As Boolean


Private Sub chk_Click(): If IsDate("01/" & mskPer) Then Call Mostrar_Facturas
End Sub

Private Sub Cmd_Click(index As Integer)
'variables locales
Select Case index
    '
    Case 0  'visualizar facturas
        If cmd(0).Tag = 0 Then
            cmd(0).Caption = "Detener"
            cmd(0).Tag = 1
            Call Mostrar_Facturas
            
        Else
            cmd(0).Tag = 0
            cmd(0).Caption = "Ver Facturas"
            blnDetener = True
        End If
        
    Case 1  'imprimir
        If FlexFact.TextMatrix(1, 0) <> "" Then Call Printer_Factura
        
    Case 2  'salir
        If blnpreoceso Then
            blnSalir = True
        Else
            Unload Me
            Set frmFactIva = Nothing
        End If
    
            
End Select
'
End Sub

Private Sub FlexFact_Click()
With FlexFact
    If .Col = .Cols - 1 And .Row >= 1 Then
        .Col = .Cols - 1
        Set .CellPicture = img(IIf(.CellPicture = img(0), 1, 0))
    End If
End With
End Sub

Private Sub Form_Load()
'
Call centra_titulo(FlexFact, True)
img(0).Picture = LoadResPicture("UNCHECKED", vbResBitmap)
img(1).Picture = LoadResPicture("CHECKED", vbResBitmap)
FlexFact.Col = FlexFact.Cols - 1
FlexFact.Row = 1
Set FlexFact.CellPicture = img(0)
FlexFact.CellPictureAlignment = flexAlignCenterCenter
FlexFact.ColAlignment(3) = flexAlignCenterCenter
'
End Sub


Private Sub Mostrar_Facturas()
'variables locales
Dim n As Long, errLocal As Long
Dim strSql As String
Dim Periodo As Date, PeriodoF As Date
Dim rstInm(2) As ADODB.Recordset
Dim CnnInm As ADODB.Connection

mskPer.AllowPrompt = True

If Not IsDate("01/" & mskPer) Then
    MsgBox "Introdujo un período no válido", vbCritical, App.ProductName
Else
    FlexFact.Enabled = False
    MousePointer = vbHourglass
    Call rtnBitacora("Listar Facturas período " & mskPer)
    blnProceso = True
    
    Call rtnLimpiar_Grid(FlexFact)
    FlexFact.Rows = 2
    FlexFact.Col = FlexFact.Cols - 1
    FlexFact.Row = 1
    Set FlexFact.CellPicture = img(0)
    FlexFact.CellPictureAlignment = flexAlignCenterCenter
    n = 1
    Set rstInm(0) = New ADODB.Recordset
    rstInm(0).CursorLocation = adUseClient
    
    '
    Periodo = "01/" & mskPer
    PeriodoF = DateAdd("M", 1, Periodo)
    PeriodoF = DateAdd("D", -1, PeriodoF)

    strSql = "SELECT * FROM Cpp WHERE Fact LIKE 'F%' AND (Frecep " _
    & ">=#" & Format(Periodo, "mm/dd/yyyy") & "# AND Frecep<=#" & _
    Format(PeriodoF, "mm/dd/yy") & "#) ORDER BY Frecep,CodInm"
    
    rstInm(0).Open strSql, cnnConexion, adOpenKeyset, adLockOptimistic, adCmdText
    
    If CHK.Value = vbChecked Then rstInm(0).Filter = "Emitida = False"

    
    If Not (rstInm(0).EOF And rstInm(0).BOF) Then
        rstInm(0).MoveFirst
         
        Do
            '
            DoEvents
            MousePointer = vbHourglass
            '
            If blnSalir Then
                Unload Me
                Set frmFactIva = Nothing
                Exit Sub
            ElseIf blnDetener Then
                blnDetener = False
                Me.FlexFact.Enabled = True
                MousePointer = vbDefault
                Exit Sub
            End If
            '
            With FlexFact
            
                If n > 1 Then
                    .AddItem rstInm(0)!CodInm
                Else
                    .TextMatrix(n, 0) = rstInm(0)!CodInm
                End If
                .Col = .Cols - 1
                .Row = n
                Set .CellPicture = img(rstInm(0)!Emitida)
                .CellPictureAlignment = flexAlignCenterCenter
                .TextMatrix(n, 1) = rstInm(0)!NDoc
                .TextMatrix(n, 2) = Format(rstInm(0)!Total, "#,##0.00 ")
                .TextMatrix(n, 3) = Format(rstInm(0)!Frecep, "dd-mm-yy")
                    
            End With
            n = n + 1
            rstInm(0).MoveNext
            
        Loop Until rstInm(0).EOF
        FlexFact.Enabled = True
        
    End If
    MousePointer = vbDefault
    rstInm(0).Close
    Set rstInm(0) = Nothing
    blnProceso = False
    blnDetener = False
    cmd(0).Caption = "Ver Facturas"
    cmd(0).Tag = 0
End If
End Sub

Private Sub Printer_Factura()
'variables locales
Dim F&, J&, i&
Dim CodInm$, Nfact$
Dim crReport As CRAXDRT.Report
Dim crApp As CRAXDRT.Application
Dim crSection As CRAXDRT.Section
Dim crObject As Object
Dim crSubReportObj As CRAXDRT.Report '
Dim crxSRObject As CRAXDRT.SubreportObject
Dim frmLocal As frmView
'
With FlexFact
'
    Set crApp = New CRAXDRT.Application
    Set crReport = New CRAXDRT.Report
    
    For i = 1 To .Rows - 1
        .Col = .Cols - 1
        .Row = i
        If .CellPicture = img(1) Then
            '
            CodInm = .TextMatrix(i, 0)
            Nfact = .TextMatrix(i, 1)
            'aqui imprime el reporte
            MousePointer = vbHourglass
                        
            With crReport
                '
                Set crReport = crApp.OpenReport(gcReport & "fact_IVA.rpt", 1)
                
                crReport.Database.Tables(1).Location = gcPath & "\sac.mdb"
                crReport.Database.Tables(2).Location = gcPath & "\sac.mdb"
                crReport.FormulaFields.GetItemByName("CodInm").Text = "'" + CodInm + "'"
                crReport.FormulaFields.GetItemByName("Ndoc").Text = "'" + Nfact + "'"
                
                '
                For Each crSection In crReport.Sections
                    For Each crObject In crSection.ReportObjects
                        If crObject.Kind = crSubreportObject Then
                            Set crxSRObject = crObject
                            Set crSubReportObj = crxSRObject.OpenSubreport
                            crSubReportObj.Database.Tables(1).Location = gcPath & "\sac.mdb"
                            'crSubReportObj.Database.Tables(2).Location = gcPath & "\" & CodInm & "\inm.mdb"
                            crSubReportObj.Database.Tables(2).Location = gcPath & "\sac.mdb"
                            crSubReportObj.Database.Tables(3).Location = gcPath & "\" & CodInm & "\inm.mdb"
                            crSubReportObj.Database.Tables(4).Location = gcPath & "\sac.mdb"
                            crSubReportObj.FormulaFields.GetItemByName("sFecha").Text = "'" & IIf(IsDate(fecha), fecha, "") & "'"
                        End If
                    Next
                Next
                '
                If opt(0) Then  'salida por impresora
                    crReport.DisplayProgressDialog = False
                    crReport.PrintOut False
                Else
                    '
                    
                    
                    Set frmLocal = New frmView
                    frmLocal.crView.ReportSource = crReport
            
            
                    frmLocal.crView.ViewReport
                    While frmLocal.crView.IsBusy
                        DoEvents
                    Wend
                    If Screen.Width / Screen.TwipsPerPixelX = 1024 Then
                        frmLocal.crView.Zoom 120
                    Else
                        frmLocal.crView.Zoom 1
                    End If
                    frmLocal.Caption = "Factura " & Nfact
                    frmLocal.Show
                    
                    '
                End If
                
                
                Call rtnBitacora("Emisión Factura " & Nfact & " del Inm.:" & CodInm & " Mes: " _
                & mskPer)
                If Err <> 0 Then
                '
                    MsgBox Err.Description, vbCritical, "Error " & Err
                    Call rtnBitacora(Err.Description)
                    '
                End If
                
            End With
            MousePointer = vbDefault
            '
            'Set crReporte = Nothing
        End If
        '
    Next
    '
End With
'
End Sub
