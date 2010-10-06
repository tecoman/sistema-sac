VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form FrmEntregaCaja 
   Caption         =   "Entrega General de Caja"
   ClientHeight    =   5970
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8535
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   5970
   ScaleWidth      =   8535
   WindowState     =   2  'Maximized
   Begin VB.Frame FraEntrega 
      Caption         =   "Transf."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4665
      Index           =   3
      Left            =   5950
      TabIndex        =   22
      Top             =   2000
      Width           =   2600
      Begin MSFlexGridLib.MSFlexGrid GridTPago 
         Height          =   3615
         Index           =   3
         Left            =   100
         TabIndex        =   23
         Top             =   500
         Width           =   2235
         _ExtentX        =   3942
         _ExtentY        =   6376
         _Version        =   393216
         BackColorBkg    =   -2147483636
         ScrollTrack     =   -1  'True
         FocusRect       =   2
         ScrollBars      =   2
         SelectionMode   =   1
         BorderStyle     =   0
         Appearance      =   0
      End
   End
   Begin VB.CommandButton CmdEntrega 
      Caption         =   "&Salir"
      Height          =   960
      Index           =   1
      Left            =   4725
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   330
      Width           =   1140
   End
   Begin VB.CommandButton CmdEntrega 
      Caption         =   "&Imprimir"
      Height          =   960
      Index           =   0
      Left            =   3570
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   330
      Width           =   1140
   End
   Begin VB.Frame Frame1 
      Height          =   1425
      Left            =   540
      TabIndex        =   6
      Top             =   225
      Width           =   2865
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         DrawMode        =   14  'Copy Pen
         Index           =   3
         X1              =   1230
         X2              =   2740
         Y1              =   1065
         Y2              =   1065
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         DrawMode        =   14  'Copy Pen
         Index           =   2
         X1              =   1215
         X2              =   2750
         Y1              =   615
         Y2              =   615
      End
      Begin VB.Line Line1 
         BorderWidth     =   2
         Index           =   1
         X1              =   1230
         X2              =   2730
         Y1              =   1065
         Y2              =   1065
      End
      Begin VB.Line Line1 
         BorderWidth     =   2
         Index           =   0
         X1              =   1230
         X2              =   2730
         Y1              =   615
         Y2              =   615
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "FECHA:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   0
         Left            =   120
         TabIndex        =   10
         Top             =   390
         Width           =   1215
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "MM/DD/AA"
         Height          =   285
         Index           =   1
         Left            =   1395
         TabIndex        =   9
         Top             =   390
         Width           =   1215
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "CAJERO:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   2
         Left            =   120
         TabIndex        =   8
         Top             =   795
         Width           =   1215
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Label1"
         Height          =   285
         Index           =   3
         Left            =   1395
         TabIndex        =   7
         Top             =   795
         Width           =   1215
      End
   End
   Begin VB.Frame FraEntrega 
      Caption         =   "Detalle Efectivo:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4665
      Index           =   2
      Left            =   8655
      TabIndex        =   4
      Top             =   2000
      Width           =   5000
      Begin MSFlexGridLib.MSFlexGrid GridTPago 
         Height          =   3615
         Index           =   2
         Left            =   135
         TabIndex        =   5
         Top             =   495
         Width           =   4095
         _ExtentX        =   7223
         _ExtentY        =   6376
         _Version        =   393216
         Cols            =   3
         BackColorBkg    =   -2147483636
         FocusRect       =   2
         HighLight       =   2
         ScrollBars      =   2
         PictureType     =   1
         BorderStyle     =   0
         Appearance      =   0
      End
   End
   Begin VB.Frame FraEntrega 
      Caption         =   "Cheques"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4665
      Index           =   1
      Left            =   540
      TabIndex        =   3
      Top             =   2000
      Width           =   2600
      Begin MSFlexGridLib.MSFlexGrid GridTPago 
         Height          =   3615
         Index           =   1
         Left            =   100
         TabIndex        =   0
         Top             =   500
         Width           =   2100
         _ExtentX        =   3704
         _ExtentY        =   6376
         _Version        =   393216
         BackColor       =   16777215
         BackColorBkg    =   -2147483636
         ScrollTrack     =   -1  'True
         FocusRect       =   2
         ScrollBars      =   2
         BorderStyle     =   0
         Appearance      =   0
      End
   End
   Begin VB.Frame FraEntrega 
      Caption         =   "Depositos"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4665
      Index           =   0
      Left            =   3245
      TabIndex        =   1
      Top             =   2000
      Width           =   2600
      Begin MSFlexGridLib.MSFlexGrid GridTPago 
         Height          =   3615
         Index           =   0
         Left            =   100
         TabIndex        =   2
         Top             =   500
         Width           =   2235
         _ExtentX        =   3942
         _ExtentY        =   6376
         _Version        =   393216
         BackColorBkg    =   -2147483636
         ScrollTrack     =   -1  'True
         FocusRect       =   2
         ScrollBars      =   2
         SelectionMode   =   1
         BorderStyle     =   0
         Appearance      =   0
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   200
      Left            =   4245
      Top             =   795
   End
   Begin VB.Line Line1 
      Index           =   9
      X1              =   8730
      X2              =   11025
      Y1              =   8415
      Y2              =   8415
   End
   Begin VB.Line Line1 
      Index           =   8
      X1              =   8730
      X2              =   11025
      Y1              =   8475
      Y2              =   8475
   End
   Begin VB.Line Line1 
      Index           =   7
      X1              =   8730
      X2              =   11025
      Y1              =   8085
      Y2              =   8085
   End
   Begin VB.Line Line1 
      Index           =   6
      X1              =   8745
      X2              =   11040
      Y1              =   8025
      Y2              =   8025
   End
   Begin VB.Line Line1 
      Index           =   5
      X1              =   8715
      X2              =   11010
      Y1              =   7665
      Y2              =   7665
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   12
      Left            =   8730
      TabIndex        =   21
      Top             =   8145
      Width           =   2265
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "0,00"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   7
      Left            =   8730
      TabIndex        =   20
      Top             =   7740
      Width           =   2265
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Diferencia:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   11
      Left            =   6480
      TabIndex        =   19
      Top             =   8145
      Width           =   2145
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Total General Caja:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   10
      Left            =   6480
      TabIndex        =   18
      Top             =   7740
      Width           =   2145
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Saldo Apertura:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   9
      Left            =   6480
      TabIndex        =   17
      Top             =   7335
      Width           =   2145
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Sub -Total Caja:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   5
      Left            =   6480
      TabIndex        =   16
      Top             =   6930
      Width           =   2145
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "0,00"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   8
      Left            =   8730
      TabIndex        =   15
      Top             =   7335
      Width           =   2265
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "0,00"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   6
      Left            =   8730
      TabIndex        =   13
      Top             =   6930
      Width           =   2265
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      Index           =   4
      X1              =   1515
      X2              =   5940
      Y1              =   7215
      Y2              =   7215
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "FIRMA:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   4
      Left            =   570
      TabIndex        =   12
      Top             =   7005
      Width           =   1215
   End
End
Attribute VB_Name = "FrmEntregaCaja"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim vDenominacion(0 To 20)
Public Contador As Integer
Dim K As Integer

    '---------------------------------------------------------------------------------------------
    Private Sub CmdEntrega_Click(index As Integer)  '
    '---------------------------------------------------------------------------------------------
    '
    Select Case index
    '
        Case 0  'Imprimir
    '   ---------------------
            FrmEntregaCaja.BackColor = &HFFFFFF
            Frame1.BackColor = &HFFFFFF
            For i = 0 To 2
                FraEntrega(i).BackColor = &HFFFFFF
                GridTPago(i).BackColorBkg = &HFFFFFF
                If i < 2 Then CmdEntrega(i).Visible = False
            Next i
            Printer.Orientation = 2 'vertical
            PrintForm
            Printer.Orientation = 1 'horizontal
            FrmEntregaCaja.BackColor = &H8000000F
            Frame1.BackColor = &H8000000F
            For i = 0 To 2
                FraEntrega(i).BackColor = &H8000000F
                GridTPago(i).BackColorBkg = &H8000000F
                If i < 2 Then CmdEntrega(i).Visible = True
            Next i
        Case 1  'Cerrar
    '   ---------------------
            Unload FrmEntregaCaja
            Set FrmEntregaCaja = Nothing
    '
    End Select
    '
    End Sub

    '---------------------------------------------------------------------------------------------
    Private Sub Form_Load() '
    '---------------------------------------------------------------------------------------------
    '
    Dim RstDenominacion As ADODB.Recordset
    '
    Set RstDenominacion = New ADODB.Recordset
    CmdEntrega(0).Picture = LoadResPicture("print", vbResIcon)
    CmdEntrega(1).Picture = LoadResPicture("Salir", vbResIcon)
    '
    apertura = Format(CurSaldoOpen, "#,##0.00")
    Label1(8) = apertura: Label1(6) = apertura
    '
    With RstDenominacion
        .Open "SELECT * FROM Denominacion ORDER By Billetes DESC", _
            cnnConexion, adOpenStatic, adLockOptimistic
        .MoveFirst
        i = 0
        Do Until .EOF
            vDenominacion(i) = Format(.Fields("Billetes"), "#,##0.000")
            i = i + 1
            .MoveNext
        Loop
    Dim IntParada As Integer
    IntParada = i - 1
    End With
    RstDenominacion.Close
    Set RstDenominacion = Nothing
    For i = 0 To 3
        If i <> 2 Then
            With GridTPago(i)
                .ColWidth(0) = 400
                .ColWidth(1) = 1500
                .Width = 2300
                .TextMatrix(0, 0) = "Nº": .TextMatrix(0, 1) = "MONTO"
                .TextMatrix(1, 0) = 1
            End With
        End If
    Next
    With GridTPago(2)
        .ColWidth(0) = 1700
        .ColWidth(1) = 1000
        .ColWidth(2) = 1700
        .Width = 4400
        .TextMatrix(0, 0) = "DENOMINACION": .TextMatrix(0, 1) = "CANTIDAD"
        .TextMatrix(0, 2) = "MONTO"
        .Rows = IntParada + 2
        For J = 0 To IntParada
            .TextMatrix(J + 1, 0) = vDenominacion(J)
        Next
    End With
    Label1(1) = Format(Date, "DD/MM/YYYY")
    Label1(3) = gcUsuario
End Sub
    '---------------------------------------------------------------------------------------------
    Private Sub GridTPago_EnterCell(index As Integer)   '
    '---------------------------------------------------------------------------------------------
    '
    Select Case index
        
        Case 0, 1, 3 'GridDepositos y GridCheques
            With GridTPago(index)
                If .Text <> "" Then
                    .Text = CCur(.Text)
                End If
            End With
        
        Case 2  'Grid Efectivo
    '
    End Select
    '
    End Sub

    '---------------------------------------------------------------------------------------------
    Private Sub GridTPago_KeyPress(index As Integer, KeyAscii As Integer)   '
    '---------------------------------------------------------------------------------------------
    '
    Call Validacion(KeyAscii, "0123456789.,")
    If KeyAscii = 0 Then Exit Sub
    If KeyAscii = 46 Then KeyAscii = 44 'CONVIERTE PUNTO(.) EN COMA(,)
    '
    Select Case index
    '
            Case 0, 1, 3 'Grid de Depósitos y Grid Cheques
    '       ---------------------------
                With GridTPago(index)
                    If KeyAscii = 8 Then
                        If Len(.Text) > 0 Then
                            .Text = Left(.Text, Len(.Text) - 1)
                        End If
                    
                    ElseIf KeyAscii = 13 Then
                        If .Row < .Rows - 1 Then
                            .Row = .Row + 1
                        Else
                            IntContador = .TextMatrix(.Rows - 1, 0)
                            .AddItem (IntContador + 1)
                            .Row = .Row + 1
                            If IntContador >= 14 Then .TopRow = .Rows - 14
                            
                        End If
                    Else
                        .TextMatrix(.RowSel, 1) = .TextMatrix(.RowSel, 1) & Chr(KeyAscii)
                    End If
                    
                End With
            
            Case 2  'Grid de Efectivo
    '       ---------------------------
                With GridTPago(2)
    '
                    If .Col = 1 Then
                        If KeyAscii = 8 Then
                            If Len(.Text) > 0 Then
                                .Text = Left(.Text, Len(.Text) - 1)
                            End If
                        ElseIf KeyAscii = 13 Then
                            If .Row < .Rows - 1 Then
                                .Row = .Row + 1
                            Else
                                .Row = 1
                            End If
                        Else
                            .TextMatrix(.RowSel, 1) = .TextMatrix(.RowSel, 1) & Chr(KeyAscii)
                        End If
                    End If
                End With
    '
    End Select
    '
    End Sub

    '---------------------------------------------------------------------------------------------
    Private Sub GridTPago_KeyUp(index As Integer, KeyCode As Integer, Shift As Integer) '
    '---------------------------------------------------------------------------------------------
    '
    If KeyCode = 46 Then
    '
        Select Case index
    '
            Case 0, 1, 2, 3 'Grid de Depositos y Grid de Efectivo
    '       ---------------------------
                If GridTPago(index).Col = 1 Then
                    GridTPago(index).Text = ""
                End If
    '
        End Select
    '
    End If
    '
    End Sub

    '---------------------------------------------------------------------------------------------
    Private Sub GridTPago_LeaveCell(index As Integer)   '
    '---------------------------------------------------------------------------------------------
    '
    Select Case index
    '
        Case 0, 1, 3 'Grid Depositos y Grid Cheques
    '   -------------------------------
            With GridTPago(index)
                .TextMatrix(.RowSel, 1) = Format(.TextMatrix(.RowSel, 1), "#,##0.00")
    
            Dim curMonto As Currency
            Dim Count As Integer
            For i = 1 To .Rows - 1
                If .TextMatrix(i, 1) <> "" Then
                    curMonto = _
                    curMonto + .TextMatrix(i, 1)
                    Count = Count + 1
                End If
            Next
            
            FraEntrega(index) = Count & "-" & IIf(index = 0, "Depósitos Bs. ", IIf(index = 3, "Transf. Bs. ", "Cheques Bs. ")) & Format(curMonto, "#,##0.00")
            Label1(6) = Format(CajaGeneral, "#,##0.00")
            End With
        
        Case 2  'Grid Efectivo
    '   -------------------------------
            With GridTPago(2)
    
                If .Col = 1 And (.TextMatrix(.RowSel, 1)) <> "" Then
                    .TextMatrix(.RowSel, 2) = _
                    Format(.TextMatrix(.RowSel, 1) * .TextMatrix(.RowSel, 0), "#,##0.00")
                Else
                    If (.TextMatrix(.RowSel, 1)) = "" Then .TextMatrix(.RowSel, 2) = ""
                End If
                Dim CurEfectivo As Currency
                CurEfectivo = CurSaldoOpen * -1
                For i = 1 To .Rows - 1
                    If .TextMatrix(i, 2) <> "" Then
                        CurEfectivo = CurEfectivo + .TextMatrix(i, 2)
                        Count = Count + 1
                    End If
                Next
                
                FraEntrega(2) = "Detalle Efectivo: Bs. " + Format(CurEfectivo, "#,##0.00")
                Label1(6) = Format(CajaGeneral, "#,##0.00")
            End With
    
    End Select
    '
    End Sub



Public Function CajaGeneral() As Currency
Dim Columna As Integer
CajaGeneral = CurSaldoOpen
For J = 0 To 3
With GridTPago(J)
    If J = 0 Or J = 1 Or J = 3 Then
        Columna = 1
    Else
        Columna = 2
    End If
    
    For K = 1 To .Rows - 1
        
        If .TextMatrix(K, Columna) <> "" Then
            CajaGeneral = _
            CajaGeneral + CCur(.TextMatrix(K, Columna))
        End If
        
    Next

End With
Next
End Function

    '------------------------------------------------------------------------------------
    Private Sub Label1_Change(index As Integer) '
    '------------------------------------------------------------------------------------
    '
    If index = 6 Then
        Label1(7) = Format(CCur(Label1(6)) - CCur(Label1(8)), "#,##0.00")
    End If
    '
    End Sub

