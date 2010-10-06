VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmGasMen 
   Caption         =   "Gastos Mensuales"
   ClientHeight    =   8310
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9885
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8310
   ScaleWidth      =   9885
   Tag             =   "1"
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmd 
      Caption         =   "Salir"
      Height          =   765
      Index           =   0
      Left            =   6375
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   6405
      Width           =   1065
   End
   Begin VB.CommandButton cmd 
      Caption         =   "Imprimir"
      Height          =   765
      Index           =   1
      Left            =   5160
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   6405
      Width           =   1065
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid Flex 
      Height          =   4800
      Left            =   795
      TabIndex        =   0
      Top             =   945
      Width           =   6615
      _ExtentX        =   11668
      _ExtentY        =   8467
      _Version        =   393216
      Cols            =   5
      FixedCols       =   0
      RowHeightMin    =   315
      BackColorFixed  =   -2147483646
      ForeColorFixed  =   -2147483639
      BackColorSel    =   65280
      ForeColorSel    =   0
      BackColorBkg    =   -2147483636
      GridColor       =   -2147483633
      AllowBigSelection=   0   'False
      AllowUserResizing=   1
      FormatString    =   "^Cod.Gasto|<Descripción|>Monto|^Nº Cheque|^Verif."
      BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   5
   End
   Begin MSAdodcLib.Adodc aDO 
      Height          =   330
      Left            =   1365
      Top             =   6495
      Visible         =   0   'False
      Width           =   2790
      _ExtentX        =   4921
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Gastos Mensuales"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.Label lbl 
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   1
      Left            =   1920
      TabIndex        =   2
      Top             =   495
      Width           =   1215
   End
   Begin VB.Label lbl 
      Caption         =   "Período:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   0
      Left            =   840
      TabIndex        =   1
      Top             =   450
      Width           =   1215
   End
   Begin VB.Image img 
      Enabled         =   0   'False
      Height          =   255
      Index           =   1
      Left            =   705
      Picture         =   "frmGasMen.frx":0000
      Stretch         =   -1  'True
      Top             =   390
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image img 
      Enabled         =   0   'False
      Height          =   180
      Index           =   0
      Left            =   735
      Picture         =   "frmGasMen.frx":0083
      Stretch         =   -1  'True
      Top             =   135
      Visible         =   0   'False
      Width           =   180
   End
End
Attribute VB_Name = "frmGasMen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public OrigenD As String

Private Sub Cmd_Click(Index As Integer)
'
Select Case Index
    '
    Case 0
    
        Unload Me
        Set frmGasMen = Nothing
        '
    Case 1
        Call FrmAdmin.Reporte_GM(, crPantalla)
        '
End Select
End Sub

Private Sub Flex_Click()
If Flex.RowSel >= 1 And Flex.ColSel = 4 And Flex.CellPicture <> 0 Then Marca
End Sub

Private Sub Flex_KeyPress(KeyAscii As Integer)
If Flex.RowSel >= 1 And Flex.ColSel = 4 And Flex.CellPicture <> 0 And KeyAscii = 13 Then Marca
End Sub

Private Sub Form_Load()
'variables locales
Flex.Tag = "1500|8000|1500|1500|500"
Call centra_titulo(Flex, True)
'
With Ado
    .CursorLocation = adUseClient
    .ConnectionString = cnnConexion.ConnectionString
    .LockType = adLockOptimistic
    .CommandType = adCmdText
    .RecordSource = OrigenD
    .Refresh
End With
'
Call Mostrar
Unload FrmRepFact
cmd(0).Picture = LoadResPicture("Salir", vbResIcon)
cmd(1).Picture = LoadResPicture("Print", vbResIcon)
'
End Sub

Private Sub Form_Resize()
Const Distancia& = 200
If FrmAdmin.WindowState <> vbMinimized Then
    Flex.Width = ScaleWidth - (Flex.Left * 2)
    Flex.Height = ScaleHeight - Flex.Top - cmd(0).Height - (Distancia * 2)
    cmd(0).Top = Flex.Height + Flex.Top + Distancia
    cmd(0).Left = Flex.Left + Flex.Width - cmd(0).Width
    cmd(1).Top = cmd(0).Top
    cmd(1).Left = cmd(0).Left - Distancia - cmd(1).Width
End If
End Sub

Private Sub Mostrar()
'variables locales
Dim I As Long, N As Long
Dim Gasto As String
Dim subTotal As Double
'
Call rtnLimpiar_Grid(Flex)
With Ado.Recordset
    If Not .EOF And Not .BOF Then
        .Sort = "CodGasto"
        .MoveFirst
        Lbl(1) = UCase(Format(!Cargado, "MMM-YYYY"))
        Flex.Rows = .RecordCount + 1
        Do
            I = I + 1
            Flex.TextMatrix(I, 0) = !codGasto
            Flex.TextMatrix(I, 1) = !Detalle
            Flex.TextMatrix(I, 2) = Format(!Monto, "#,##0.00 ")
            Flex.TextMatrix(I, 3) = Format(!IDCheque, "000000")
            Flex.Row = I
            Flex.Col = Flex.Cols - 1
            If !check Then
                Set Flex.CellPicture = img(0)
            Else
                Set Flex.CellPicture = img(1)
            End If
            Flex.CellPictureAlignment = flexAlignCenterCenter
            If Gasto = "" Then
                subTotal = !Monto
                Gasto = !codGasto
                N = 1
            End If
            .MoveNext
            If Not .EOF Then
                If !codGasto = Gasto Then
                    subTotal = subTotal + !Monto
                    N = N + 1
                Else
                    
                    If N > 1 Then
                        
                        I = I + 1
                        Flex.AddItem ("")
                        Flex.TextMatrix(I, 1) = "TOTAL GASTO " & Gasto
                        Flex.Col = 1
                        Flex.Row = I
                        Flex.CellAlignment = flexAlignRightCenter
                        Flex.CellFontBold = True
                        Flex.TextMatrix(I, 2) = Format(subTotal, "#,##0.00 ")
                        Flex.Col = 2
                        Flex.CellAlignment = flexAlignRightCenter
                        Flex.CellFontBold = True
                        Flex.CellBackColor = vbGreen
                    End If
                    N = 1
                    subTotal = !Monto
                    Gasto = !codGasto
                End If
            End If
        Loop Until .EOF
    End If
End With
End Sub

Private Sub Marca()
'variables locales
Dim Criterio As String
Dim Col&, Fila&

Fila = Flex.RowSel
'
Criterio = "IDCheque =" & Flex.TextMatrix(Fila, 3) & " AND CodGasto ='" & _
Flex.TextMatrix(Fila, 0) & "' AND Detalle ='" & Flex.TextMatrix(Fila, 1) & _
"' AND Monto = " & Replace(CCur(Flex.TextMatrix(Fila, 2)), ",", ".")
'
With Ado.Recordset
    .Filter = Criterio
    If Not .EOF Then
    
        Flex.Col = Flex.ColSel
        Flex.Row = Fila
        Set Flex.CellPicture = img(IIf(Flex.CellPicture = img(1), 0, 1))
        Flex.CellPictureAlignment = flexAlignCenterCenter
        .Update "Check", IIf(Flex.CellPicture = img(0), -1, 0)
    Else
        .Filter = adFilterNone
    End If
End With

End Sub

