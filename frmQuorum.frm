VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmQuorum 
   Caption         =   "Quórum"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   WindowState     =   2  'Maximized
   Begin VB.CheckBox Check1 
      Caption         =   "Votación 2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   8055
      TabIndex        =   30
      Top             =   5085
      Width           =   1260
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Votación 1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Index           =   1
      Left            =   8070
      TabIndex        =   29
      Top             =   3150
      Width           =   1245
   End
   Begin VB.Frame fraQ 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Index           =   3
      Left            =   7920
      TabIndex        =   16
      Top             =   5085
      Width           =   3780
      Begin VB.TextBox txt 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   5
         Left            =   1065
         TabIndex        =   18
         Text            =   "000"
         Top             =   360
         Width           =   555
      End
      Begin VB.CommandButton cmd 
         Caption         =   "Reiniciar"
         Height          =   330
         Index           =   2
         Left            =   210
         Style           =   1  'Graphical
         TabIndex        =   32
         Top             =   1290
         Width           =   1260
      End
      Begin VB.TextBox txt 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   4
         Left            =   1875
         Locked          =   -1  'True
         TabIndex        =   17
         Text            =   "0,000000"
         Top             =   360
         Width           =   1695
      End
      Begin VB.Label lbl 
         Caption         =   "Nº Votos"
         Height          =   225
         Index           =   10
         Left            =   270
         TabIndex        =   22
         Top             =   420
         Width           =   720
      End
      Begin VB.Label lbl 
         Alignment       =   2  'Center
         BackColor       =   &H80000002&
         BackStyle       =   0  'Transparent
         Caption         =   "0,00%"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000009&
         Height          =   180
         Index           =   7
         Left            =   210
         TabIndex        =   19
         Top             =   915
         Width           =   3435
      End
      Begin VB.Label lbl 
         BackColor       =   &H80000002&
         Height          =   285
         Index           =   8
         Left            =   240
         TabIndex        =   20
         Top             =   885
         Width           =   30
      End
      Begin VB.Label lbl 
         BackColor       =   &H8000000C&
         BorderStyle     =   1  'Fixed Single
         Height          =   345
         Index           =   9
         Left            =   210
         TabIndex        =   21
         Top             =   855
         Width           =   3450
      End
   End
   Begin VB.CommandButton cmd 
      Caption         =   "&Salir"
      Height          =   870
      Index           =   0
      Left            =   10395
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   6990
      Width           =   1290
   End
   Begin VB.Frame fraQ 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Index           =   2
      Left            =   7935
      TabIndex        =   8
      Top             =   3225
      Width           =   3780
      Begin VB.CommandButton cmd 
         Caption         =   "Reiniciar"
         Height          =   330
         Index           =   1
         Left            =   210
         Style           =   1  'Graphical
         TabIndex        =   31
         Top             =   1275
         Width           =   1260
      End
      Begin VB.TextBox txt 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   3
         Left            =   1875
         Locked          =   -1  'True
         TabIndex        =   10
         Text            =   "0,000000"
         Top             =   375
         Width           =   1695
      End
      Begin VB.TextBox txt 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   2
         Left            =   1065
         TabIndex        =   9
         Text            =   "000"
         Top             =   375
         Width           =   555
      End
      Begin VB.Label lbl 
         Alignment       =   2  'Center
         BackColor       =   &H80000002&
         BackStyle       =   0  'Transparent
         Caption         =   "0,00%"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000009&
         Height          =   180
         Index           =   6
         Left            =   210
         TabIndex        =   15
         Top             =   915
         Width           =   3435
      End
      Begin VB.Label lbl 
         BackColor       =   &H80000002&
         Height          =   285
         Index           =   5
         Left            =   240
         TabIndex        =   14
         Top             =   885
         Width           =   30
      End
      Begin VB.Label lbl 
         BackColor       =   &H8000000C&
         BorderStyle     =   1  'Fixed Single
         Height          =   345
         Index           =   4
         Left            =   210
         TabIndex        =   13
         Top             =   855
         Width           =   3450
      End
      Begin VB.Label lbl 
         Caption         =   "Nº Votos"
         Height          =   225
         Index           =   3
         Left            =   270
         TabIndex        =   11
         Top             =   435
         Width           =   720
      End
   End
   Begin VB.Frame fraQ 
      Caption         =   "Quórum"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1995
      Index           =   1
      Left            =   7950
      TabIndex        =   3
      Top             =   1035
      Width           =   3780
      Begin VB.TextBox txt 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   9
         Left            =   1500
         Locked          =   -1  'True
         TabIndex        =   28
         Text            =   "000"
         Top             =   1020
         Width           =   510
      End
      Begin VB.TextBox txt 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   8
         Left            =   1500
         Locked          =   -1  'True
         TabIndex        =   27
         Text            =   "0,000000"
         Top             =   1380
         Width           =   2100
      End
      Begin VB.TextBox txt 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   7
         Left            =   2385
         TabIndex        =   25
         Text            =   "0,000000"
         Top             =   660
         Width           =   1215
      End
      Begin VB.TextBox txt 
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
         Height          =   360
         Index           =   6
         Left            =   1500
         TabIndex        =   24
         Text            =   "0,00%"
         Top             =   660
         Width           =   885
      End
      Begin VB.TextBox txt 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   0
         Left            =   1500
         Locked          =   -1  'True
         TabIndex        =   5
         Text            =   "0,00000"
         Top             =   300
         Width           =   2100
      End
      Begin VB.TextBox txt 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   1
         Left            =   2010
         Locked          =   -1  'True
         TabIndex        =   4
         Text            =   "0,000000"
         Top             =   1020
         Width           =   1590
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         Caption         =   "Diferencia:"
         Height          =   225
         Index           =   12
         Left            =   135
         TabIndex        =   26
         Top             =   1455
         Width           =   1245
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         Caption         =   "Requerido:"
         Height          =   225
         Index           =   11
         Left            =   135
         TabIndex        =   23
         Top             =   735
         Width           =   1245
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         Caption         =   "Total Alícuota:"
         Height          =   225
         Index           =   1
         Left            =   135
         TabIndex        =   7
         Top             =   375
         Width           =   1245
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         Caption         =   "Asistencia:"
         Height          =   225
         Index           =   2
         Left            =   135
         TabIndex        =   6
         Top             =   1095
         Width           =   1245
      End
   End
   Begin VB.Frame fraQ 
      Caption         =   "Propietarios"
      Height          =   6855
      Index           =   0
      Left            =   225
      TabIndex        =   1
      Top             =   1035
      Width           =   7575
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid grid 
         CausesValidation=   0   'False
         Height          =   6180
         Left            =   270
         TabIndex        =   2
         Tag             =   "800|3500|1000|600"
         Top             =   405
         Width           =   7125
         _ExtentX        =   12568
         _ExtentY        =   10901
         _Version        =   393216
         Cols            =   5
         FixedCols       =   0
         BackColorFixed  =   -2147483646
         ForeColorFixed  =   -2147483639
         BackColorSel    =   16777215
         ForeColorSel    =   0
         BackColorBkg    =   -2147483636
         GridColor       =   -2147483633
         AllowBigSelection=   0   'False
         FocusRect       =   2
         GridLinesUnpopulated=   2
         AllowUserResizing=   1
         MousePointer    =   99
         FormatString    =   "Apto.|Propietario|Alícuota|Asist|Voto"
         MouseIcon       =   "frmQuorum.frx":0000
         _NumberOfBands  =   1
         _Band(0).Cols   =   5
      End
   End
   Begin VB.Image Img 
      Enabled         =   0   'False
      Height          =   150
      Index           =   0
      Left            =   0
      Top             =   15
      Visible         =   0   'False
      Width           =   165
   End
   Begin VB.Image Img 
      Enabled         =   0   'False
      Height          =   150
      Index           =   1
      Left            =   315
      Top             =   0
      Visible         =   0   'False
      Width           =   165
   End
   Begin VB.Label lbl 
      Caption         =   "Quórum"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000002&
      Height          =   660
      Index           =   0
      Left            =   195
      TabIndex        =   0
      Top             =   180
      Width           =   2115
   End
End
Attribute VB_Name = "frmQuorum"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
    '---------------------------------------------------------------------------------------------
    '   Módulo Quorum
    '
    '---------------------------------------------------------------------------------------------
    'variables locales
    Dim vecProp()
    Dim Votacion As Integer
    Const Codigo& = 0
    Const Nombre& = 2
    Const Alicuota& = 6
    '-------------
    
    Private Sub Check1_Click(Index As Integer)
    '
    Votacion = 0
    If Index = 1 Then
        If Check1(1).Value = vbChecked Then Check1(2).Value = vbUnchecked
    ElseIf Index = 2 Then
        If Check1(2).Value = vbChecked Then Check1(1).Value = vbUnchecked
    End If
    For I = 1 To 2
        If Check1(I).Value = vbChecked Then Votacion = I
    Next I
    '
    End Sub

    Private Sub cmd_Click(Index As Integer)
    Select Case Index
        Case 0
            Unload Me
            Set frmQuorum = Nothing
            
        Case 1: Call Reiniciar(3, 2, 6, 5)
        
        Case 2: Call Reiniciar(4, 5, 7, 8)
        
    End Select
    End Sub

    Private Sub Form_Load()
    
    Dim rstQ As New ADODB.Recordset
    Dim I As Integer
    Dim k As Integer
    Dim AliTotal As Double
    '
    'vecProp(Campo,Nregistro)
    Img(0).Picture = LoadResPicture("UnChecked", vbResBitmap)
    Img(1).Picture = LoadResPicture("Checked", vbResBitmap)
    cmd(0).Picture = LoadResPicture("salir", vbResIcon)
    '
    Call centra_titulo(grid, True)
    Set grid.FontFixed = LetraTitulo(LoadResString(527), 9, , True)
    Set grid.Font = LetraTitulo(LoadResString(528), 8)
    rstQ.Open "Propietarios", cnnOLEDB + mcDatos, adOpenKeyset, adLockOptimistic, adCmdTable
    vecProp = rstQ.GetRows
    rstQ.Close
    Set rstQ = Nothing
    '
    Call centra_titulo(grid, True)
    Call rtnLimpiar_Grid(grid)
    grid.Rows = UBound(vecProp, 2) + 1
    '
    For I = 1 To UBound(vecProp, 2)
    
        grid.TextMatrix(I, 0) = vecProp(Codigo, I - 1)
        grid.TextMatrix(I, 1) = IIf(IsNull(vecProp(Nombre, I - 1)), "", vecProp(Nombre, I - 1))
        grid.TextMatrix(I, 2) = Format(vecProp(Alicuota, I - 1), "0.000000")
        AliTotal = AliTotal + CDbl(vecProp(Alicuota, I - 1))
        For k = 3 To 4
            grid.Col = k
            grid.Row = I
            Set grid.CellPicture = Img(0)
            grid.CellPictureAlignment = flexAlignCenterCenter
        Next
        
    Next I
    '
    txt(0) = Format(AliTotal, "#,##0.000000")
    grid.ColAlignment(0) = flexAlignCenterCenter
    grid.Row = 1
    grid.Col = 0
    grid.ColSel = grid.Cols - 1
    '
    End Sub

    Private Sub grid_Click()
    '
    If grid.ColSel <> 4 Then
        If CCur(txt(7)) = 0 Then
            MsgBox "Falta parámetro. Quorúm requerido", vbInformation, App.ProductName
            Exit Sub
        Else
            grid.Col = 3
        End If
    Else
        grid.Col = 3
        grid.Row = grid.RowSel
        If grid.CellPicture = Img(0) Then
            MsgBox "Este propietario no pueda votar porque no está presente", vbInformation, _
            App.ProductName
            Exit Sub
        End If
        grid.Col = 4
    End If
    '
    grid.Row = grid.RowSel
    Set grid.CellPicture = IIf(grid.CellPicture = Img(0), Img(1), Img(0))
    If grid.Col = 3 Then    'quórum
        txt(1) = IIf(grid.CellPicture = Img(0), grid.TextMatrix(grid.Row, 2) * -1, _
        grid.TextMatrix(grid.Row, 2)) + CCur(txt(1))
        txt(9) = IIf(grid.CellPicture = Img(0), CLng(txt(9)) - 1, CLng(txt(9)) + 1)
        txt(8) = CCur(txt(7)) - CCur(txt(1))
        txt(1) = Format(txt(1), "#,##0.000000")
        txt(9) = Format(txt(9), "000")
        txt(8) = Format(txt(8), "#,##0.000000")
    Else    'votación
        If Votacion = 1 Then
            Call Vota(3, 2, 6, 5)
        ElseIf Votacion = 2 Then
            Call Vota(4, 5, 7, 8)
        End If
    End If
    
    End Sub


    Private Sub grid_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Or KeyAscii = Asc("A") Or KeyAscii = Asc("a") Then Call grid_Click
    If KeyAscii = Asc("V") Or KeyAscii = Asc("v") Then
            grid.Col = 4
            Call grid_Click
    End If
    grid.Col = grid.ColSel
    End Sub



    Private Sub txt_KeyPress(Index As Integer, KeyAscii As Integer)
    '
    Select Case Index
        
        Case 6  'porcentaje req. para quorum
            If KeyAscii = 46 Then KeyAscii = 44
            Call Validacion(KeyAscii, "0123456789%,")
            'si presiona enter o la tecla "%" efectua el cálculo
            If KeyAscii = 13 Or KeyAscii = Asc("%") Then
                KeyAscii = 0
                If InStr(txt(6), "%") <> 0 Then txt(6) = Left(txt(6), InStr(txt(6), "%") - 1)
                txt(7) = CCur(txt(0)) * CCur(txt(6)) / 100
                txt(7) = Format(txt(7), "#,##0.000000")
                txt(6) = Format(txt(6) / 100, "0.00%")
                txt(8) = CCur(txt(7)) - CCur(txt(1))
                txt(8) = Format(txt(8), "#,##0.000000")
            End If
            
        Case 7  'cantidad en números de pers. req. para quorum
            If KeyAscii = 44 Then KeyAscii = 46 'convierte punto a coma
            Call Validacion(KeyAscii, "0123456789,")
            If KeyAscii = 13 Then
                txt(6) = (CCur(txt(7)) * 100) / CCur(txt(0))
                txt(6) = Format(txt(6), "0%")
            End If
            
    End Select
    '
    End Sub

    Private Sub Vota(A%, B%, C%, D%)
    '
    lbl(D).Visible = False
    txt(A) = IIf(grid.CellPicture = Img(0), grid.TextMatrix(grid.Row, 2) * -1, _
    grid.TextMatrix(grid.Row, 2)) + CCur(txt(A))
    txt(B) = IIf(grid.CellPicture = Img(0), txt(B) - 1, txt(B) + 1)
    lbl(C) = (CCur(txt(A)) * 100) / CCur(txt(1))
    lbl(D).Width = IIf((lbl(C) * 3390) / 100 < 0, 0, (lbl(C) * 3390) / 100)
    'formato a los controles
    txt(A) = Format(txt(A), "#,##0.000000")
    txt(B) = Format(txt(B), "000")
    lbl(C) = Format(lbl(C) / 100, "0.00%")
    lbl(D).Visible = True
    '----------
    End Sub

    Private Sub Reiniciar(A%, B%, C%, D%)
    txt(A) = Format(0, "#,##0.000000")
    txt(B) = Format(0, "000")
    lbl(D).Width = 0
    lbl(C) = Format(0 / 100, "0.00%")
    End Sub
