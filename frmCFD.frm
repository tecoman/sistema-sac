VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmCFD 
   Caption         =   "Cuadre Fondo - Deuda"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdCFD 
      Caption         =   "&Recuperar"
      Height          =   450
      Index           =   4
      Left            =   10785
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   1485
      Width           =   945
   End
   Begin VB.CommandButton cmdCFD 
      Caption         =   "&Guardar"
      Height          =   450
      Index           =   3
      Left            =   10785
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   975
      Width           =   945
   End
   Begin VB.CommandButton cmdCFD 
      Caption         =   "&Actualizar"
      Height          =   450
      Index           =   2
      Left            =   10785
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   465
      Width           =   945
   End
   Begin VB.CommandButton cmdCFD 
      Caption         =   "&Salir"
      Height          =   765
      Index           =   1
      Left            =   10800
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   6840
      Width           =   1005
   End
   Begin VB.CommandButton cmdCFD 
      Caption         =   "&Imprimir"
      Height          =   765
      Index           =   0
      Left            =   9720
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   6840
      Width           =   1005
   End
   Begin VB.Frame fraCFD 
      Height          =   6615
      Index           =   2
      Left            =   105
      TabIndex        =   14
      Top             =   120
      Visible         =   0   'False
      Width           =   11700
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid gridCFD 
         Height          =   6015
         Index           =   2
         Left            =   240
         TabIndex        =   15
         Top             =   360
         Width           =   10380
         _ExtentX        =   18309
         _ExtentY        =   10610
         _Version        =   393216
         BackColor       =   16777215
         FixedCols       =   0
         BackColorFixed  =   -2147483646
         ForeColorFixed  =   -2147483639
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
   End
   Begin VB.Frame fraCFD 
      Height          =   6615
      Index           =   1
      Left            =   105
      TabIndex        =   16
      Top             =   120
      Visible         =   0   'False
      Width           =   11700
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid gridCFD 
         Height          =   6015
         Index           =   1
         Left            =   240
         TabIndex        =   17
         Top             =   360
         Width           =   10275
         _ExtentX        =   18124
         _ExtentY        =   10610
         _Version        =   393216
         BackColor       =   16777215
         FixedCols       =   0
         BackColorFixed  =   -2147483646
         ForeColorFixed  =   -2147483639
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
   End
   Begin VB.Label lblCFD 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "0,00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   7
      Left            =   7920
      TabIndex        =   13
      Top             =   7200
      Width           =   1560
   End
   Begin VB.Label lblCFD 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "0,00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   4
      Left            =   4800
      TabIndex        =   12
      Top             =   7200
      Width           =   1560
   End
   Begin VB.Label lblCFD 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "0,00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   3
      Left            =   3240
      TabIndex        =   11
      Top             =   7200
      Width           =   1560
   End
   Begin VB.Label lblCFD 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "0,00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   2
      Left            =   1680
      TabIndex        =   10
      Top             =   7200
      Width           =   1560
   End
   Begin VB.Label lblCFD 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "0,00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   11
      Left            =   6360
      TabIndex        =   9
      Top             =   7200
      Width           =   1560
   End
   Begin VB.Label lblCFD 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "0,00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   12
      Left            =   120
      TabIndex        =   8
      Top             =   7200
      Width           =   1560
   End
   Begin VB.Label lblCFD 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000002&
      Caption         =   "Diferencia"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000009&
      Height          =   255
      Index           =   6
      Left            =   7995
      TabIndex        =   7
      Top             =   6960
      Width           =   1560
   End
   Begin VB.Label lblCFD 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000002&
      Caption         =   "Facturacion"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000009&
      Height          =   255
      Index           =   5
      Left            =   4875
      TabIndex        =   6
      Top             =   6960
      Width           =   1560
   End
   Begin VB.Label lblCFD 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000002&
      Caption         =   "Cheq.Dev."
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000009&
      Height          =   255
      Index           =   22
      Left            =   3315
      TabIndex        =   5
      Top             =   6960
      Width           =   1560
   End
   Begin VB.Label lblCFD 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000002&
      Caption         =   "Cajas"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000009&
      Height          =   255
      Index           =   21
      Left            =   1755
      TabIndex        =   4
      Top             =   6960
      Width           =   1560
   End
   Begin VB.Label lblCFD 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000002&
      Caption         =   "Deuda Actual"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000009&
      Height          =   255
      Index           =   20
      Left            =   6435
      TabIndex        =   3
      Top             =   6960
      Width           =   1560
   End
   Begin VB.Label lblCFD 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000002&
      Caption         =   "Deuda Inicial"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000009&
      Height          =   255
      Index           =   1
      Left            =   195
      TabIndex        =   2
      Top             =   6960
      Width           =   1560
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H80000018&
      BackStyle       =   1  'Opaque
      BorderStyle     =   6  'Inside Solid
      BorderWidth     =   2
      DrawMode        =   15  'Merge Pen Not
      FillColor       =   &H00FFFFFF&
      Height          =   765
      Left            =   120
      Shape           =   4  'Rounded Rectangle
      Top             =   6840
      Width           =   9495
   End
End
Attribute VB_Name = "frmCFD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Caso As Integer
Const Deuda% = 1
Const Fondo% = 2
Const Moneda$ = "#,##0.00 ;(#,##0.00) "
Dim CFD As New STCFD001 'crea una instancia de la Dll generada \WINDOWS\system\STCFD001.dll


Private Sub cmdCFD_Click(Index As Integer)
Dim errlocal As Long    'variables locales
Dim i As Integer
'-------
Select Case Index
    'Imprime el reporte de cuadre
    Case 0
        MousePointer = vbHourglass
        CFD.Print_Cuadre (Caso)
        mcTitulo = IIf(Caso = Deuda, "Cuadre Deuda al ", "Cuadre Fondo al ") & Date
        mcReport = IIf(Caso = Deuda, "cfd_deuda.Rpt", "cfd_fondo.rpt")
        mcCrit = ""
        FrmReport.Show
        MousePointer = vbDefault
        
    Case 1: Unload Me   'Cierra el formulario
    
    'Guardar un temporal con la información introducida hasta ahora
    Case 3: CFD.Guardar_Inf (Caso)
    
    'Recupera la información guarda en el temporal
    Case 4: CFD.Recuperar_Inf (Caso)
    With gridCFD(Caso)
        .Col = 2
        For i = 1 To .Rows - 1
            .Row = i
            If .Text <> "" Then Call gridCFD_LeaveCell(Caso)
        Next
    End With
    
    'Refresca la presentación de los datos
    Case 2: errlocal = CFD.Config_Grid(Caso, lblCFD(12), lblCFD(11), lblCFD(7))
    
End Select
'
End Sub

Private Sub Form_Load()
'variables locales
Dim errlocal&, i%

cmdCFD(1).Picture = LoadResPicture("Salir", vbResIcon)
cmdCFD(0).Picture = LoadResPicture("Print", vbResIcon)
CFD.Directorio = gcPath & "\"
CFD.Cadena_Conexion = cnnOLEDB + gcPath + "\sac.mdb"
Set CFD.Grid = gridCFD(Caso)
fraCFD(Caso).Visible = True
For i = 1 To 2
    With gridCFD(i)
        Set .FontFixed = LetraTitulo(LoadResString(527), 7, True)
        Set .Font = LetraTitulo(LoadResString(528), 8)
    End With
Next i
errlocal = CFD.Config_Grid(Caso, lblCFD(12), lblCFD(11), lblCFD(7))
If Caso = Deuda Then
    Me.Caption = "Cuadre Deuda"
    lblCFD(1) = "Deuda Inicial"
    lblCFD(22) = "Cheq. Dev."
    lblCFD(5) = "Facturación"
    lblCFD(20) = "Deuda Actual"
Else
    Me.Caption = "Cuadre Fondo"
    lblCFD(1) = "Fondo Inicial"
    lblCFD(22) = "Egresos"
    lblCFD(5) = "Ingresos"
    lblCFD(20) = "Fondo Actual"
    lblCFD(2).Enabled = False
End If
Me.Show
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set CFD = Nothing
Set frmCFD = Nothing
End Sub


Private Sub gridCFD_EnterCell(Index As Integer)
'Ocurre cuando una celda toma el foco
'Select Case Index
    'Case 1
        Select Case gridCFD(1).Col
            Case 2, 3, 4: If Not gridCFD(Index).Text = "" Then gridCFD(Index).Text = CCur(gridCFD(Index).Text)
        End Select
'End Select
End Sub

Private Sub gridCFD_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
'Ocurre cuando el usuario suelta determinada tecla
'Lo utilizaremos para borrar un caracter dentro de la celda activa
'y para eliminar el contenido todal de la celda activa

Select Case gridCFD(Index).Col
    Case 2, 3, 4
    If KeyCode = 8 And gridCFD(Index) <> "" Then 'Borra el último elemento introducido
        gridCFD(Index).Text = Left(gridCFD(Index).Text, Len(gridCFD(Index).Text) - 1)
    ElseIf KeyCode = 46 Then    'Borrar todo el contenido
        gridCFD(Index).Text = ""
    ElseIf Shift = 2 And KeyCode = 67 Then  'Copiar
        Clipboard.Clear
        Clipboard.SetText gridCFD(Index).Text
    ElseIf Shift = 2 And KeyCode = 86 Then  'Pegar
        gridCFD(Index).Text = Clipboard.GetText
        With gridCFD(Index)
            .FillStyle = flexFillRepeat
            .Text = Clipboard.GetText
            .FillStyle = flexFillSingle
        End With
    ElseIf Shift = 2 And KeyCode = 88 Then  'Cortar
        Clipboard.SetText gridCFD(Index).Text
        gridCFD(Index).Text = ""
    End If
End Select
'
End Sub

Private Sub gridCFD_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii = Asc(".") Then KeyAscii = Asc(",")
Call Validacion(KeyAscii, "=-0123456789,+")
If KeyAscii = 13 Then
    Call gridCFD_LeaveCell(Index)
    gridCFD(Index).Row = IIf(gridCFD(Index).Row + 1 > gridCFD(Index).Rows - 1, 1, gridCFD(Index).Row + 1)
    Call gridCFD_EnterCell(Index)
End If
If KeyAscii = 0 Or KeyAscii < 26 Then Exit Sub
Select Case gridCFD(Index).ColSel
    Case 2, 3, 4: gridCFD(Index).Text = gridCFD(Index).Text + Chr(KeyAscii)
End Select

End Sub

Private Sub gridCFD_LeaveCell(Index As Integer)
'Ocurre cuando un celda pierde el foco
Dim Temp As String
Dim PosSig As Integer
Dim Positivo As Boolean
Dim Total As Currency

If Left(gridCFD(Index).Text, 1) = "=" Then   'EFECTUA SUMA
    Temp = Right(gridCFD(Index).Text, Len(gridCFD(Index)) - 1)
    Do
        PosSig = InStr(Temp, "+")
        If Not PosSig = 0 Then
            Total = Total + Mid(Temp, 1, PosSig - 1)
            Temp = Mid(Temp, PosSig + 1, Len(Temp))
        Else
            If Temp <> "" Then
                Total = Total + Temp
                Temp = ""
            End If
        End If
    Loop Until Temp = ""
    gridCFD(Index).Text = Total
End If
If Not IsNumeric(gridCFD(Index).Text) Then gridCFD(Index).Text = ""
'Select Case Index
    'Case 1

    Select Case gridCFD(Index).Col
        Case 2, 3, 4
            CFD.Sumatoria gridCFD(Index).Col, gridCFD(Index).Row, lblCFD(gridCFD(Index).Col)
            gridCFD(Index).Text = Format(gridCFD(Index).Text, Moneda)
            lblCFD(7) = Format(CCur(lblCFD(12)) + CCur(lblCFD(2)) + CCur(lblCFD(4)) - CCur(lblCFD(11) - CCur(lblCFD(3))), Moneda)
    End Select
'End Select
'
End Sub

