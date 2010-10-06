VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmCajaCta 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Form1"
   ClientHeight    =   2040
   ClientLeft      =   2130
   ClientTop       =   5505
   ClientWidth     =   5520
   ControlBox      =   0   'False
   DrawMode        =   1  'Blackness
   DrawStyle       =   2  'Dot
   FillStyle       =   6  'Cross
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2040
   ScaleWidth      =   5520
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdCta 
      Appearance      =   0  'Flat
      Caption         =   "&Cancelar"
      Height          =   435
      Index           =   1
      Left            =   3930
      MaskColor       =   &H00FFFFFF&
      TabIndex        =   2
      Top             =   1485
      Width           =   1440
   End
   Begin VB.CommandButton cmdCta 
      Appearance      =   0  'Flat
      Caption         =   "&Aceptar"
      Height          =   435
      Index           =   0
      Left            =   2385
      MaskColor       =   &H00FFFFFF&
      TabIndex        =   1
      Top             =   1500
      Width           =   1440
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid gridCta 
      Height          =   1230
      Left            =   120
      TabIndex        =   0
      Tag             =   "300|2000|2000|300"
      Top             =   135
      Width           =   5265
      _ExtentX        =   9287
      _ExtentY        =   2170
      _Version        =   393216
      Rows            =   4
      Cols            =   4
      FixedCols       =   0
      BackColorFixed  =   -2147483646
      ForeColorFixed  =   -2147483639
      BackColorBkg    =   -2147483636
      GridColor       =   -2147483633
      GridLinesFixed  =   0
      BorderStyle     =   0
      Appearance      =   0
      FormatString    =   "ID|Nº Cuenta|Banco|Sel"
      _NumberOfBands  =   1
      _Band(0).Cols   =   4
   End
   Begin VB.Image Img 
      Height          =   195
      Index           =   0
      Left            =   525
      Stretch         =   -1  'True
      Top             =   1110
      Width           =   195
   End
   Begin VB.Image Img 
      Height          =   195
      Index           =   1
      Left            =   135
      Stretch         =   -1  'True
      Top             =   150
      Width           =   195
   End
End
Attribute VB_Name = "frmCajaCta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public strTitulo$
Dim Marca As Boolean

Private Sub cmdCta_Click(Index As Integer)
Marca = False
Select Case Index
    Case 1: strTitulo = "": Me.Hide
    Case 0: If ftnCuenta = False Then Me.Hide
End Select
End Sub

Private Sub Form_Load()
Img(0).Picture = LoadResPicture("Unchecked", vbResBitmap)
Img(1).Picture = LoadResPicture("Checked", vbResBitmap)
Me.Caption = strTitulo
Call centra_titulo(gridCta, True)
Call llenar_grid
Me.Show vbModal, FrmAdmin
End Sub

Private Sub llenar_grid()
'On Error Resume Next
Dim i%, j%, K% 'variables locales
With gridCta
    j = 1
    K = UBound(Matriz_A, 2)
    .Rows = K + 2
    For i = 0 To K
         .TextMatrix(j, 0) = Matriz_A(0, i)
         .TextMatrix(j, 1) = Matriz_A(1, i)
         .TextMatrix(j, 2) = Matriz_A(2, i)
         .Row = j
         .Col = 3
         Set .CellPicture = Img(0)
         .CellPictureAlignment = flexAlignCenterCenter
         j = j + 1
    Next i
End With
End Sub

'-------------------------------------------------------------------------------------------------
'   Función:    ftnCuenta
'
'   Devuelte el identificador de la cuenta seleccionada por el usuario
'-------------------------------------------------------------------------------------------------
Private Function ftnCuenta() As Boolean
Dim i%, m% 'variables locales

With gridCta
    .Col = 3
    For i = 1 To .Rows - 1
        .Row = i
        If .CellPicture = Img(1) Then
            m = m + 1
            If m > 1 Then
                ftnCuenta = MsgBox("Tiene seleccionada más de una cuenta.", vbExclamation, _
                App.ProductName)
                FrmMovCaja.NCta = ""
                strTitulo = ""
                Exit Function
            End If
            FrmMovCaja.NCta = Trim(.TextMatrix(i, 0))
            strTitulo = .TextMatrix(i, 2)
            Exit Function
        End If
    Next i
    If m = 0 Then
        ftnCuenta = MsgBox("Debe seleccionar por lo menos una cuenta", vbInformation, _
        App.ProductName)
    End If
End With
End Function

Private Sub gridCta_Click()
'
With gridCta
    If .RowSel >= 1 And .ColSel = 3 Then
        If .CellPicture = Img(0) Then
            If Marca Then
                MsgBox "Imposible marcar otra cuenta, si no desmarca la anterior...", _
                vbInformation, App.ProductName
                Exit Sub
            Else
                Set .CellPicture = Img(1)
                Marca = True
            End If
        Else
            Set .CellPicture = Img(0)
            Marca = False
        End If
            .CellPictureAlignment = flexAlignCenterCenter
    End If
End With
End Sub

Private Sub gridCta_KeyPress(KeyAscii As Integer): Call gridCta_Click
End Sub
