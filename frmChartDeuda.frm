VERSION 5.00
Begin VB.Form frmChartDeuda 
   Caption         =   "Estadistico Inmueble"
   ClientHeight    =   4125
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7140
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   4125
   ScaleWidth      =   7140
   Begin VB.Frame fra 
      Caption         =   "Introduzca los campos requeridos:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1605
      Index           =   0
      Left            =   255
      TabIndex        =   2
      Top             =   375
      Width           =   14430
      Begin VB.ComboBox cmb 
         Height          =   315
         Index           =   5
         Left            =   1935
         Sorted          =   -1  'True
         TabIndex        =   11
         Top             =   915
         Width           =   5070
      End
      Begin VB.ComboBox cmb 
         Height          =   315
         Index           =   0
         ItemData        =   "frmChartDeuda.frx":0000
         Left            =   510
         List            =   "frmChartDeuda.frx":0002
         TabIndex        =   7
         Top             =   915
         Width           =   1425
      End
      Begin VB.ComboBox cmb 
         Height          =   315
         Index           =   1
         ItemData        =   "frmChartDeuda.frx":0004
         Left            =   7815
         List            =   "frmChartDeuda.frx":002F
         TabIndex        =   6
         Top             =   915
         Width           =   1185
      End
      Begin VB.ComboBox cmb 
         Height          =   315
         Index           =   2
         ItemData        =   "frmChartDeuda.frx":009F
         Left            =   10605
         List            =   "frmChartDeuda.frx":00CA
         TabIndex        =   5
         Top             =   915
         Width           =   1185
      End
      Begin VB.ComboBox cmb 
         Height          =   315
         Index           =   3
         Left            =   9000
         TabIndex        =   4
         Top             =   915
         Width           =   1455
      End
      Begin VB.ComboBox cmb 
         Height          =   315
         Index           =   4
         Left            =   11790
         TabIndex        =   3
         Top             =   915
         Width           =   1455
      End
      Begin VB.Label lbl 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Código. del Inmueble"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   1
         Left            =   480
         TabIndex        =   10
         Top             =   570
         Width           =   6495
      End
      Begin VB.Label lbl 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Inicio Período"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   2
         Left            =   7830
         TabIndex        =   9
         Top             =   570
         Width           =   2595
      End
      Begin VB.Label lbl 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Final Período"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   3
         Left            =   10605
         TabIndex        =   8
         Top             =   570
         Width           =   2640
      End
   End
   Begin VB.CommandButton cmd 
      Caption         =   "&Cerrar"
      Height          =   960
      Index           =   2
      Left            =   10770
      Picture         =   "frmChartDeuda.frx":013A
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   3015
      Width           =   1215
   End
   Begin VB.CommandButton cmd 
      Caption         =   "&Imprimir"
      Height          =   960
      Index           =   1
      Left            =   9555
      Picture         =   "frmChartDeuda.frx":0444
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   3015
      Width           =   1215
   End
End
Attribute VB_Name = "frmChartDeuda"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'variables locales
Dim rstlocal As New ADODB.Recordset
Dim Inm As String
Dim curDA As Currency
Dim pIni As Date
Dim pFin As Date
Dim pPeriodo As Date


Private Sub cmb_Click(Index As Integer)
'variables locales
Dim strCriterio As String
'
If Index = 0 Then
    strCriterio = "CodInm ='" & Cmb(0) & "'"
ElseIf Index = 5 Then
    strCriterio = "Nombre Like '*" & Cmb(5) & "*'"
End If
'
If strCriterio <> "" Then
'
    With rstlocal
        .MoveFirst
        .Find strCriterio
        If Not .EOF And Not .BOF Then
            Cmb(0) = !CodInm
            Cmb(5) = !Nombre
            Inm = !CodInm
            curDA = !Deuda
        Else
            MsgBox "No se encuentra coincidencia con el criterio de búsqueda", vbInformation, _
            App.ProductName
        End If
        '
    End With
    '
End If
'

End Sub

Private Sub cmb_KeyPress(Index As Integer, KeyAscii As Integer)
'combierte todo a mayuscular
KeyAscii = Asc(UCase(Chr(KeyAscii)))
If Index = 0 Then If KeyAscii > 26 Then If InStr("0123456789", Chr(KeyAscii)) = 0 Then KeyAscii = 0
    

If KeyAscii = 13 Then
    If Index = 0 Or Index = 5 Then
    
        Call cmb_Click(Index)
    Else
        SendKeys vbTab
    End If
End If
End Sub

Private Sub cmd_Click(Index As Integer)
'variables locales
Select Case Index

    Case 0  'graficar
        'If Not novalida Then Call grafica
    
    Case 1 'imprimir
    
    Case 2  'close
        Unload Me
        Set frmChartDeuda = Nothing
        '
End Select
'
End Sub

Private Sub Form_Load()

With rstlocal
    .Open "Inmueble", cnnConexion, adOpenStatic, adLockReadOnly, adCmdTable
    .Filter = "Inactivo = False"
    If Not .EOF And Not .BOF Then
        .MoveFirst
        Do
            Cmb(0).AddItem !CodInm
            Cmb(5).AddItem !Nombre
            .MoveNext
        Loop Until .EOF
    End If
    '
    
End With
For i = 3 To 4
    For j = 2001 To Year(Date): Cmb(i).AddItem j
    Next
Next
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
rstlocal.Close
Set rstlocal = Nothing
End Sub

