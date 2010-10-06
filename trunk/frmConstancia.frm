VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmConstancia 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Constancia de Trabajo"
   ClientHeight    =   5220
   ClientLeft      =   5880
   ClientTop       =   1965
   ClientWidth     =   6525
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5220
   ScaleWidth      =   6525
   Begin VB.Frame fra 
      Caption         =   "Tipo Reporte:"
      Height          =   1215
      Index           =   2
      Left            =   1920
      TabIndex        =   10
      Top             =   3840
      Width           =   1455
      Begin VB.OptionButton opt 
         Caption         =   "Con Sueldo"
         Height          =   495
         Index           =   3
         Left            =   120
         TabIndex        =   12
         Top             =   240
         Value           =   -1  'True
         Width           =   1215
      End
      Begin VB.OptionButton opt 
         Caption         =   "Sin Sueldo"
         Height          =   375
         Index           =   2
         Left            =   120
         TabIndex        =   11
         Top             =   720
         Width           =   1215
      End
   End
   Begin VB.Frame fra 
      Caption         =   "Salida:"
      Height          =   1215
      Index           =   1
      Left            =   240
      TabIndex        =   7
      Top             =   3840
      Width           =   1455
      Begin VB.OptionButton opt 
         Caption         =   "Impresora"
         Height          =   375
         Index           =   1
         Left            =   120
         TabIndex        =   9
         Top             =   720
         Width           =   1215
      End
      Begin VB.OptionButton opt 
         Caption         =   "Ventana"
         Height          =   495
         Index           =   0
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Value           =   -1  'True
         Width           =   1215
      End
   End
   Begin VB.CommandButton cmd 
      Caption         =   "&Imprimir"
      Height          =   375
      Index           =   2
      Left            =   4080
      TabIndex        =   6
      Top             =   4680
      Width           =   1095
   End
   Begin MSComctlLib.ListView lst 
      Height          =   1935
      Left            =   240
      TabIndex        =   5
      Top             =   1800
      Width           =   6135
      _ExtentX        =   10821
      _ExtentY        =   3413
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      HideColumnHeaders=   -1  'True
      AllowReorder    =   -1  'True
      Checkboxes      =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      NumItems        =   6
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Nombre"
         Object.Width           =   6068
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   1
         Text            =   "Cédula"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Sueldo"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Inm"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Estado"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Codigo"
         Object.Width           =   0
      EndProperty
   End
   Begin VB.CommandButton cmd 
      Caption         =   "&Salir"
      Height          =   375
      Index           =   1
      Left            =   5280
      TabIndex        =   4
      Top             =   4680
      Width           =   1095
   End
   Begin VB.Frame fra 
      Caption         =   "Buscar Empleado:"
      Height          =   1335
      Index           =   0
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   6135
      Begin VB.CommandButton cmd 
         Caption         =   "&Buscar"
         Height          =   375
         Index           =   0
         Left            =   4800
         TabIndex        =   3
         Top             =   840
         Width           =   1095
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Left            =   1920
         TabIndex        =   2
         Top             =   360
         Width           =   3975
      End
      Begin VB.ComboBox cmb 
         Height          =   315
         ItemData        =   "frmConstancia.frx":0000
         Left            =   240
         List            =   "frmConstancia.frx":000A
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   360
         Width           =   1695
      End
   End
End
Attribute VB_Name = "frmConstancia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmd_Click(Index As Integer)
Select Case Index
    Case 1
        Unload Me
        Set frmConstancia = Nothing
    
    Case 0
        Call BuscaEmpleado
        
    Case 2
        Call ImprimeReportes
        
End Select
End Sub

Private Sub Form_Load()
cmb.ListIndex = 0   'seleccion por defecto (cedula)

End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii))) 'Convierte en Mayuscula
If KeyAscii <> 13 Then
    If cmb.ListIndex = 0 Then
        Call Validacion(KeyAscii, "0123456789")
    Else
        Call Validacion(KeyAscii, " ABCDEFGHIJKLMNÑOPQRSTUWXYZ")
    End If
Else
    'BUSQUEDA EMPLEADO
    Call BuscaEmpleado
End If
End Sub


Private Sub BuscaEmpleado()
Dim strSQL As String
Dim Criterio As String
Dim rstlocal As ADODB.Recordset

If cmb.ListIndex = 0 Then
    Criterio = "WHERE Cedula =" & Text1.Text
Else
    Criterio = "WHERE Nombres Like '%" & Text1.Text & _
    "%' or Apellidos like '%" & Text1.Text & "%'"
End If
strSQL = "SELECT Emp.Nombres,Emp.Apellidos,Emp.Cedula," & _
    "Emp_Estado.Estado,Emp.Sueldo,Emp.CodInm,Emp.CodEmp " & _
    "FROM Emp " & _
    "INNER JOIN Emp_Estado On Emp_Estado.CodEstado = Emp.CodEstado " & Criterio

Set rstlocal = New ADODB.Recordset

rstlocal.Open strSQL, cnnConexion, adOpenDynamic, adLockOptimistic, adCmdText

lst.ListItems.Clear

If Not (rstlocal.EOF And rstlocal.BOF) Then
    lst.Checkboxes = True
    Do
        I = I + 1
        lst.ListItems.Add , , rstlocal("Nombres") & ", " & rstlocal("Apellidos")
        lst.ListItems(I).ListSubItems.Add , , Format(rstlocal("Cedula"), "#,##0")
        lst.ListItems(I).ListSubItems.Add , , rstlocal("sueldo")
        lst.ListItems(I).ListSubItems.Add , , rstlocal("codinm")
        lst.ListItems(I).ListSubItems.Add , , rstlocal("Estado")
        lst.ListItems(I).ListSubItems.Add , , rstlocal("CodEmp")
        rstlocal.MoveNext
    Loop Until rstlocal.EOF
Else
    lst.Checkboxes = False
    lst.ListItems.Add , , "---NO EXISTE COINCIDENCIA---"
End If
rstlocal.Close
Set rstlocal = Nothing

End Sub


Private Sub ImprimeReportes()
Dim Codigo As Long
Dim Indice As Integer
Dim conSueldo As Boolean
Dim Salida As crSalida
Dim Sueldo As Double

conSueldo = opt(3).Value
Salida = IIf(opt(0), crPantalla, crImpresora)

With lst

    For Each Item In lst.ListItems
        If Item.Checked Then
        
            Indice = Item.Index
            Codigo = CLng(.ListItems(Indice).ListSubItems(5).Text)
            Sueldo = CDbl(.ListItems(Indice).ListSubItems(2).Text)
            Call ImprimirConstancia(Codigo, conSueldo, Salida, Item.Text, Sueldo)
        
        End If
    Next
    
End With

End Sub
