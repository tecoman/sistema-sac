VERSION 5.00
Object = "{65E121D4-0C60-11D2-A9FC-0000F8754DA1}#2.0#0"; "mschrt20.ocx"
Begin VB.Form frmChartDeuda 
   Caption         =   "Estadistico Inmueble"
   ClientHeight    =   4125
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7140
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   4125
   ScaleWidth      =   7140
   WindowState     =   2  'Maximized
   Begin VB.Frame fra 
      Height          =   6600
      Index           =   1
      Left            =   240
      TabIndex        =   12
      Top             =   2220
      Width           =   14430
      Begin MSChart20Lib.MSChart MSChart1 
         Height          =   5745
         Left            =   345
         OleObjectBlob   =   "frmChartDeuda.frx":0000
         TabIndex        =   14
         Top             =   420
         Width           =   13500
      End
   End
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
      TabIndex        =   3
      Top             =   375
      Width           =   14430
      Begin VB.ComboBox cmb 
         Height          =   315
         Index           =   5
         Left            =   1935
         Sorted          =   -1  'True
         TabIndex        =   13
         Top             =   915
         Width           =   5070
      End
      Begin VB.ComboBox cmb 
         Height          =   315
         Index           =   0
         ItemData        =   "frmChartDeuda.frx":2356
         Left            =   510
         List            =   "frmChartDeuda.frx":2358
         TabIndex        =   8
         Top             =   915
         Width           =   1425
      End
      Begin VB.ComboBox cmb 
         Height          =   315
         Index           =   1
         ItemData        =   "frmChartDeuda.frx":235A
         Left            =   7815
         List            =   "frmChartDeuda.frx":2385
         TabIndex        =   7
         Top             =   915
         Width           =   1185
      End
      Begin VB.ComboBox cmb 
         Height          =   315
         Index           =   2
         ItemData        =   "frmChartDeuda.frx":23F5
         Left            =   10605
         List            =   "frmChartDeuda.frx":2420
         TabIndex        =   6
         Top             =   915
         Width           =   1185
      End
      Begin VB.ComboBox cmb 
         Height          =   315
         Index           =   3
         Left            =   9000
         TabIndex        =   5
         Top             =   915
         Width           =   1455
      End
      Begin VB.ComboBox cmb 
         Height          =   315
         Index           =   4
         Left            =   11790
         TabIndex        =   4
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
         TabIndex        =   11
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
         TabIndex        =   10
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
         TabIndex        =   9
         Top             =   570
         Width           =   2640
      End
   End
   Begin VB.CommandButton cmd 
      Caption         =   "&Cerrar"
      Height          =   960
      Index           =   2
      Left            =   13455
      Picture         =   "frmChartDeuda.frx":2490
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   8970
      Width           =   1215
   End
   Begin VB.CommandButton cmd 
      Caption         =   "&Imprimir"
      Height          =   960
      Index           =   1
      Left            =   12240
      Picture         =   "frmChartDeuda.frx":279A
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   8970
      Width           =   1215
   End
   Begin VB.CommandButton cmd 
      Caption         =   "&Graficar"
      Height          =   960
      Index           =   0
      Left            =   11025
      Picture         =   "frmChartDeuda.frx":2AA4
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   8970
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
    strCriterio = "CodInm ='" & cmb(0) & "'"
ElseIf Index = 5 Then
    strCriterio = "Nombre Like '*" & cmb(5) & "*'"
End If
'
If strCriterio <> "" Then
'
    With rstlocal
        .MoveFirst
        .Find strCriterio
        If Not .EOF And Not .BOF Then
            cmb(0) = !CodInm
            cmb(5) = !Nombre
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
        If Not novalida Then Call grafica
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
            cmb(0).AddItem !CodInm
            cmb(5).AddItem !Nombre
            .MoveNext
        Loop Until .EOF
    End If
    '
    
End With
For I = 3 To 4
    For j = 2001 To Year(Date): cmb(I).AddItem j
    Next
Next
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
rstlocal.Close
Set rstlocal = Nothing
End Sub

Private Sub grafica()
'variables locales
Dim rstGraf As New ADODB.Recordset
Dim cnnGraf As New ADODB.Connection
Dim curIni As Currency
Dim curPago As Currency
Dim curFact As Currency
Dim fIni As Date
Dim fFin As Date
Dim strSQL As String

'limpia el contenido de la tabla chartDeuda
cnnGraf.Open cnnOLEDB & gcPath & "\" & Inm & "\inm.mdb"

cnnGraf.Execute "DELETE * FROM ChartDeuda"
'
'comienza a efectuar los calculos
'pagos recibidos entre la fecha Inicial y la fecha actual
strSQL = "SELECT Sum(P.Monto) AS Pagos FROM MovimientoCaja as C INNER JOIN Periodos as P ON" _
& " C.IDRecibo = P.IDRecibo WHERE (((P.CodGasto)='900030' Or (P.CodGasto)='900001') AND ((C.Fec" _
& "haMovimientoCaja) Between #" & Format(DateAdd("d", 1, pIni), "mm/dd/yy") & "# And #" & _
Format(Date, "mm/dd/yy") & "#) AND ((C.InmuebleMovimientoCaja)='" & Inm & "'));"

rstGraf.Open strSQL, cnnConexion, adOpenKeyset, adLockOptimistic, adCmdText
If Not rstGraf.BOF And Not rstGraf.EOF Then
    If Not IsNull(rstGraf("Pagos")) Then curPago = rstGraf("Pagos")
End If
rstGraf.Close
'
'calculo de lo facturado entre la fecha inicial hasta la fecha

strSQL = "SELECT Sum(Facturado) AS Factu FROM Factura WHERE FechaFactura >#" & _
Format(pIni, "mm/dd/yy") & "#"

rstGraf.Open strSQL, cnnGraf, adOpenKeyset, adLockOptimistic, adCmdText
If Not rstGraf.EOF And Not rstGraf.BOF Then
    If Not IsNull(rstGraf("Factu")) Then curFact = rstGraf("Factu")
End If
rstGraf.Close
'
'// tengo la deuda al inicio del periodo
curIni = curDA + curPago - curFact

cnnGraf.Execute "INSERT INTO ChartDeuda(Mes,Deuda,Facturado,Pagos) VALUES ('" & pIni & "','" _
& curIni & "','" & curFact & "','" & curPago & "')"
pIni = DateAdd("d", 1, pIni)

Do
    
    pPeriodo = DateAdd("m", 1, pPeriodo)
    
    strSQL = "SELECT * FROM Factura WHERE Periodo =#" & Format(pPeriodo, "m/d/yy") & "#"
    rstGraf.Open strSQL, cnnGraf, adOpenKeyset, adLockOptimistic, adCmdText
    If Not rstGraf.EOF And Not rstGraf.BOF Then
        If Not IsNull(rstGraf("FechaFactura")) Then fFin = rstGraf("FechaFactura")
    End If
    rstGraf.Close
    '
    'calculo lo facturado
    curFact = 0
    strSQL = "SELECT Sum(Facturado) AS Factu FROM Factura WHERE FechaFactura IN (SELECT FechaFa" _
    & "ctura FROM Factura WHERE Periodo =#" & Format(pPeriodo, "m/d/yy") & "#)"
    
    rstGraf.Open strSQL, cnnGraf, adOpenKeyset, adLockOptimistic, adCmdText
    
    If Not rstGraf.EOF And Not rstGraf.BOF Then
        If Not IsNull(rstGraf("Factu")) Then curFact = rstGraf("Factu")
    End If
    rstGraf.Close
    '
    'calculo lo pagado
    curPago = 0
    strSQL = "SELECT Sum(P.Monto) AS Pagos FROM MovimientoCaja as C INNER JOIN Periodos as " _
    & "P ON C.IDRecibo = P.IDRecibo WHERE (((P.CodGasto)='900030' Or (P.CodGasto)='900001') " _
    & "AND ((C.FechaMovimientoCaja) Between #" & Format(pIni, "mm/dd/yy") & "# And #" & _
    Format(fFin, "mm/dd/yy") & "#) AND ((C.InmuebleMovimientoCaja)='" & Inm & "'));"

    rstGraf.Open strSQL, cnnConexion, adOpenKeyset, adLockOptimistic, adCmdText
    If Not rstGraf.BOF And Not rstGraf.EOF Then
        If Not IsNull(rstGraf("Pagos")) Then curPago = rstGraf("Pagos")
    End If
    rstGraf.Close
    '
    curIni = curIni + curFact - curPago
    
    cnnGraf.Execute "INSERT INTO ChartDeuda(Mes,Deuda,Facturado,Pagos) VALUES ('" & fFin & "','" _
    & curIni & "','" & curFact & "','" & curPago & "')"
    
    pIni = DateAdd("D", 1, fFin)
    
Loop Until fFin = pFin
'
End Sub

Private Function novalida() As Boolean
Dim rstFecFact As New ADODB.Recordset

If Inm = "" Then
    novalida = MsgBox("Falta el Código del Inmueble", vbExclamation, App.ProductName)
    Exit Function
End If
If Not IsDate("01/" & cmb(1) & "/" & cmb(3)) Then
    novalida = MsgBox("Introdujo en valor no válido en el campo 'Inicio Período'", vbExclamation, _
    App.ProductName)
    Exit Function
Else
    pIni = "01/" & cmb(1) & "/" & cmb(3)
    pPeriodo = pIni
    rstFecFact.Open "Select FechaFactura FROM Factura WHERE Periodo =#" & Format(pIni, "m/d/yy") & "#", cnnOLEDB & _
    gcPath & "\" & Inm & "\inm.mdb", adOpenKeyset, adLockOptimistic, adCmdText
    If Not rstFecFact.EOF And Not rstFecFact.BOF Then
        pIni = rstFecFact("FechaFactura")
    End If
    rstFecFact.Close
    
End If
If Not IsDate("01/" & cmb(2) & "/" & cmb(4)) Then
    novalida = MsgBox("Introdujo en valor no válido en el campo 'Final Período'", vbExclamation, _
    App.ProductName)
    Exit Function
Else
    pFin = "01/" & cmb(2) & "/" & cmb(4)
    rstFecFact.Open "SELECT FechaFactura FROM Factura WHERE Periodo =#" & Format(pFin, "m/d/yy") & "#", cnnOLEDB & gcPath & "\" & Inm & "\inm.mdb", adOpenKeyset, adLockOptimistic, adCmdText
    If Not rstFecFact.EOF And Not rstFecFact.BOF Then
        pFin = rstFecFact("FechaFactura")
    End If
    rstFecFact.Close
    
End If
If pFin < pIni Then
    novalida = MsgBox("El Final el período no puede ser un valor menor que el inicio", _
    vbExclamation, App.ProductName)
End If
'
Set rstFecFact = Nothing

End Function
