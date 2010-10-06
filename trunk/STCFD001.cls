VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "STCFD001"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Const Deuda% = 1
Const Fondo% = 2
Const Moneda$ = "#,##0.00 ;(#,##0.00) "
Const Archivo$ = "CFD.log"
Dim ObjCnn As New ADODB.Connection
'variables locales para almacenar los valores de las propiedades
Private mvarCadena_Conexion As String 'copia local
'variables locales para almacenar los valores de las propiedades
Private mvarGrid As Object 'copia local
Public Directorio As String

Public Property Set Grid(ByVal vData As Object)
'se usa al asignar un objeto a la propiedad, en la parte izquierda de una instrucción Set.
'Syntax: Set x.Grid = Form1
    Set mvarGrid = vData
End Property


Public Property Get Grid() As Object
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.Grid
    Set Grid = mvarGrid
End Property

Public Property Let Cadena_Conexion(ByVal vData As String)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.Cadena_Conexion = 5
    mvarCadena_Conexion = vData
End Property


Public Property Get Cadena_Conexion() As String
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.Cadena_Conexion
    Cadena_Conexion = mvarCadena_Conexion
End Property


Public Function Config_Grid(Opcion As Integer, lblI As Object, LblA As Object, _
lblD As Object) As Long
'
Dim i As Integer
Dim Respuesta As Integer

With Grid
    If Opcion = Deuda Then
        .FormatString = "Cod.Edif.| Deuda Anterior | Caja | Cheq. Dev. | Facturacion | Totales " _
        & "| Deuda Actual | Diferencia"
        .ColWidth(0) = 800
    Else
        .FormatString = "Cod.Edif.| Fondo Anterior | |Egresos | Ingresos | Totales | Fondo Actual | Diferencia"
        .ColWidth(0) = 1000
    End If
    
    .Row = 0
    .Col = 0
    .CellAlignment = flexAlignCenterCenter
    For i = 1 To .Cols - 1
        .ColWidth(i) = 1300
        .Col = i
        .CellAlignment = flexAlignCenterCenter
    Next
    .ColAlignment(0) = flexAlignCenterCenter
    If Not Opcion = Deuda Then .ColWidth(2) = 0
    Call Mostrar_Info(Grid, lblI, LblA, lblD, Opcion)
    Call RtnColorFondo
    'Verifica si existe un archivo temporal
    If Dir(Directorio & Opcion & Archivo) <> "" Then
        Respuesta = MsgBox("Existe información almacenada en el temporal. ¿Desea recuperarla?", _
        vbYesNo, "ST SAC")
        If Respuesta = vbYes Then Call Recuperar_Inf(Opcion)
    End If
End With
End Function

Private Sub RtnColorFondo()
'
With Grid
    .Col = 0
    .Row = 1
    .ColSel = 1
    .RowSel = .Rows - 1
    .FillStyle = flexFillRepeat
    .CellBackColor = &H80000018
    .Col = 5
    .Row = 1
    .ColSel = 7
    .RowSel = .Rows - 1
    .FillStyle = flexFillRepeat
    .CellBackColor = &H80000018
    .FillStyle = flexFillSingle
    .Col = 0
    .Row = 1
End With
'
End Sub

Private Sub Mostrar_Info(Grid As Object, lblI As Label, LblA As Label, lblD As Label, _
Optional Opcion As Integer)

Dim K As Integer    'Variables locales
Dim objRst As New ADODB.Recordset
Dim strSql As String

ObjCnn.Open Cadena_Conexion
If Opcion = Deuda Then  'Esta cuadrando la deuda

    'strSql = "SELECT CodInm,DeudaIni,DeudaAct FROM Inmueble WHERE Caja not In (SELECT Caja FROM In" _
    & "mueble WHERE CodInm='9999') AND Not Inactivo UNION SELECT '9999',Sum(DeudaIni),Sum(Deuda" _
    & "Act) FROM Inmueble WHERE Caja IN (SELECT Caja FROM Inmueble WHERE CodInm='9999') AND Not In" _
    & "activo  AND CodInm<>'9999';"
    strSql = "SELECT CodInm,DeudaIni,DeudaAct FROM Inmueble WHERE  Not Inactivo AND CodInm <>'" _
    & sysCodInm & "' ORDER BY CodInm"
Else    'se esta cuadrando el fondo de reserva
    'strSql = "SELECT CodInm,FondoIni,Fondo FROM Inmueble WHERE Caja not In (SELECT Caja FROM In" _
    & "mueble WHERE CodInm='9999') AND Not Inactivo UNION SELECT '9999',Sum(FondoIni),Sum(Fondo" _
    & ") FROM Inmueble WHERE Caja IN (SELECT Caja FROM Inmueble WHERE CodInm='9999') AND Not In" _
    & "activo  AND CodInm<>'9999';"
    strSql = "SELECT CodInm,FondoIni,Fondo FROM Inmueble WHERE Not Inactivo AND CodInm <>'" _
    & sysCodInm & "' ORDER BY CodInm"
End If
'
objRst.Open strSql, ObjCnn, adOpenStatic, adLockReadOnly
With objRst
    If .EOF Or .BOF Then Exit Sub
    Grid.Rows = .RecordCount + 1
    .MoveFirst
    If Opcion = Deuda Then
        Do
            K = K + 1
            Grid.TextMatrix(K, 0) = !CodInm
            Grid.TextMatrix(K, 1) = Format(!DeudaIni, Moneda)
            Grid.TextMatrix(K, 5) = Format(!DeudaIni, Moneda)
            Grid.TextMatrix(K, 6) = Format(!DeudaAct, Moneda)
            Grid.TextMatrix(K, 7) = Format(!DeudaIni - !DeudaAct, Moneda)
            lblI = Format(lblI + !DeudaIni, Moneda)
            LblA = Format(LblA + !DeudaAct, Moneda)
            .MoveNext
        Loop Until .EOF
        lblD = Format(CCur(lblI) - CCur(LblA), Moneda)
    Else '"Cod.Edif.|Fondo Anterior|Egresos|Ingresos|Totales|Fondo Actual|Diferencia"
        Do
            K = K + 1
            Grid.TextMatrix(K, 0) = !CodInm
            Grid.TextMatrix(K, 1) = Format(!FondoIni, Moneda)
            Grid.TextMatrix(K, 5) = Format(!FondoIni, Moneda)
            Grid.TextMatrix(K, 6) = Format(!Fondo, Moneda)
            Grid.TextMatrix(K, 7) = Format(!FondoIni - !Fondo, Moneda)
            lblI = Format(lblI + !FondoIni, Moneda)
            LblA = Format(LblA + !Fondo, Moneda)
            .MoveNext
        Loop Until .EOF
        lblD = Format(CCur(lblI) - CCur(LblA), Moneda)
    End If
    objRst.Close
End With

Set objRst = Nothing
ObjCnn.Close
End Sub

Public Sub Sumatoria(Columna As Integer, Fila As Integer, Etiqueta As Object)
'
Dim i As Integer    'variables locales
Dim Total As Currency
'
With Grid
    'Totaliza la linea
    For i = 1 To 4
        Total = Total + IIf(.TextMatrix(Fila, i) = "", 0, .TextMatrix(Fila, i))
    Next i
    .TextMatrix(Fila, 5) = Format(Total, Moneda)
    .TextMatrix(Fila, 7) = Format(Total - CCur(.TextMatrix(Fila, 6)), Moneda)
    Total = 0
    For i = 1 To .Rows - 1
        If Not .TextMatrix(i, Columna) = "" Then
            Total = Total + .TextMatrix(i, Columna)
        End If
    Next
    Etiqueta = Format(Total, Moneda)
End With
'
End Sub

'----------------------------------------------------------------------------------------
'   Rutina: Guardar_inf
'
'   Guarda la información introducida por el usuario hasta el el momento
'   en un archivo temporal, para ser recuperada cuando se necesite
'----------------------------------------------------------------------------------------
Public Sub Guardar_Inf(Opcion As Integer)
'Variables locales
Dim numFichero As Integer
Dim strArchivo As String
Dim booG As Boolean
'
On Error Resume Next
With Grid
    numFichero = FreeFile
    strArchivo = Directorio & Opcion & Archivo
    Open strArchivo For Output As numFichero
        J = 1
        Do
            If Opcion = Deuda Then  'cuadre de deuda
                If .TextMatrix(J, 2) <> "" Or .TextMatrix(J, 3) <> "" Or .TextMatrix(J, 4) <> "" Then
                    Write #numFichero, J, .TextMatrix(J, 2), .TextMatrix(J, 3), .TextMatrix(J, 4)
                    booG = True
                End If
            Else    'cuadre de fondo
                If .TextMatrix(J, 4) <> "" Or .TextMatrix(J, 3) <> "" Then
                    Write #numFichero, J, .TextMatrix(J, 3), .TextMatrix(J, 4)
                    booG = True
                End If
            End If
            J = J + 1
        Loop Until J > .Rows - 1
    Close numFichero
End With
If Not booG And Dir(strArchivo) <> "" Then Kill (strArchivo)
If Err.Number <> 0 Then
    MsgBox "Error al guardar la información. No se pudo generar el archivo" & vbCrLf & _
    Err.Description, vbCritical
End If
'
End Sub

Public Sub Recuperar_Inf(Opcion As Integer)
'Variables locales
Dim numFichero As Integer
Dim strArchivo As String
'-----------
On Error Resume Next
numFichero = FreeFile
strArchivo = Directorio & Opcion & Archivo
'
Open strArchivo For Input As numFichero
With Grid
    '
    If Opcion = Deuda Then    'Cuadre de Deuda
        Do
            Input #numFichero, intL, PCaja, CheqDev, Facturacion
            .TextMatrix(intL, 2) = PCaja
            .TextMatrix(intL, 3) = CheqDev
            .TextMatrix(intL, 4) = Facturacion
        Loop Until EOF(numFichero)
    Else    'Cuade de Fondo
        Do
            Input #numFichero, intL, Debitos, Creditos
                .TextMatrix(intL, 3) = Debitos
                .TextMatrix(intL, 4) = Creditos
        Loop Until EOF(numFichero)
    End If
    '
End With
Close numFichero
'
End Sub

'----------------------------------------------------------------------------------------
'   Rutina: Print_cuadre
'
'   Entrada:    Opcion Interger, cual es el reporte que se quiere
'
'   Agrega toda la información que se encuentra en pantalla a un archivo
'----------------------------------------------------------------------------------------
Public Sub Print_Cuadre(Opcion As Integer)
'Variables locales
Dim strSql As String
ObjCnn.Open Cadena_Conexion
With Grid
    If Opcion = Deuda Then  'Cuadre de deuda
        ObjCnn.Execute "DELETE * FROM STCFDDeuda"
        For i = 1 To .Rows - 1
            strSql = "INSERT INTO STCFDDeuda (CodInm,DAnterior,Caja,CheqDev,Facturacion,Totales" _
            & ",DActual,Diferencia) VALUES('" & .TextMatrix(i, 0) & "','" & .TextMatrix(i, 1) & _
            "','" & IIf(.TextMatrix(i, 2) = "", 0, .TextMatrix(i, 2)) & "','" & _
            IIf(.TextMatrix(i, 3) = "", 0, .TextMatrix(i, 3)) & "','" & _
            IIf(.TextMatrix(i, 4) = "", 0, .TextMatrix(i, 4)) & "','" & .TextMatrix(i, 5) & _
            "','" & .TextMatrix(i, 6) & "','" & .TextMatrix(i, 7) & "');"
            ObjCnn.Execute strSql
        Next
    Else    'Cuadre de fondo
        ObjCnn.Execute "DELETE * FROM STCFDFondo"
        For i = 1 To .Rows - 1
            strSql = "INSERT INTO  STCFDFondo (CodInm,FAnterior,Debitos,Creditos,Totales,FActua" _
            & "l,Diferencia) VALUES('" & .TextMatrix(i, 0) & "','" & .TextMatrix(i, 1) & "','" & _
            IIf(.TextMatrix(i, 3) = "", 0, .TextMatrix(i, 3)) & "','" & _
            IIf(.TextMatrix(i, 4) = "", 0, .TextMatrix(i, 4)) & "','" & .TextMatrix(i, 5) & "','" _
            & .TextMatrix(i, 6) & "','" & .TextMatrix(i, 7) & "');"
            ObjCnn.Execute strSql
        Next
    End If
    '
End With
ObjCnn.Close
'
End Sub


Private Sub Class_Terminate()
Set ObjCnn = Nothing
End Sub


