VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "stRFactura"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "Revertir Facturacion de un mes en específico"
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'Dll STREVFACT001*********************************************************************************
'   Reverso de facturacion
'
'variables locales para almacenar los valores de las propiedades
Private mvarCodigo_Inmueble As String 'copia local
Private mvarPeriodo As Date 'copia local
Dim conexionP As New ADODB.Connection
Dim conexionS As New ADODB.Connection
Dim Proveedor As String
'Const Proveedor$ = "Provider=Microsoft.Jet.OLEDB.3.51;Data Source="
'variables locales para almacenar los valores de las propiedades
Private mvarsConexionP As String 'copia local
Private mvarsConexionS As String 'copia local
'variables locales para almacenar los valores de las propiedades
Private mvarRuta As String 'copia local
Private mvarReporte As String 'copia local
Private INI(1 To 7) As String

Public Property Let Reporte(ByVal vData As String)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.Reporte = 5
    mvarReporte = vData
End Property

Public Property Get Reporte() As String
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.Reporte
    Reporte = mvarReporte
End Property



Public Property Let Ruta(ByVal vData As String)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.Ruta = 5
    mvarRuta = vData
End Property


Public Property Get Ruta() As String
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.Ruta
    Ruta = mvarRuta
End Property




Public Property Let sConexionS(ByVal vData As String)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.sConexionS = 5
    mvarsConexionS = vData
End Property


Public Property Get sConexionS() As String
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.sConexionS
    sConexionS = mvarsConexionS
End Property



Public Property Let sConexionP(ByVal vData As String)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.sConexionP = 5
    mvarsConexionP = vData
End Property


Public Property Get sConexionP() As String
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.sConexionP
    sConexionP = mvarsConexionP
End Property




Public Property Let Periodo(ByVal vData As Date)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.Periodo = 5
    mvarPeriodo = vData
End Property


Public Property Get Periodo() As Date
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.Periodo
    Periodo = mvarPeriodo
End Property


Public Property Let Codigo_Inmueble(ByVal vData As String)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.Codigo_Inmueble = 5
mvarCodigo_Inmueble = vData
End Property


Public Property Get Codigo_Inmueble() As String
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.Codigo_Inmueble
Codigo_Inmueble = mvarCodigo_Inmueble
End Property

'-------------------------------------------------------------------------------------------------
'   Rutina:     Revertir_Facturacion
'
'   Coloca todos los pagos seleccionados en el depósito de abonos a futuro
'-------------------------------------------------------------------------------------------------
Public Function Revertir_Facturacion() As Boolean
Attribute Revertir_Facturacion.VB_UserMemId = 0
'Variables locales
conexionP.CursorLocation = adUseClient
conexionP.Open Proveedor & sConexionP
conexionS.CursorLocation = adUseClient
conexionS.Open Proveedor & sConexionS
Dim strSql As String, strMes As String
Dim strPC As String, strAC As String, strNF As String
Dim strAF As String, strCartas As String, strTelegrama As String
Dim Fecha_Facturacion As String
Dim rstPE As New ADODB.Recordset
Dim i As Integer
'
'confirma que el período haya sido facturado
rstPE.Open "SELECT * FROM Factura WHERE Periodo=#" & Periodo & "# AND Fact Not Like 'CH%';", _
conexionS, adOpenStatic, adLockReadOnly
If rstPE.RecordCount <= 0 Then
    Revertir_Facturacion = MsgBox("Período No está facturado, vea el estado de cuenta", _
    vbOKOnly + vbInformation)
    Exit Function
End If
'Elimina todos los reportes de la facturacion de la carpeta reportes
'del inmueble
For i = 1 To 6
    If Not Dir(Ruta & INI(i) & Reporte) = "" And i <> 6 Then Kill Ruta & INI(i) & Reporte
Next i
If Dir(Ruta & "DC" & Format(Periodo, "ddyyyy") & ".rpt") <> "" Then Kill Ruta & "DC" & Format(Periodo, "ddyyyy") & ".rpt"
If Dir(Ruta & "DG" & Format(Periodo, "ddyyyy") & ".rpt") <> "" Then Kill Ruta & "DG" & Format(Periodo, "ddyyyy") & ".rpt"
rstPE.Close
rstPE.Open "SELECT * FROM Inmueble WHERE CodInm='" & Codigo_Inmueble & "';", conexionP, _
adOpenStatic, adLockReadOnly
'//Asigna valores a las variables locales
'//Código de gastos
strPC = rstPE!CodPagoCondominio 'Pago de condominio
strAC = rstPE!CodAbonoCta   'Abono a cuenta
strAF = rstPE!CodAbonoFut   'Abono a futuro
strCarta = rstPE!CodCarta   'Carta de cobro
strTelegrama = rstPE!CodTelegrama   'Telegrama
rstPE.Close
strMes = Format(Periodo, "dd-yy")
'//Selecciona los pagos efectuados por los clientes al período reversado
strSql = "SELECT Periodos.*, MovimientoCaja.AptoMovimiento" _
& "Caja  FROM MovimientoCaja INNER JOIN Periodos ON MovimientoCaja.IDRecibo =" _
& "Periodos.IDRecibo WHERE (((Periodos.IDRecibo) Like '" & Right(Codigo_Inmueble, 2) & "%') AND (" _
& "(Periodos.Periodo)='" & strMes & "') AND ((Periodos.CodGasto)='" & strPC & "'OR (Periodos.Co" _
& "dGasto)='" & strAC & "'));"
'
rstPE.Open strSql, conexionP, adOpenDynamic, adLockOptimistic
With rstPE
'
    If .RecordCount > 0 Then    'Si existen pagos los coloca en abono a futuro
        .MoveFirst
        Do
            conexionP.Execute "INSERT INTO TDFAbonos(IDRecibo,Monto) VALUES('" & !IDRecibo _
            & "','" & !Monto & "')"
            .Update "CodGasto", strAF
            .Update "Descripcion", "ABONO A PROX. FACTURACION"
            .MoveNext
        Loop Until .EOF
        '
    End If
    '
    .Close
End With

'-----------------------------------------
'   Actualiza estados de cheques devueltos, gastos de cartas y telegramas
'   además de las cuentas de fondos
'-----------------------------------------
strSql = "SELECT freg FROM Factura WHERE Periodo=#" & Periodo & "#;"
rstPE.Open strSql, conexionS, adOpenKeyset, adLockOptimistic
If Not rstPE.EOF Or Not rstPE.BOF Then
    Fecha_Facturacion = Format(rstPE.Fields(0), "mm/dd/yyyy")
    'Actualiza los cobros cheques devueltos
    conexionS.Execute "UPDATE ChequeDevuelto SET Recuperado=False WHERE Freg=#" _
    & Fecha_Facturacion & "#;"
End If
rstPE.Close
strSql = "* FROM Cpp WHERE Fact='F" & Codigo_Inmueble & "' AND Detalle LIKE 'HONOR" _
& "ARIO%' AND Freg=#" & Fecha_Facturacion & "#;"
rstPE.Open "SELECT " & strSql, conexionP, adOpenKeyset, adLockOptimistic
If Not rstPE.EOF Or Not rstPE.BOF Then
    If rstPE!Estatus = "PAGADO" Then
        Revertir_Facturacion = MsgBox("La factura por honorarios y gestiones le emitieron cheque," _
        & vbCrLf & "verifique por favor....", vbInformation, "Reverso Facturación")
    End If
    'Elimina la cuenta por pagar por concepto de honorarios y gestiones
    conexionS.Execute "DELETE * FROM AsignaGasto WHERE Ndoc='" & rstPE!NDoc & "';"
    conexionS.Execute "DELETE * FROM Cargado WHERE Ndoc='" & rstPE!NDoc & "';"
    conexionP.Execute "DELETE " & strSql
End If
rstPE.Close
'Actualiza los cobros por cartas enviadas
strSql = "Update Propietarios SET Carta=True WHERE Codigo IN (SELECT CodApto FROM GastoNoComun" _
& " WHERE GastoNoComun.Periodo=#" & Periodo & "# AND GastoNoComun.CodGasto='" & strCarta & "');"
conexionS.Execute strSql
'Actualiza los cobros por telegramas enviados
strSql = "Update Propietarios SET Carta=True WHERE Codigo IN (SELECT CodApto FROM GastoNoComun" _
& " WHERE GastoNoComun.Periodo=#" & Periodo & "# AND GastoNoComun.CodGasto='" & strTelegrama & "');"
conexionS.Execute strSql
'Elimina el movimiento de la cuenta de fondo (MovFondo - Pago_Inf)
conexionS.Execute "DELETE FROM Pago_Inf WHERE Periodo=#" & Periodo & "#;"
conexionS.Execute "DELETE FROM MovFondo WHERE Periodo=#" & Periodo & "# AND Tipo='FA';"
'
rstPE.Open "SELECT CodGasto,cargado,Descripcion,Monto FROM AsignaGasto WHERE CodGasto IN (S" _
& "ELECT CodGasto FROM Tgastos WHERE Fondo=True;) AND Cargado=#" & Periodo & "# UNION SE" _
& "LECT CodGasto,Periodo,concepto, Sum(Monto)  FROM GastoNoComun GROUP BY CodGasto, Concept" _
& "o, Periodo HAVING CodGasto In (Select CodGasto FROM Tgastos WHERE Fondo=True) AND Period" _
& "o=#" & Periodo & "#;", conexionS, adOpenKeyset, adLockOptimistic
    
If Not rstPE.EOF Or Not rstPE.BOF Then
    With rstPE
    .MoveFirst
    Do Until .EOF
        'Actualiza el saldo actual, Saldo del mes Facturado
        conexionS.Execute "UPDATE Tgastos SET SaldoActual=SaldoActual - '" & CCur(!Monto) & "',Sa" _
        & "ldo" & Format(Periodo, "d") & "= SaldoActual - '" & CCur(!Monto) & "' WHE" _
        & "RE Codgasto='" & !codGasto & "';"
        'Inserta el movimiento en TDF MovFondo
        'conexionS.Execute "INSERT INTO movFondo (CodGasto,Fecha,Tipo,Periodo,Concepto,Debe,Haber)" _
        & " VALUES ('" & !CodGasto & "',Date(),'RF',#" & Periodo & "#,'" & !Descripcion _
        & "','" & CCur(!Monto) & "',0);"
        .MoveNext
    Loop
    End With
End If
rstPE.Close
'---------------------------------------------------------------------------------------------
'   Elimina la factuacion señalada y actualiza las deudas
'---------------------------------------------------------------------------------------------
strNF = Format(Periodo, "ddyy") & Right(gcCodInm, 3)
'actualiza la deuda del inmueble
Dim curDeuda As Currency
Dim CurFac As Currency
Dim CurFR As Currency
'---------------------------------------------------------------------------------------------
'Actualiza la deuda de los propietarios
conexionS.Execute "UPDATE Propietarios INNER JOIN Factura ON Propietarios.Codigo = Factura.codpro" _
& "p SET Propietarios.Deuda = Propietarios.Deuda - Factura.Facturado, Propietarios.Reci" _
& "bos=Propietarios.Recibos - 1 WHERE Factura.FACT LIKE '" & strNF & "%'"
'Elimina facturacion del mes
conexionS.Execute "DELETE FROM Factura WHERE FACT LIKE '" & strNF & "%'"
'
rstPE.Open "SELECT SUM(Saldo) FROM Factura;", conexionS, adOpenKeyset, adLockOptimistic
curDeuda = IIf(rstPE.EOF Or rstPE.BOF, 0, rstPE.Fields(0))
conexionP.Execute "UPDATE Inmueble SET Deuda='" & curDeuda & "' WHERE CodInm='" _
& Codigo_Inmueble & "'"
rstPE.Close
'
rstPE.Open "SELECT T1.InmuebleMovimientoCaja,Sum(T2.Monto) FROM MovimientoCaja " _
& "as T1 INNER JOIN TDFAbonos AS T2 ON T1.IDRecibo=T2.IDRecibo WHERE T1.InmuebleMovimientoC" _
& "aja='" & Codigo_Inmueble & "' GROUP BY T1.InmuebleMovimientoCaja", conexionP, adOpenKeyset, _
adLockOptimistic
curDeuda = IIf(rstPE.BOF Or rstPE.EOF, 0, rstPE.Fields(0))
conexionP.Execute "UPDATE Inmueble SET Deuda=Deuda - '" & curDeuda & "' WHERE CodInm='" _
& Codigo_Inmueble & "'"
rstPE.Close
Set rstPE = Nothing
conexionP.Close
conexionS.Close
Set conexionP = Nothing
Set conexionS = Nothing
'
End Function

''-------------------------------------------------------------------------------------------------
''   Rutina:     Buscar_gastosE
''
''
''-------------------------------------------------------------------------------------------------
'Private Sub Buscar_GastosE(strUbicaBD As String, strPeriodo As String)
''Variables locales
'strSql = "SELECT freg FROM Factura WHERE Periodo=#" & strPeriodo & "#;"
'rstGE.Open strSql, cnnBD, adOpenKeyset, adLockOptimistic
''
'Fecha_Facturacion = Format(rstGE.Fields(0), "mm/dd/yyy")
'rstGE.Close
''Actualiza los cobros cheques devueltos
'cnnBD.Execute "UPDATE ChequeDevuelto SET Recuperado=False WHERE Freg=#" & Fecha_Facturacion & "#;"
''Actualiza los cobros por cartas enviadas
'strSql = "Update Propietarios SET Carta=True WHERE Codigo IN (SELECT CodApto FROM GastoNoComun" _
'& " WHERE GastoNoComun.Periodo=#" & strPeriodo & "# AND GastoNoComun.CodGasto='" & strCarta & "';"
'cnnBD.Execute strSql
''Actualiza los cobros por telegramas enviados
'strSql = "Update Propietarios SET Carta=True WHERE Codigo IN (SELECT CodApto FROM GastoNoComun" _
'& " WHERE GastoNoComun.Periodo=#" & strPeriodo & "# AND GastoNoComun.CodGasto='" & strTelegrama & "';"
'cnnBD.Execute strSql
''Actualiza los movimientos de fondos
''
'rstGE.Open "SELECT CodGasto,cargado,Descripcion,Monto FROM AsignaGasto WHERE CodGasto IN (S" _
'& "ELECT CodGasto FROM Tgastos WHERE Fondo=True;) AND Cargado=#" & strPeriodo & "# UNION SE" _
'& "LECT CodGasto,Periodo,concepto, Sum(Monto)  FROM GastoNoComun GROUP BY CodGasto, Concept" _
'& "o, Periodo HAVING CodGasto In (Select CodGasto FROM Tgastos WHERE Fondo=True) AND Period" _
'& "o=#" & strPeriodo & "#;", cnnBD, adOpenKeyset, adLockOptimistic
'
'If Not rstGE.EOF Or Not rstGE.BOF Then
'    With rstGE
'    .MoveFirst
'    Do Until .EOF
'        'Actualiza el saldo actual, Saldo del mes Facturado
'        cnn.Execute "UPDATE Tgastos SET SaldoActual=SaldoActual - '" & CCur(!Monto) & "',Sa" _
'        & "ldo" & Format(strPeriodo, "d") & "= SaldoActual - '" & CCur(!Monto) & "' WHE" _
'        & "RE Codgasto='" & !CodGasto & "';"
'        'Inserta el movimiento en TDF MovFondo
'        cnn.Execute "INSERT INTO movFondo (CodGasto,Fecha,Tipo,Periodo,Concepto,Debe,Haber)" _
'        & " VALUES ('" & !CodGasto & "',Date(),'RF',#" & strPeriodo & "#,'" & !Descripcion _
'        & "','" & CCur(!Monto) & "',0);"
'        .MoveNext
'    Loop
'    rstGE.Close
'    Set rstGE = Nothing
'    '
'    End With
'End Sub

'Private Sub Eliminar_Facturacion()
'
'strNF = Format(strPeriodo, "ddyy") & Right(gcCodInm, 3)
''actualiza la deuda del inmueble
'Dim curDeuda, CurFac, CurFR As Currency
'Dim rstDeuda As ADODB.Recordset
'Set rstDeuda = New ADODB.Recordset
''---------------------------------------------------------------------------------------------
''Actualiza la deuda de los propietarios
'cnn.Execute "UPDATE Propietarios INNER JOIN Factura ON Propietarios.Codigo = Factura.codpro" _
'& "p SET Propietarios.Deuda = Propietarios.Deuda - Factura.Facturado, Propietarios.Reci" _
'& "bos=Propietarios.Recibos - 1 WHERE Factura.FACT) LIKE '" & strNF & "*'"
''Elimina facturacion del mes
'cnn.Execute "DELETE FROM Fractura WHERE FACT LIKE '" & strNF & "*'"
''
'rstDeuda.Open "SELECT SUM(Saldo) FROM Factura;", cnn, adOpenKeyset, adLockOptimistic
'curDeuda = IIf(rstDeuda.EOF Or rstDeuda.BOF, 0, rstDeuda.Fields(0))
'cnnConexion.Execute "UPDATE Inmueble SET Deuda='" & curDeuda & "' WHERE CodInm='" _
'& gcCodInm & "'"
'rstDeuda.Close
''
'rstDeuda.Open "SELECT T1.InmuebleMovimientoCaja,Sum(T2.Monto) FROM MovimientoCaja " _
'& "as T1 INNER JOIN TDFAbonos AS T2 ON T1.IDRecibo=T2.IDRecibo WHERE T1.InmuebleMovimientoC" _
'& "aja='" & gcCodInm & "' GROUP BY T1.InmuebleMovimientoCaja", cnnConexion, adOpenKeyset, _
'adLockOptimistic
'curDeuda = IIf(rstDeuda.BOF Or rstDeuda.EOF, 0, rstDeuda.Fields(0))
'cnnConexion.Execute "UPDATE Inmueble SET Deuda=Deuda - '" & curDeuda & "' WHERE CodInm='" _
'& gcCodInm & "'"
'rstDeuda.Close
'Set rstDeuda = Nothing
''
'End Sub

Private Sub Class_Initialize()
INI(1) = "CF"
INI(2) = "AC"
INI(3) = "AN"
INI(4) = "NC"
INI(5) = "PR"
INI(6) = "F"
INI(7) = "R"
Proveedor = GetSetting(App.EXEName, "Entorno", "Proveedor", "")
End Sub

Private Sub Class_Terminate()
On Error Resume Next
'conexionP.Close
'conexionS.Close
Set conexionP = Nothing
Set conexionS = Nothing
End Sub
