VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCaja"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'local variable(s) to hold property value(s)
Private mvarTaquilla As Integer 'local copy
Private mvarFormadePago As String 'local copy
Private mvarNumeroDocumento As String 'local copy
Private mvarFechaDocumento As String 'local copy
Private mvarMonto As Double 'local copy
Private mvarFacturas As Collection 'local copy

Public Function guardar_pago() As Boolean
Attribute guardar_pago.VB_UserMemId = 0
Dim sCodInm As String
Dim sFact As String
Dim nPago As Double

For I = 0 To mvarFacturas.Count
    sFact = Mid(mvarFacturas.Item, 4, 3)
    If sCodInm <> sFact Then
        sCodInm = "2" & sFact
    End If
    nPago = pagar_factura(sFact, sCodInm)
Next

End Function

Private Function nPago(Factura, Inmueble)

End Function

Public Property Set Facturas(ByVal vData As Collection)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.Facturas = Form1
    Set mvarFacturas = vData
End Property


Public Property Get Facturas() As Collection
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Facturas
    Set Facturas = mvarFacturas
End Property



Public Property Let Monto(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Monto = 5
    mvarMonto = vData
End Property


Public Property Get Monto() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Monto
    Monto = mvarMonto
End Property



Public Property Let FechaDocumento(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.FechaDocumento = 5
    mvarFechaDocumento = vData
End Property


Public Property Get FechaDocumento() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.FechaDocumento
    FechaDocumento = mvarFechaDocumento
End Property



Public Property Let NumeroDocumento(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.NumeroDocumento = 5
    mvarNumeroDocumento = vData
End Property


Public Property Get NumeroDocumento() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.NumeroDocumento
    NumeroDocumento = mvarNumeroDocumento
End Property


Public Property Let FormadePago(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.FormadePago = 5
    mvarFormadePago = vData
End Property


Public Property Get FormadePago() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.FormadePago
    FormadePago = mvarFormadePago
End Property



Public Property Let Taquilla(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Taquilla = 5
    mvarTaquilla = vData
End Property


Public Property Get Taquilla() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Taquilla
    Taquilla = mvarTaquilla
End Property



