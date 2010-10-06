VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "STEndoso"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'variables locales para almacenar los valores de las propiedades
Private mvarCuenta As String 'copia local
Private mvarBanco As String 'copia local
Private mvarTitular As String 'copia local

Public Property Let Titular(ByVal vData As String)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.Titular = 5
    mvarTitular = vData
End Property


Public Property Get Titular() As String
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.Titular
    Titular = mvarTitular
End Property



Public Property Let Banco(ByVal vData As String)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.Banco = 5
    mvarBanco = vData
End Property


Public Property Get Banco() As String
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.Banco
    Banco = mvarBanco
End Property


Public Property Let Cuenta(ByVal vData As String)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.Cuenta = 5
    mvarCuenta = vData
End Property


Public Property Get Cuenta() As String
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.Cuenta
    Cuenta = mvarCuenta
End Property

Function Endosar() As Boolean
'Función que endosa en cheque devuelve 'Verdadero' si se origina algún error
Dim strCadena As String
On Error GoTo salir
'Printer.PaperSize = 257
Printer.Orientation = 1
Printer.Font.Name = "Courier New"
Printer.Font.Size = 10
Printer.FontBold = True
strCadena = "unicamente para ser depositado en la cuenta # " & mvarCuenta & " en el banco " _
& mvarBanco & " a nombre de " & mvarTitular
Printer.Print
strCadena = Formatear(strCadena)
Printer.EndDoc
salir:
If Err.Number <> 0 Then
    Endosar = MsgBox("El papel solicitado no está configurado en su impresora," & vbCrLf _
    & "Retire el cheque de la bandeja y consulte al administrador del sistema.", vbInformation)
End If
End Function

Private Function Formatear(ByVal Cadena As String) As String
'Da formato al texto que se va imprimir, al centro del documento
'en varias líneas--------------------
Dim strTemp, strLinea, strPalabra As String 'Variables locales
Dim lngPos As Long
'
Do Until Cadena = ""
    lngPos = InStr(Cadena, " ")
    If lngPos = 0 Then
        strPalabra = Cadena
        Cadena = ""
    Else
        strPalabra = Left(Cadena, lngPos)
        Cadena = Right(Cadena, Len(Cadena) - lngPos)
    End If
    strTemp = strLinea & strPalabra
    If Len(strTemp) > 34 Then
        'Printer.CurrentX
        Formatear = strLinea '& vbCrLf
        Printer.CurrentX = (Printer.ScaleWidth / 2) - (Printer.TextWidth(Formatear) / 2)
        Printer.Print UCase(Formatear)
        strTemp = ""
        strLinea = strPalabra
    Else
        strLinea = strTemp
    End If
Loop
'
Formatear = strLinea
Printer.CurrentX = (Printer.ScaleWidth / 2) - (Printer.TextWidth(Formatear) / 2)
Printer.Print UCase(Formatear)
End Function
