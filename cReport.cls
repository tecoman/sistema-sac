VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 1  'vbSimpleBound
  DataSourceBehavior  = 1  'vbDataSource
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ctlReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Attribute VB_Ext_KEY = "Member0" ,"Item"
Option Explicit
'variables locales para almacenar los valores de las propiedades
Private mReport As CRAXDRT.Report   'copia local
Private mApp As CRAXDRT.Application 'copia local
Private mFrm As frmView 'copia local
Private mvarReporte As String 'copia local
Private mvarSalida As crSalida  'copia local
Private mvarFormulas() As String 'copia local
Private mvarParametros() 'copia local
Private mvarOrigenDatos() As String 'copia local
Private mvarTituloVentana As String 'copia local
Private mvarFormuladeSeleccion As String 'copia local
Private mvarArbol As Boolean    'copia local
Private mvarZoom As Integer     'copia local
Private Formula As Boolean
Private Parametro As Boolean
Private Origen As Boolean, orden As Boolean
Public Event Error(ID As Long, Descripcion As String)  'error local
Private mvarArchivoSalida As String 'copia local
Private mvarOrdenRegistros() As String 'copia local
'variables locales para almacenar los valores de las propiedades
Private mvarFormatoSalida As CRExportFormatType 'copia local


Public Property Let FormatoSalida(ByVal vData As CRExportFormatType)
    mvarFormatoSalida = vData
End Property

Public Property Get FormatoSalida() As CRExportFormatType
    FormatoSalida = mvarFormatoSalida
End Property

Public Property Let OrdenRegistros(Index As Integer, ByVal vData As String)
If Index = 0 Then
    ReDim mvarOrdenRegistros(0)
Else
    ReDim Preserve mvarOrdenRegistros(UBound(mvarOrdenRegistros) + 1)
End If
mvarOrdenRegistros(Index) = vData
orden = True
End Property


Public Property Get OrdenRegistros(Index As Integer) As String
OrdenRegistros = mvarOrdenRegistros(Index)
End Property

Public Property Let ArchivoSalida(ByVal vData As String)
    mvarArchivoSalida = vData
End Property

Public Property Get ArchivoSalida() As String
    ArchivoSalida = mvarArchivoSalida
End Property

Public Property Let Zoom(ByVal vData As Integer)
If vData = 0 Then
    mvarZoom = 1
Else
    mvarZoom = vData
End If
End Property

Public Property Get Zoom() As Integer
Zoom = mvarZoom
End Property

Public Property Let ArbolGrupo(ByVal vData As Boolean)
mvarArbol = vData
End Property

Public Property Get ArbolGrupo() As Boolean
ArbolGrupo = mvarArbol
End Property

Public Property Let FormuladeSeleccion(ByVal vData As String)
mvarFormuladeSeleccion = vData
End Property

Public Property Get FormuladeSeleccion() As String
FormuladeSeleccion = mvarFormuladeSeleccion
End Property


Public Property Let TituloVentana(ByVal vData As String)
mvarTituloVentana = vData
End Property


Public Property Get TituloVentana() As String
TituloVentana = mvarTituloVentana
End Property

Public Property Let OrigenDatos(Index As Integer, ByVal vData As String)
If Index = 0 Then
    ReDim mvarOrigenDatos(0)
Else
    ReDim Preserve mvarOrigenDatos(UBound(mvarOrigenDatos) + 1)
End If
mvarOrigenDatos(Index) = vData
Origen = True
End Property

Public Property Get OrigenDatos(Index As Integer) As String
OrigenDatos = mvarOrigenDatos(Index)
End Property

Public Property Let Formulas(Index As Integer, ByVal vData As String)
If Index = 0 Then
    ReDim mvarFormulas(0)
Else
    ReDim Preserve mvarFormulas(UBound(mvarFormulas) + 1)
End If
mvarFormulas(Index) = vData
Formula = True
End Property

Public Property Get Formulas(Index As Integer) As String
Formulas = mvarFormulas(Index)
End Property

Public Property Let Parametros(Index As Integer, ByVal vData)
If Index = 0 Then
    ReDim mvarParametros(0)
Else
    ReDim Preserve mvarParametros(UBound(mvarParametros) + 1)
End If
mvarParametros(Index) = vData
Parametro = True
End Property

Public Property Get Parametros(Index As Integer)
Parametros = mvarParametros(Index)
End Property

Public Property Let Salida(ByVal vData As crSalida)
mvarSalida = vData
End Property

Public Property Get Salida() As crSalida
Salida = mvarSalida
End Property

Public Property Let Reporte(ByVal vData As String)
If Dir$(vData) = "" Then
    RaiseEvent Error(48, "No se encuentra el reporte " & vData)
Else
    mvarReporte = vData
End If
End Property

Public Property Get Reporte() As String
Attribute Reporte.VB_UserMemId = 0
Reporte = mvarReporte
End Property

Public Function Imprimir(Optional Copias As Long, Optional DataSource As ADODB.Recordset) As Long
'variables locales
Dim I As Integer, P As Integer, j As Integer
Dim mFormulasF As CRAXDRT.FormulaFieldDefinitions
Dim mFormulaF As CRAXDRT.FormulaFieldDefinition
Dim crParamDefs As CRAXDRT.ParameterFieldDefinitions
Dim crParamDef As CRAXDRT.ParameterFieldDefinition

'
On Error Resume Next

If Reporte = "" Then Exit Function
Set mReport = mApp.OpenReport(Reporte, 1)

If Not mReport.HasSavedData Then
    'si no recive un ADODB.Recordset
    If Origen Then
        For I = 0 To UBound(mvarOrigenDatos)
            mReport.Database.Tables(I + 1).Location = OrigenDatos(I)
        Next
    End If
    If mvarFormuladeSeleccion <> "" Then mReport.RecordSelectionFormula = FormuladeSeleccion

    If Formula Then
        Set mFormulasF = mReport.FormulaFields
        For I = 0 To UBound(mvarFormulas)
            P = InStr(mvarFormulas(I), "=") - 1
            For Each mFormulaF In mFormulasF
                If mvarFormulas(I) <> "" Then
                    If UCase(mFormulaF.FormulaFieldName) = Trim(UCase(Left(mvarFormulas(I), P))) Then
                        mFormulaF.Text = Mid(Formulas(I), P + 2, Len(Formulas(I)))
                    End If
                End If
            Next
        Next
    End If
    'parametros
    If Parametro Then
        Set crParamDefs = mReport.ParameterFields
        I = 0
        For Each crParamDef In crParamDefs
            crParamDef.AddCurrentValue (Parametros(I))
            I = I + 1
        Next
    End If
    
End If

If orden Then
    Dim DefCampo As CRAXDRT.DatabaseFieldDefinition
    Dim Tabla As String, campo As String, Direccion As String
    
    For I = 0 To UBound(mvarOrdenRegistros)
        Direccion = Left(mvarOrdenRegistros(I), 1)
        Tabla = UCase(Mid(mvarOrdenRegistros(I), 2, InStr(mvarOrdenRegistros(I), ".") - 2))
        campo = Mid(mvarOrdenRegistros(I), InStr(mvarOrdenRegistros(I), ".") + 1, 100)
        
        For j = 1 To mReport.Database.Tables.Count
            If UCase(mReport.Database.Tables(j).Name) = Tabla Then
                Set DefCampo = mReport.Database.Tables(j).Fields.GetItemByName(campo)
                mReport.RecordSortFields.Add DefCampo, crAscendingOrder
            End If
        Next
        
    Next
End If
'If Not IsMissing(DataSource) Then
'    mReport.Database.SetDataSource DataSource, 3, 1
'End If
If Salida = 1 Then
    mReport.PrintOut False, IIf(Copias = 0, 1, Copias)
ElseIf Salida = crPantalla Then
    mFrm.Caption = TituloVentana
    mFrm.crView.ReportSource = mReport
    mFrm.crView.ViewReport
    While mFrm.crView.IsBusy
        DoEvents
    Wend
    mFrm.crView.DisplayGroupTree = mvarArbol
    mFrm.crView.EnableGroupTree = mvarArbol
    mFrm.crView.Zoom (IIf(mvarZoom = 0, 1, mvarZoom))
    mFrm.Show
    mFrm.Visible = True
ElseIf Salida = crArchivoDisco Then
    mReport.ExportOptions.DestinationType = crEDTDiskFile
    'crEFTCrystalReport
    'crEFTPortableDocFormat
    mReport.ExportOptions.FormatType = mvarFormatoSalida
    mReport.ExportOptions.DiskFileName = mvarArchivoSalida
    mReport.Export False
End If
If Err <> 0 And Err <> 401 Then
    Imprimir = Err 'MsgBox("Error al imprimir el reporte:" & _
    Err.Description, vbCritical, "Error " & Err)
    RaiseEvent Error(Err, "Ocurrio el siguiente error al imprimir el reporte:" & vbCrLf & _
    Err.Description)
End If
'
End Function

Private Sub Class_Initialize()
'crea una instancia de los objetos
Set mReport = New CRAXDRT.Report
Set mApp = New CRAXDRT.Application
Set mFrm = New frmView

End Sub


Private Sub Class_Terminate()
'elimina de memoria los objeto creador
Set mReport = Nothing
Set mApp = Nothing
Set mFrm = Nothing
End Sub
