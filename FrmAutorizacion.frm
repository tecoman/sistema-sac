VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form FrmAutorizacion 
   Caption         =   "Autorizacion Deducciones"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   WindowState     =   2  'Maximized
   Begin MSAdodcLib.Adodc AdoTemDeducciones 
      Height          =   330
      Left            =   60
      Top             =   6765
      Visible         =   0   'False
      Width           =   4560
      _ExtentX        =   8043
      _ExtentY        =   582
      ConnectMode     =   3
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   1
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "Provider=Microsoft.Jet.OLEDB.3.51;Persist Security Info=False;Data Source=F:\sac\DATOS\2501\Inm.mdb"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.3.51;Persist Security Info=False;Data Source=F:\sac\DATOS\2501\Inm.mdb"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "SELECT * FROM TemDeducciones WHERE Autoriza =0;"
      Caption         =   "ADOTemDeducciones"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Cancelar"
      Height          =   660
      Index           =   1
      Left            =   9030
      TabIndex        =   8
      Top             =   6645
      Width           =   2040
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Aceptar"
      Height          =   660
      Index           =   0
      Left            =   6795
      TabIndex        =   7
      Top             =   6645
      Width           =   2040
   End
   Begin VB.Frame Frame3 
      Caption         =   "Detalle Deducciones"
      Height          =   4050
      Index           =   1
      Left            =   270
      TabIndex        =   9
      Top             =   2325
      Width           =   11205
      Begin MSFlexGridLib.MSFlexGrid FlexDeducciones 
         Height          =   3420
         Left            =   255
         TabIndex        =   10
         Tag             =   "1280|6850|1950"
         Top             =   390
         Width           =   10620
         _ExtentX        =   18733
         _ExtentY        =   6033
         _Version        =   393216
         Cols            =   3
         FixedCols       =   0
         ForeColor       =   -2147483646
         BackColorFixed  =   -2147483646
         ForeColorFixed  =   -2147483639
         BackColorBkg    =   -2147483636
         GridColor       =   -2147483636
         Enabled         =   0   'False
         FocusRect       =   0
         MergeCells      =   2
         FormatString    =   "^Cod.Gasto |Descripcion |Monto"
      End
   End
   Begin VB.Frame Frame3 
      Height          =   1500
      Index           =   2
      Left            =   300
      TabIndex        =   0
      Top             =   600
      Width           =   11190
      Begin VB.TextBox txt 
         Height          =   300
         Index           =   1
         Left            =   9255
         Locked          =   -1  'True
         TabIndex        =   16
         Top             =   840
         Width           =   1530
      End
      Begin VB.TextBox txt 
         Height          =   300
         Index           =   0
         Left            =   9255
         Locked          =   -1  'True
         TabIndex        =   15
         Top             =   360
         Width           =   1530
      End
      Begin MSDataListLib.DataCombo DataCombo1 
         Bindings        =   "FrmAutorizacion.frx":0000
         Height          =   315
         Index           =   1
         Left            =   1335
         TabIndex        =   5
         Top             =   810
         Width           =   1140
         _ExtentX        =   2011
         _ExtentY        =   556
         _Version        =   393216
         ListField       =   "Codigo"
         BoundColumn     =   "Nombre"
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo DataCombo1 
         Height          =   315
         Index           =   0
         Left            =   1335
         TabIndex        =   2
         Top             =   315
         Width           =   1140
         _ExtentX        =   2011
         _ExtentY        =   556
         _Version        =   393216
         ListField       =   "CodInm"
         BoundColumn     =   "Nombre"
         Text            =   ""
         Object.DataMember      =   ""
      End
      Begin MSDataListLib.DataCombo DataCombo1 
         Height          =   315
         Index           =   2
         Left            =   2490
         TabIndex        =   3
         Top             =   315
         Width           =   4890
         _ExtentX        =   8625
         _ExtentY        =   556
         _Version        =   393216
         ListField       =   "Nombre"
         BoundColumn     =   "CodInm"
         Text            =   ""
         Object.DataMember      =   ""
      End
      Begin MSDataListLib.DataCombo DataCombo1 
         Bindings        =   "FrmAutorizacion.frx":0019
         Height          =   315
         Index           =   3
         Left            =   2490
         TabIndex        =   6
         Top             =   810
         Width           =   4890
         _ExtentX        =   8625
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         ListField       =   "Nombre"
         BoundColumn     =   "CodInm"
         Text            =   ""
         Object.DataMember      =   "CmdListNombre"
      End
      Begin VB.Label Label16 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "N°:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   10
         Left            =   8805
         TabIndex        =   14
         Top             =   390
         Width           =   270
      End
      Begin VB.Label Label16 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Periodo:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   11
         Left            =   8400
         TabIndex        =   13
         Top             =   885
         Width           =   675
      End
      Begin VB.Label Label16 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Propietario:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   4
         Left            =   315
         TabIndex        =   4
         Top             =   862
         Width           =   930
      End
      Begin VB.Label Label16 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Inmueble :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   1
         Left            =   360
         TabIndex        =   1
         Top             =   367
         Width           =   885
      End
   End
   Begin MSAdodcLib.Adodc ADOcontrol 
      Height          =   330
      Left            =   120
      Top             =   1815
      Visible         =   0   'False
      Width           =   4725
      _ExtentX        =   8334
      _ExtentY        =   582
      ConnectMode     =   16
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   1
      LockType        =   3
      CommandType     =   1
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "AdoPropietario"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.Label Label16 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Total Honorarios:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   2
      Left            =   3060
      TabIndex        =   12
      Top             =   6870
      Visible         =   0   'False
      Width           =   1590
   End
   Begin VB.Label Label16 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   0
      Left            =   4695
      TabIndex        =   11
      Top             =   6825
      Visible         =   0   'False
      Width           =   1995
   End
End
Attribute VB_Name = "FrmAutorizacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'-------------------------------------------------------------------------------------------------
'Modulo de Autorización Deducciones Rev.20/08/20052
'-------------------------------------------------------------------------------------------------
'Variables Globales a nivel de módulo
Dim objRstLocal As New ADODB.Recordset
Dim StrRutaInmueble$, strPeriodo$, strMesMora$
Dim DouHonoMorosidad As Double
Dim ObjCnn As New ADODB.Connection
Public Inm$, apto$
'-------------------------------------------------------------------------------------------------


'Rev.14/08/2002------------------------------------------------------------------------------
Private Sub Command1_Click(Index As Integer)
'Botones {Aceptar/Cancelar}------------------------------------------------------------------
Dim X As Integer

Select Case Index
'
    Case 0  'Aceptar
'   --------------------
        Command1(0).Enabled = False
        Dim CnnAutoriza As New ADODB.Connection
        CnnAutoriza.Open AdoTemDeducciones.ConnectionString
        CnnAutoriza.Execute "UPDATE TemDeducciones SET Autoriza = -1,Usuario='" & gcUsuario _
        & "' WHERE IDPeriodos = '" & strPeriodo & "'", X
        With FlexDeducciones
            Call rtnBitacora("Autoriza (" & X & ") Deducciones Apto:" & DataCombo1(1) & _
            " Inmueble: " & DataCombo1(0))
        End With
        '
        If X > 0 Then
            MsgBox "Deducciones Autorizadas" & Chr(13) & "Presione Enter para continuar...", _
            vbInformation, App.ProductName
        Else
            MsgBox "Autorización no procesada", vbInformation, App.ProductName
        End If
        '
        CnnAutoriza.Close
        Set CnnAutoriza = Nothing
        Unload FrmAutorizacion
        Set FrmAutorizacion = Nothing
'
    Case 1  'Cancelar
'   ---------------------
        Unload Me
        Set FrmAutorizacion = Nothing
        
End Select
'
End Sub


'    '-----------------------------------------
'    Private Sub RtnLis(StrSource As String) '|
'    '-----------------------------------------
'    '
'    '****************************************************************************
'    '              SE LLENAN LAS LISTAS DE LOS OBJETOS DATACOMBO CON
'    '       DATOS DE LA TABLA TEMDEDUCCIONES CON VALORES QUE NO SE REPITEN
'    '****************************************************************************
'    With AdoTemDeducciones
'
'        .ConnectionString = cnnOLEDB & strRutaInmueble
'        .CommandType = adCmdText
'        .RecordSource = StrSource
'        .Refresh
'        If .Recordset.EOF Then _
'            MsgBox "No Existen Deducciones Solicitadas para este inmueble": Exit Sub
'        strPeriodo = .Recordset.Fields("IDPeriodos")
'        DataCombo1(1).Text = Mid$(strPeriodo, 3, 4)
'        Label16(13) = " " & Right(strPeriodo, 4)
'        Label16(12) = " " & .Recordset.Fields("NumFact")
'
'    End With
'    '
'    End Sub

'-------------------------------------------------
Private Sub DataCombo1_Change(Index As Integer) '|
'-------------------------------------------------
'*********************************************************
' SE SELECCIONA EL CASO DE ACUEROD AL INDICE DEL DATACOMBO
' MUESTRA EL CODIGO O EL NOMBRE DEL INMUEBLE SELECCIONADO
'*********************************************************
If Len(DataCombo1(Index)) < 4 Then Exit Sub
Select Case Index
    
    Case 0  'CUANDO CAMBIA EL VALOR COD.INMUEBLE
        
        DataCombo1(2) = DataCombo1(0).BoundText
        
    Case 1  'CUANDO CAMBIA EL VALOR COD.PROPIETARIO
        DataCombo1(3) = DataCombo1(1).BoundText
    
    Case 2  'CUANDO CAMBIA EL VALOR NOMBRE DEL INMUEBLE
        DataCombo1(0) = DataCombo1(2).BoundText
End Select
End Sub

    '----------------------------------------------
    Private Sub RtnAsignaCampos(rst As ADODB.Recordset) '|
    '----------------------------------------------
    '
    
    With rst
        
        If .EOF Then MsgBox "Veirifique Codigo de Inmueble...": Exit Sub
        DataCombo1(0) = .Fields("CodInm")
        DataCombo1(2) = .Fields("Nombre")
        StrRutaInmueble = gcPath & .Fields("Ubica") & "inm.mdb"
        strMesMora = .Fields("MesesMora")
        DouHonoMorosidad = .Fields("HonoMorosidad")
        
    End With
    '
    End Sub

    
'-----------------------------------------------------------------
Private Sub DataCombo1_Click(Index As Integer, Area As Integer) '|
'-----------------------------------------------------------------
If Area = 2 Then   'click en un elemento de la lista

    Select Case Index
        Case 0, 2
            Call RtnInmueble("CodInm", DataCombo1(0))
        Case 1, 3
            Call RtnBuscaDeduccion
    End Select
End If
End Sub

    '------------------------------------------------------------------------
    Private Sub DataCombo1_KeyPress(Index As Integer, KeyAscii As Integer) '|
    '------------------------------------------------------------------------
    KeyAscii = Asc(UCase(Chr(KeyAscii)))    'Convierte todo a mayusculas

    If KeyAscii = 13 Then
        Select Case Index
        '
        Case 0  'CODIGO DE INMUEBLE
        '
        If DataCombo1(0) = "" Then DataCombo1(2).SetFocus: Exit Sub
        If Len(DataCombo1(0)) = 4 Then 'SI EL CODIGO DE INMUEBLE = 4 DIGITOS ENTONCES
        '
            Call RtnInmueble("CodInm", DataCombo1(0))
        '
        End If
        '
        Case 1  'CODIGO DE PROPIETARIO
        '
            Call RtnBuscaDeduccion
        '
        Case 2  'NOMBRE DEL INMUEBLE
        '
            Call RtnInmueble("Inmueble.Nombre", DataCombo1(0))
        '
        End Select
        '
    End If
    End Sub

    '-------------------------
    Private Sub Form_Load() '|
    '-------------------------
    'variables locales
    Dim strSql$
    '
    strSql = "SELECT CodInm,Nombre FROM Inmueble ORDER BY CodInm"
    objRstLocal.Open strSql, cnnConexion, adOpenKeyset, adLockReadOnly, adCmdText
    Set DataCombo1(0).RowSource = objRstLocal
    Set objRstLocal = New ADODB.Recordset
    strSql = "SELECT CodInm,Nombre FROM Inmueble ORDER BY Nombre"
    objRstLocal.Open strSql, cnnConexion, adOpenKeyset, adLockReadOnly, adCmdText
    Set DataCombo1(2).RowSource = objRstLocal
    '
    Command1(0).Enabled = False
    Call centra_titulo(FlexDeducciones, True)
    
    If Inm <> "" And apto <> "" Then
    
        DataCombo1(0) = Inm
        Call DataCombo1_Click(0, 2)
        DataCombo1(1) = apto
        Call DataCombo1_Click(1, 2)
        Show
        Inm = "": apto = ""
        
    End If
    
'    Set objRst = New ADODB.Recordset
'    objRst.Open "Sol_autorizacion", cnnConexion, adOpenKeyset, adLockOptimistic, adCmdTable
'    If Not objRst.EOF Or Not objRst.BOF Then
'        objRst.MoveFirst
'        DataCombo1(0) = objRst!CodInm
'        DataCombo1_Click 0, 2
'        DataCombo1(2) = objRst!apto
'
'    End If
    End Sub

    '--------------------------------
    Public Sub RtnBuscaDeduccion() '|
    '--------------------------------
        
    If DataCombo1(1) = "" Then
        MsgBox "Debe Introducir un Valor..", vbInformation, App.ProductName
        DataCombo1(1).SetFocus
        Exit Sub
    End If
    AdoTemDeducciones.ConnectionString = ADOcontrol.ConnectionString
    AdoTemDeducciones.RecordSource = _
        "SELECT * FROM TemDeducciones WHERE Idperiodos LIKE '" _
        & Right(DataCombo1(0), 2) + DataCombo1(1) & "%' AND Autoriza=FALSE"
    AdoTemDeducciones.Refresh
    With AdoTemDeducciones.Recordset
        i = 1: FlexDeducciones.Rows = .RecordCount + 1
        If Not .EOF Then
            strPeriodo = .Fields("IDPeriodos")
            txt(0) = " " & IIf(IsNull(.Fields("NumFact")), "", .Fields("NumFact"))
            If IsDate("01/" & Left(strPeriodo, 2)) Then
                txt(1) = " " & UCase(Format(CDate("01/" & Left(strPeriodo, 2)), "MMM-YY"))
            Else
                txt(1) = " " & Left(Right(strPeriodo, 4), 2) & "-" & _
                Right(Right(strPeriodo, 4), 2)
            End If
            
            .MoveFirst
            
            Do
                
                With FlexDeducciones
                    'SI LA DEDUCCION ES INTERESES DE MORA MUESTRA EL TOTAL EN PANTALLA
                    If AdoTemDeducciones.Recordset.Fields("Titulo") Like "*HONORARIOS*ABOGADO*" Then
                        Set ObjCnn = New ADODB.Connection
                        ObjCnn.Open AdoTemDeducciones.ConnectionString
                        Set objRstLocal = New ADODB.Recordset
                        objRstLocal.Open "SELECT Count(Factura.Periodo) AS Cuenta, Sum(Factura.Saldo" _
                        & ") AS Deuda, Factura.codprop FROM Factura WHERE (((Factura.codprop)='" _
                        & DataCombo1(1) & "') and ((Factura.saldo) <> 0)) GROUP BY Factura.codp" _
                        & "rop", ObjCnn, adOpenKeyset, adLockBatchOptimistic, adCmdText
                        
                        If objRstLocal.Fields("Cuenta") > CInt(strMesMora) Then
                            Label16(2).Visible = True
                            Label16(0).Visible = True
                            'Asigna el resultado a la etiqueta
                            Label16(0) = Format(CLng(objRstLocal.Fields("Deuda") _
                                * DouHonoMorosidad / 100), "#,##0.00")
                        End If
                        objRstLocal.Close
                        Set objRstLocal = Nothing
                        ObjCnn.Close
                        Set ObjCnn = Nothing
                        
                    End If
                    'distribuye la información de c/registro
                    .TextMatrix(i, 0) = AdoTemDeducciones.Recordset("CodGasto")
                    .TextMatrix(i, 1) = AdoTemDeducciones.Recordset("Titulo")
                    .TextMatrix(i, 2) = Format(AdoTemDeducciones.Recordset("Monto"), "#,##0.00")
                    i = i + 1: AdoTemDeducciones.Recordset.MoveNext
                End With
            Loop Until .EOF
            Command1(0).Enabled = True
            If apto = "" And Inm = "" Then Command1(1).SetFocus
        Else
        
            MsgBox "El Propietario del '" & DataCombo1(1) & "' No Tiene Solicitud Deducciones ", _
            vbInformation, App.ProductName
            
            'DataCombo1(1).SetFocus
            
        End If
        
    End With
    AdoTemDeducciones.Recordset.Close
End Sub

'rutina que busca la información del inmuble seleccionado-----------------------------------------
Private Sub RtnInmueble(StrCampo As String, datControl As DataCombo)
'-------------------------------------------------------------------------------------------------
'
'Dim REg As Long

With FrmAdmin.objRst
    'If .State = 0 Then .Open
    Reg = .Bookmark
    .MoveFirst
    .Find StrCampo & " LIKE '%" & datControl & "%'"
    If Not .EOF Then
        Call RtnAsignaCampos(FrmAdmin.objRst)
        '.Move REg
    Else
        MsgBox "No se encuentra la información del inmueble '" & datControl & "'", _
        vbExclamation, App.ProductName
    End If
End With
ADOcontrol.ConnectionString = cnnOLEDB & StrRutaInmueble
ADOcontrol.RecordSource = "Select * from Propietarios order by codigo"
ADOcontrol.Refresh
If apto = "" And Inm = "" Then DataCombo1(1).SetFocus
'
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
objRstLocal.Close
Set objRstLocal = Nothing
ObjCnn.Close
Set ObjCnn = Nothing
End Sub
