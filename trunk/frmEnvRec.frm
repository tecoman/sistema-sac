VERSION 5.00
Begin VB.Form frmEnvRec 
   AutoRedraw      =   -1  'True
   Caption         =   "Imprimir recibos por enviar..."
   ClientHeight    =   3555
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6225
   Icon            =   "frmEnvRec.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3555
   ScaleWidth      =   6225
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmd 
      Height          =   345
      Index           =   3
      Left            =   1530
      Picture         =   "frmEnvRec.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   1620
      Width           =   375
   End
   Begin VB.CommandButton cmd 
      Height          =   345
      Index           =   2
      Left            =   1530
      Picture         =   "frmEnvRec.frx":0588
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   1965
      Width           =   375
   End
   Begin VB.ListBox list 
      Height          =   2205
      Index           =   1
      Left            =   1920
      Sorted          =   -1  'True
      TabIndex        =   6
      Top             =   885
      Width           =   1215
   End
   Begin VB.ListBox list 
      Height          =   2205
      Index           =   0
      Left            =   255
      Sorted          =   -1  'True
      TabIndex        =   5
      Top             =   885
      Width           =   1215
   End
   Begin VB.CommandButton cmd 
      Caption         =   "&Aceptar"
      Height          =   495
      Index           =   1
      Left            =   4800
      TabIndex        =   1
      Top             =   2235
      Width           =   1215
   End
   Begin VB.CommandButton cmd 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   495
      Index           =   0
      Left            =   4800
      TabIndex        =   0
      Top             =   2835
      Width           =   1215
   End
   Begin VB.Label lbl 
      Alignment       =   2  'Center
      Caption         =   "Seleccion"
      Height          =   285
      Index           =   2
      Left            =   1965
      TabIndex        =   4
      Top             =   360
      Width           =   1515
   End
   Begin VB.Label lbl 
      Alignment       =   2  'Center
      Caption         =   "Inmuebles"
      Height          =   270
      Index           =   1
      Left            =   255
      TabIndex        =   3
      Top             =   390
      Width           =   1500
   End
   Begin VB.Label lbl 
      Caption         =   "Selecciones de la lista de inmueble los inmuebles de los cuales quiere imprimir los recibos de pago."
      Height          =   1635
      Index           =   0
      Left            =   3690
      TabIndex        =   2
      Top             =   330
      Width           =   2280
   End
End
Attribute VB_Name = "frmEnvRec"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Cmd_Click(Index As Integer)
'variables locales
Dim Respuesta As Integer
Dim FP(2) As String
Dim rpReporte As ctlReport
Dim rstPago As ADODB.Recordset
Dim errLocal As Long
'--------------------------------------
Select Case Index

    Case 0  'cerra formulario
        Me.Hide
        
    Case 1 'procesar
    If List(1).List(0) = "" Then
        MsgBox "Debe seleccionar por lo menos un inmueble de la lista", vbInformation, _
        App.ProductName
        Exit Sub
    End If
    
    Respuesta = MsgBox("Desea llevar a cabo este proceso", vbYesNo + vbQuestion, App.ProductName)
            
    If Respuesta = vbYes Then
    
        cmd(0).Enabled = False
        cmd(1).Enabled = False
        lbl(0) = "Iniciando Proceso, Espere un momento por favor...."
        
        'crea una instancia del objeto ADODB.Recordset
        Set rstPago = New ADODB.Recordset

        rstPago.Open "qdfRE", cnnConexion, adOpenKeyset, adLockOptimistic, adCmdTable
        
        '
        With rstPago
        
            .Filter = strFiltro
            .Sort = "InmuebleMovimientoCaja, AptoMovimientoCaja"
            '
            If Not .EOF And Not .BOF Then
                'Set ctlReport = FrmAdmin.rptReporte
                .MoveFirst
                Do  'imprime los recibos de pago hasta fin de archivo
                    
                    If Not !Print Then
                        
                        lbl(0) = "Inm. " & !InmuebleMovimientoCaja & " Imprimiendo Fact. '" & _
                        !Fact & "'"
                        DoEvents
                        FP(0) = !FPago & " - " & !NumDocumentoMovimientoCaja & " - " & _
                        !BancoDocumentoMovimientoCaja
                        FP(1) = !Fpago1 & " - " & !NumDocumentoMovimientoCaja1 & " - " & _
                        !BancoDocumentoMovimientoCaja1
                        FP(2) = !Fpago2 & " - " & !NumDocumentoMovimientoCaja2 & " - " & _
                        !BancoDocumentoMovimientoCaja2
                        If Right(!InmuebleMovimientoCaja, 3) = Mid(!Fact, 5, 3) Then
                            Call Printer_Pago(!Fact, !Monto, !Ubica, !InmuebleMovimientoCaja, _
                            !Nombre, !IDRecibo, True, 2, crImpresora, FP(0), FP(1), FP(2))
                        Else
                            SetTimer hWnd, NV_CLOSEMSGBOX, 2000&, AddressOf TimerProc
                            Call MessageBox(hWnd, "No se puede imprimir recibo #" & !Fact, _
                            App.ProductName, vbInformation)
                            Call rtnBitacora("No se imprimió el recibo #" & !Fact)
                        End If
                        DoEvents
'                        'marca el recibo como impreso
'                        cnnConexion.Execute "UPDATE Recibos_Enviar SET Print=True WHERE" _
'                        & " IDRecibo='" & !IDrecibo & "' AND Fact='" & !Fact & "';"

                    End If
                    .MoveNext
                Loop Until .EOF
                'imprimir el reporte final
                .Close
                Set rpReporte = New ctlReport
                rpReporte.Reporte = gcReport & "fact_pagoenv.rpt"
                rpReporte.OrigenDatos(0) = gcPath & "\sac.mdb"
                rpReporte.TituloVentana = "Recibos por Enviar"
                rpReporte.Salida = crImpresora
                errLocal = rpReporte.Imprimir
                Set rpReporte = Nothing
                Call rtnBitacora("Imprimir Recibos pendientes por enviar")
                If errLocal <> 0 Then
                    MsgBox Err.Description, vbCritical, _
                    Err.Number
                    Call rtnBitacora("Ocurrió el error [" & Err.Number & _
                    "] al imprimir el reporte")
                    lbl(0) = "El proceso generó algunos errores durante su ejecución"
                Else
                    cnnConexion.Execute "DELETE * FROM Recibos_Enviar WHERE Print=True;"
                    MsgBox "El proceso ha finalizado con éxito", vbInformation, _
                    App.ProductName
                    lbl(0) = "El proceso finalizó con éxito"
                End If
                cmd(0).Enabled = True
                cmd(1).Enabled = True
                '
            Else
                cmd(0).Enabled = True
                cmd(1).Enabled = True
                lbl(0) = "No existen recibos que imprimir..."
            End If
        '
        End With
        Set rstPago = Nothing
        '
    End If
    
    Case 3, 2
        Dim K As Integer, J As Integer
        
        K = IIf(Index = 3, 0, 1)
        J = IIf(K = 0, 1, 0)
        If List(K).List(0) <> "" Then
            Do
                List(J).AddItem List(K).List(0)
                List(K).RemoveItem (0)
            Loop Until List(K).List(0) = ""
            
        End If
End Select

End Sub

Private Sub Form_Load()
'variables locales
Dim rst As New ADODB.Recordset
'
rst.Open "SELECT * FROM Inmueble WHERE Inactivo = False", cnnConexion, adOpenKeyset, _
adLockOptimistic, adCmdText
If Not rst.EOF And Not rst.BOF Then
    rst.MoveFirst
    Do
        List(0).AddItem rst("CodInm")
        rst.MoveNext
    Loop Until rst.EOF
End If
'cierra i descarga el objeto
rst.Close
Set rst = Nothing

End Sub

Private Sub list_DblClick(Index As Integer)
'
Select Case Index

    Case 0
        If List(0).Text <> "" Then
            List(1).AddItem List(0).Text
            List(0).RemoveItem (List(0).ListIndex)
        End If
        
    Case 1
        If List(1).Text <> "" Then
            List(0).AddItem List(1).Text
            List(1).RemoveItem (List(1).ListIndex)
        End If
End Select
'
End Sub

Private Function strFiltro() As String
'variales locales
'
With List(1)

    strFiltro = "InmuebleMovimientoCaja ='" & .List(0) & "'"
    For i = 1 To .ListCount - 1
        strFiltro = strFiltro & " Or InmuebleMovimientoCaja ='" & .List(i) & "'"
    Next
    
End With
'
End Function
