VERSION 5.00
Object = "{20C62CAE-15DA-101B-B9A8-444553540000}#1.1#0"; "MSMAPI32.OCX"
Begin VB.Form frmSelecInm 
   Caption         =   "Enviar Avisos de Cobro vía mail"
   ClientHeight    =   7680
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4515
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   7680
   ScaleWidth      =   4515
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame fraRfact 
      Caption         =   "Periodo ( Mes - Año )"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1110
      Left            =   495
      TabIndex        =   7
      Top             =   195
      Width           =   3405
      Begin VB.ComboBox cmbRfact 
         DataField       =   "TipoMovimientoCaja"
         Height          =   315
         Index           =   1
         ItemData        =   "frmSelecInm.frx":0000
         Left            =   1875
         List            =   "frmSelecInm.frx":0002
         Sorted          =   -1  'True
         TabIndex        =   9
         Text            =   "cmbRfact"
         Top             =   435
         Width           =   1020
      End
      Begin VB.ComboBox cmbRfact 
         DataField       =   "TipoMovimientoCaja"
         Height          =   315
         Index           =   0
         ItemData        =   "frmSelecInm.frx":0004
         Left            =   465
         List            =   "frmSelecInm.frx":002F
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   435
         Width           =   1320
      End
   End
   Begin VB.ListBox List1 
      Height          =   3960
      Index           =   0
      ItemData        =   "frmSelecInm.frx":0098
      Left            =   165
      List            =   "frmSelecInm.frx":009A
      Sorted          =   -1  'True
      TabIndex        =   3
      Top             =   2610
      Width           =   1815
   End
   Begin VB.ListBox List1 
      Height          =   3960
      Index           =   1
      ItemData        =   "frmSelecInm.frx":009C
      Left            =   2385
      List            =   "frmSelecInm.frx":009E
      Sorted          =   -1  'True
      TabIndex        =   2
      Top             =   2655
      Width           =   1815
   End
   Begin VB.CommandButton cmd 
      Caption         =   "&Cancelar"
      Height          =   495
      Index           =   1
      Left            =   3075
      TabIndex        =   1
      Top             =   6945
      Width           =   1215
   End
   Begin VB.CommandButton cmd 
      Caption         =   "&Aceptar"
      Height          =   495
      Index           =   0
      Left            =   1725
      TabIndex        =   0
      Top             =   6960
      Width           =   1215
   End
   Begin MSMAPI.MAPIMessages MAPm 
      Left            =   870
      Top             =   6900
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      AddressEditFieldCount=   1
      AddressModifiable=   0   'False
      AddressResolveUI=   0   'False
      FetchSorted     =   0   'False
      FetchUnreadOnly =   0   'False
   End
   Begin MSMAPI.MAPISession MAPs 
      Left            =   240
      Top             =   6885
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DownloadMail    =   -1  'True
      LogonUI         =   -1  'True
      NewSession      =   0   'False
   End
   Begin VB.Label lbl 
      Caption         =   "2.- Selección Inmueble"
      Height          =   270
      Index           =   2
      Left            =   2475
      TabIndex        =   6
      Top             =   2325
      Width           =   1800
   End
   Begin VB.Label lbl 
      Caption         =   "1.- Lista Inmuebles"
      Height          =   270
      Index           =   1
      Left            =   180
      TabIndex        =   5
      Top             =   2325
      Width           =   1800
   End
   Begin VB.Label lbl 
      Caption         =   "Seleccione de la lista los inmuebles a los cuales desea enviarle los avisos de cobro vía e-mail:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   0
      Left            =   180
      TabIndex        =   4
      Top             =   1620
      Width           =   4140
   End
End
Attribute VB_Name = "frmSelecInm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmd_Click(Index As Integer)

Select Case Index
    Case 1  'cancelar
        Unload Me
        Set frmSelecInm = Nothing
        
    Case 0  'aceptar
        Call Enviar_Avisos
End Select

End Sub

Private Sub Form_Load()
Dim rstlocal As New ADODB.Recordset
CenterForm Me
For I = 0 To Year(Date) - 2003: cmbRfact(1).AddItem (2003 + I)
Next
'Presenta el periodo al mes actual
cmbRfact(0).Text = cmbRfact(0).List(Month(Date) - 1)
cmbRfact(1).Text = Year(Date)
    
    
rstlocal.CursorLocation = adUseClient
rstlocal.Open "Inmueble", cnnConexion, adOpenKeyset, adLockOptimistic, adCmdTable
rstlocal.Sort = "CodInm"
With rstlocal
    If Not (.EOF And .BOF) Then
        .MoveFirst
        Do
            List1(0).AddItem !CodInm
            .MoveNext
        Loop Until .EOF
    End If
    .Close
End With
Set rstlocal = Nothing

End Sub

Private Sub Enviar_Avisos()
'variables locales
Dim cPropieatrios As Collection
Dim propietario() As Variant
Dim rst As ADODB.Recordset
Dim Mensaje As String
Dim datPer1 As Date
'
datPer1 = "01/" & cmbRfact(0) & "/" & cmbRfact(1)
datPer1 = Format(datPer1, "mm/dd/yy")
'
Mensaje = "Se dispone a enviar los avisos de cobro a los inmuebles seleccionados." & _
 vbCrLf & "¿Desea continuar este proceso?"
'
If Respuesta(Mensaje) Then

    cmd(1).Enabled = False
    cmd(0).Enabled = False
    With List1(1)
        
        If .ListCount > 0 Then
            'If Not Validar_Periodo(datPer1) Then Call Enviar_ACemail(datPer1)
            
            
                Set cPropietarios = New Collection
                I = 0
                Set rst = New ADODB.Recordset
                
                Do
                    
                    If periodoEsValido(datPer1, .List(I)) Then
                    
                        cnn = cnnOLEDB & gcPath & "\" & .List(I) & "\inm.mdb"
                        
                        sql = "SELECT * FROM Propietarios WHERE email <>'' AND Demanda = False;"
                        rst.Open sql, cnn, adOpenKeyset, adLockOptimistic, adCmdText
                        
                        If Not (rst.EOF And rst.BOF) Then
                            Do
                                cPropietarios.Add (.List(I) & "|" & rst("codigo") & "|" & _
                                rst("nombre") & "|" & rst("email") & "|" & Format(datPer1, "dd/mm/yy") & _
                                "|" & Format(datPer1, "ddyy") + Right(.List(I), 3) + Format(rst("ID"), "000") & _
                                "|" & rst("codigo"))
                                rst.MoveNext
                            Loop Until rst.EOF
                        End If
                        rst.Close
                        
                    End If
                    I = I + 1
                                
                Loop Until .ListCount = I
                
                Set rst = Nothing
                If cPropietarios.Count > 0 Then
                    Set frmCalendario.cPropietarios = cPropietarios
                    frmCalendario.iniciarTemporizador 5000
                    Unload Me
                    Set frmselectinm = Nothing
                End If
            
        Else
            MsgBox "Seleccion por lo menos un inmueble de la lista", vbCritical, App.ProductName
        End If
        
    End With
    
    cmd(1).Enabled = True
    cmd(0).Enabled = True
End If
'
End Sub

Private Sub List1_Click(Index As Integer)
Select Case Index
    Case 0
        List1(1).AddItem List1(0).Text
        List1(0).RemoveItem List1(0).ListIndex
    Case 1
        List1(0).AddItem List1(1).Text
        List1(1).RemoveItem List1(1).ListIndex
End Select
End Sub

 '---------------------------------------------------------------------------------------------
    '   Funcion:    periodoEsValido
    '
    '   Devuelve True si el periodo seleccionado no ha sido facturado aun
    '---------------------------------------------------------------------------------------------
    Private Function periodoEsValido(Periodo As Date, Inmueble As String) As Boolean
    'variables locales
    Dim rstValida As New ADODB.Recordset
    Dim cnn As String
    '
    cnn = cnnOLEDB & gcPath & "\" & Inmueble & "\inm.mdb"
    rstValida.Open "SELECT * FROM Factura WHERE Fact Not LIKE 'CH*' AND Periodo=#" & Periodo _
    & "#;", cnn, adOpenStatic, adLockReadOnly
    
    If (rstValida.EOF And rstValida.BOF) Then
        periodoEsValido = MsgBox("Para el inmueble " & Inmueble & " no se ha facturado el período: " _
        & UCase(Left(cmbRfact(0), 3)) & "-" & cmbRfact(1), vbInformation + vbOKOnly, "Inmueble " & Inmueble)
    End If
    
    periodoEsValido = Not periodoEsValido
    '
    rstValida.Close
    Set rstValida = Nothing
    '
    End Function

