VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
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
      Left            =   1725
      TabIndex        =   1
      Top             =   6945
      Width           =   1215
   End
   Begin VB.CommandButton cmd 
      Caption         =   "&Aceptar"
      Height          =   495
      Index           =   0
      Left            =   3075
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
   Begin SHDocVwCtl.WebBrowser wb 
      Height          =   420
      Left            =   3555
      TabIndex        =   10
      Top             =   1275
      Visible         =   0   'False
      Width           =   795
      ExtentX         =   1402
      ExtentY         =   741
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   ""
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
Dim rstlocal As New ADODB.Recordset
Dim Mensaje As String
Dim datPer1 As Date
Dim Temporal(0 To 5) As String
Dim strDireccion As String
'
datPer1 = "01/" & cmbRfact(0) & "/" & cmbRfact(1)
datPer1 = Format(datPer1, "mm/dd/yy")
'
Mensaje = "Se dispone a enviar los avisos de cobro a los inmuebles seleccionados." & _
 vbCrLf & "¿Desea continuar este proceso?"
'
If Respuesta(Mensaje) Then

    cmd(1).Enabled = False
    With List1(1)
        If .ListCount > 0 Then
        
            Temporal(0) = mcDatos
            Temporal(1) = gcCodInm
            Temporal(2) = gcCodFondo
            Temporal(3) = gcUbica
            Temporal(4) = gcNomInm
            Temporal(5) = gnCta
            rstlocal.CursorLocation = adUseClient
            rstlocal.Open "Inmueble", cnnConexion, adOpenKeyset, adLockOptimistic, adCmdTable
            For I = 0 To (.ListCount - 1)
                rstlocal.Filter = "Codinm = '" & .List(I) & "'"
                If Not (rstlocal.EOF And rstlocal.BOF) Then
                    nEnviados = 0
                    gcCodInm = rstlocal("CodInm")
                    gcCodFondo = rstlocal("CodFondo")
                    gcUbica = rstlocal("Ubica")
                    mcDatos = gcPath + gcUbica + "Inm.mdb"
                    gcNomInm = rstlocal("Nombre")
                    If rstlocal!Caja = sysCodCaja Then
                        gnCta = CUENTA_POTE
                    Else
                        gnCta = CUENTA_INMUEBLE
                    End If
                    If Not Validar_Periodo(datPer1) Then Call Enviar_ACemail(datPer1)
                        
                    'strDireccion = "http://i.domaindlx.com/ynfantes/enviado.asp?Periodo=" & _
                    Format(datPer1, "mm/dd/yy") & "&CodInm=" & gcCodInm & "&Enviados=" & _
                    nEnviados & "&Fecha=" & Date & "&Clave=" & Format(datPer1, "mmddyy") & gcCodInm
                    
                    'Me.wb.Navigate strDireccion
                    
                    End If
                
            Next
        Else
            MsgBox "Seleccion por lo menos un inmueble de la lista", vbCritical, App.ProductName
        End If
        
    End With
    
    mcDatos = Temporal(0)
    gcCodInm = Temporal(1)
    gcCodFondo = Temporal(2)
    gcUbica = Temporal(3)
    gcNomInm = Temporal(4)
    gnCta = Temporal(5)
    cmd(1).Enabled = True
    
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
    '   Funcion:    Validar_Periodo
    '
    '   Devuelve True si el periodo seleccionado no ha sido facturado aun
    '---------------------------------------------------------------------------------------------
    Private Function Validar_Periodo(datPeriodo As Date) As Boolean
    'variables locales
    Dim rstValida As New ADODB.Recordset
    Dim cnnValida As New ADODB.Connection
    '
    cnnValida.Open cnnOLEDB & mcDatos
    rstValida.Open "SELECT * FROM Factura WHERE Fact Not LIKE 'CH*' AND Periodo=#" & datPeriodo _
    & "#;", cnnValida, adOpenStatic, adLockReadOnly
    If rstValida.RecordCount <= 1 Then
        Validar_Periodo = MsgBox("No facturado el período " _
        & UCase(Left(cmbRfact(0), 3)) & "-" & cmbRfact(1), vbInformation + vbOKOnly, "Inmueble " & gcCodInm)
    End If
    '
    Set rstValida = Nothing
    Set cnnValida = Nothing
    '
    End Function

