VERSION 5.00
Begin VB.Form frmRevNomina 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Reversar Nómina"
   ClientHeight    =   3270
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4170
   ControlBox      =   0   'False
   Icon            =   "frmRevNomina.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3270
   ScaleWidth      =   4170
   Begin VB.CommandButton Command1 
      Caption         =   "&Salir"
      Height          =   495
      Index           =   1
      Left            =   2760
      TabIndex        =   3
      Top             =   2400
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Reversar"
      Height          =   495
      Index           =   0
      Left            =   1440
      TabIndex        =   2
      Top             =   2400
      Width           =   1215
   End
   Begin VB.Frame fra 
      Caption         =   "Seleccione la nómina:"
      Height          =   1095
      Left            =   240
      TabIndex        =   1
      Top             =   1080
      Width           =   3735
      Begin VB.ComboBox cmb 
         Height          =   315
         Left            =   240
         TabIndex        =   4
         Top             =   480
         Width           =   3255
      End
   End
   Begin VB.Label lbl 
      Caption         =   "Este proceso eliminará la información relacionada con la nónima seleccionada. Registros de facturación y de pago al personal."
      Height          =   735
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   3735
   End
End
Attribute VB_Name = "frmRevNomina"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim IDNomina As Long
Dim vecNom(2) As Long


Private Sub cmb_Click()
IDNomina = vecNom(cmb.ListIndex)
End Sub

Private Sub Command1_Click(Index As Integer)
Select Case Index
    
    Case 1 'cerrar
        Unload Me
    Case 0 'reversar
        If valida_datos Then
            
            If reversar_nomina(IDNomina) Then
                MsgBox "Nómina " & cmb.Text & " reversada con éxito."
                Unload Me
            End If
            
        End If
        
End Select

End Sub

Private Sub Form_Load()

Dim strSQL As String
Dim strNom As String
Dim FechaU As Date
Dim ultNomPro As Long
Dim rstlocal As ADODB.Recordset

Set rstlocal = New ADODB.Recordset

strSQL = "SELECT IDNomina as UNP FROM Nom_Inf WHERE Efectivo = (SELECT Max(Efectivo) " _
        & "FROM Nom_Inf) AND IDNomina <> 312" & Year(Date)

CenterForm Me

rstlocal.Open strSQL, cnnConexion, adOpenKeyset, adLockReadOnly, adCmdText

If Not (rstlocal.EOF And rstlocal.BOF) Then

    ultNomPro = rstlocal("UNP")
    
    
    
    FechaU = DateSerial(Right(ultNomPro, 4), Mid(ultNomPro, 2, 2), _
    IIf(Left(ultNomPro, 1) = 1, 15, 29))
    
    J = Left(ultNomPro, 1)
    
    For i = 0 To 2
        'agrega las nominas anteriores
        If Day(FechaU) = 15 Then
            strNom = UCase(J & "  quincena mes: " & Format(DateAdd("d", -1, _
            DateAdd("m", 1 * (i * -1), "01/" & Format(FechaU, "mm/yy"))), "mmm-yyyy"))
        Else
            strNom = UCase(J & "  quincena mes: " & Format(DateAdd("d", _
            15 * (i * -1), FechaU), "mmm-yyyy"))
        End If
        cmb.AddItem strNom
        
        If Day(FechaU) = 15 Then
            vecNom(i) = J & Format(DateAdd("d", -1, _
            DateAdd("m", 1 * (i * -1), "01/" & Format(FechaU, "mm/yy"))), "mmyyyy")
        Else
            vecNom(i) = J & Format(DateAdd("d", 15 * (i * -1), FechaU), "mmyyyy")
        End If
        
        J = IIf(J = 1, 2, 1)
    Next
    
    
    
End If
End Sub

Private Function valida_datos() As Boolean

Dim rstlocal As ADODB.Recordset
Dim strMsg As String

If IDNomina = 0 Then
    MsgBox ("Debe seleccionar una nómina de la lista.")
    Exit Function
End If
Set rstlocal = New ADODB.Recordset

rstlocal.Open "SELECT IDNomina FROM Nom_Inf where IDNomina=" & IDNomina, _
cnnConexion, adOpenStatic, adLockReadOnly, adCmdText

If rstlocal.EOF And rstlocal.BOF Then
    MsgBox "Seleccione un elemento de la lista. " & vbCrLf & _
    "Si ya tiene un elemento selecconado, seleccionelo nuevamente."
    Exit Function
End If

rstlocal.Close
Set rstlocal = Nothing

strMsg = "Recuerde que este proceso efectuará " & vbCrLf & _
"cambios en la nómina y en la facuración." & vbCrLf & vbCrLf _
& "¿Esta seguro de reversar la nónima de " & vbCrLf & cmb.Text & "?"

valida_datos = Respuesta(strMsg)

End Function


