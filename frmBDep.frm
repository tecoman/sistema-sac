VERSION 5.00
Begin VB.Form frmBDep 
   Caption         =   "Busqueda Depósitos..."
   ClientHeight    =   1725
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   3150
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   1725
   ScaleWidth      =   3150
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmd 
      Caption         =   "Cerrar"
      Height          =   405
      Index           =   1
      Left            =   1920
      TabIndex        =   4
      Top             =   1230
      Width           =   1020
   End
   Begin VB.TextBox txt 
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1680
      TabIndex        =   3
      Top             =   315
      Width           =   1230
   End
   Begin VB.PictureBox pic 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00800000&
      Height          =   270
      Left            =   255
      ScaleHeight     =   210
      ScaleWidth      =   2640
      TabIndex        =   2
      Top             =   825
      Visible         =   0   'False
      Width           =   2700
   End
   Begin VB.CommandButton cmd 
      Caption         =   "Buscar"
      Height          =   405
      Index           =   0
      Left            =   735
      TabIndex        =   1
      Top             =   1230
      Width           =   1020
   End
   Begin VB.Label lbl 
      Alignment       =   2  'Center
      Caption         =   "Nº Depósito:"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   0
      Left            =   210
      TabIndex        =   0
      Top             =   330
      Width           =   1215
   End
End
Attribute VB_Name = "frmBDep"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Cmd_Click(Index As Integer)
If Index = 0 Then
    Call busca_deposito
Else
    Unload Me
    Set frmBDep = Nothing
End If
End Sub

Private Sub txt_KeyPress(KeyAscii As Integer)
Call Validacion(KeyAscii, "0123456789")
End Sub


Sub busca_deposito()
Dim rstlocal As ADODB.Recordset
Dim strSQL As String
'
If IsNumeric(txt) Then
    Set rstlocal = New ADODB.Recordset
    With rstlocal
        .CursorLocation = adUseClient
        strSQL = "SELECT * FROM TDFCheques WHERE (FPAgo='DEPOSITO' or FPago='TRANSFERENCIA') " & _
        " and NDoc Like '%" & txt & "'"
        .Open strSQL, cnnConexion, adOpenStatic, adLockReadOnly, adCmdText
        .Sort = "FechaMov ASC"
        If Not (.EOF And .BOF) Then
            .MoveFirst
            Do
                If Not .EOF Then
                    MsgBox !FPago & " Nº " & !NDoc & vbCrLf & "Descargado el " & !FechaMov & _
                    vbCrLf & "Inm: " & !CodInmueble & " Apto: " & Mid(!IDRecibo, 3, 4)
                End If
                .MoveNext
            Loop Until .EOF
        Else
            MsgBox "Depósito Nº " & txt & vbCrLf & "--NO REGISTRADO--"
            
        End If
        '
        .Close
    End With
    '
    Set rstlocal = Nothing
End If
'
End Sub
