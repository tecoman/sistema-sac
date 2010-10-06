VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmSupervision 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Supervisión"
   ClientHeight    =   2760
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5595
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2760
   ScaleWidth      =   5595
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmd 
      Caption         =   "&Salir"
      Height          =   315
      Index           =   2
      Left            =   4245
      TabIndex        =   5
      Top             =   2085
      Width           =   975
   End
   Begin VB.CommandButton cmd 
      Caption         =   "&Publicar"
      Height          =   315
      Index           =   1
      Left            =   3210
      TabIndex        =   4
      Top             =   2085
      Width           =   975
   End
   Begin VB.CommandButton cmd 
      Caption         =   "..."
      Height          =   315
      Index           =   0
      Left            =   4815
      TabIndex        =   3
      Top             =   1455
      Width           =   390
   End
   Begin VB.TextBox txt 
      Height          =   315
      Left            =   570
      TabIndex        =   1
      Top             =   1455
      Width           =   4155
   End
   Begin MSComDlg.CommonDialog dlg 
      Left            =   360
      Top             =   2070
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label lbl 
      Caption         =   "Seleccione el informe de supervisión que desea publicar en www.administradorasac.com y luego presione el botón publicar:"
      Height          =   570
      Index           =   1
      Left            =   630
      TabIndex        =   2
      Top             =   690
      Width           =   4005
   End
   Begin VB.Label lbl 
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   0
      Left            =   270
      TabIndex        =   0
      Top             =   255
      Width           =   4830
   End
End
Attribute VB_Name = "frmSupervision"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Const server  As String = "administradorasac.com"
Const user As String = "admras"
Const pass As String = "dmn+str"


Private Sub cmd_Click(Index As Integer)
Select Case Index
    Case 0  ' cargar documento word
        dlg.CancelError = True
        dlg.Filter = "Supervisiones (*.pdf)|*.pdf"
        dlg.FilterIndex = 1
        dlg.DialogTitle = "Abrir supervisión"
        dlg.ShowOpen
    
        If Err.Number = cdlCancel Then Exit Sub
        txt.Text = dlg.FileName
    Case 2 '
        Unload Me
        Set frmSupervision = Nothing
    Case 1  ' publicar supervision
        publicar_supervision (txt.Text)
        
End Select

End Sub

Private Sub Form_Load()
lbl(0) = gcCodInm & " " & gcNomInm
End Sub

Private Sub publicar_supervision(sOrigen As String)
Dim ftp As cFtp
Set ftp = New cFtp

If ftp.OpenConnection(server, user, pass) Then
    ftp.SetFTPDirectory "WWW/Supervisiones"
    If Not ftp.SimpleFTPPutFile(sOrigen, gcCodInm & ".pdf") Then
        MsgBox "No se pudo publicar el archivo en el servidor." & vbCrLf & "Inténtelo nuevamente", _
        vbCritical, App.ProductName
        Exit Sub
    End If
    ftp.CloseConnection
    Set ftp = Nothing
    MsgBox "Archivo " & sOrigen & vbCrLf & "publicado con éxito", vbInformation, App.ProductName
Else
    MsgBox "Imposible establecer comunicación con el servidor." & vbCrLf & "No se puedo publicar: " & _
    sOrigen, vbCritical, App.ProductName
End If
End Sub
