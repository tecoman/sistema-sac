VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmEmail 
   Caption         =   "Aviso de Cobro [Cuerpo del Mensaje]"
   ClientHeight    =   4500
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   6465
   Icon            =   "frmEmail.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4500
   ScaleWidth      =   6465
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdEmail 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      Height          =   735
      Index           =   1
      Left            =   5235
      Picture         =   "frmEmail.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3690
      Width           =   1215
   End
   Begin VB.CommandButton cmdEmail 
      Caption         =   "&Guardar"
      Height          =   735
      Index           =   0
      Left            =   4020
      Picture         =   "frmEmail.frx":074C
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   3690
      Width           =   1215
   End
   Begin RichTextLib.RichTextBox rtxEmail 
      Height          =   3600
      Left            =   60
      TabIndex        =   0
      Top             =   0
      Width           =   6465
      _ExtentX        =   11404
      _ExtentY        =   6350
      _Version        =   393217
      FileName        =   "F:\SAC\DATOS\email.txt"
      TextRTF         =   $"frmEmail.frx":0B8E
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "frmEmail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdEmail_Click(Index As Integer)
Select Case Index
    'Guardar
    Case 0
        rtxEmail.Text = rtxEmail.Text & vbCrLf & vbCrLf & Space(2)
        rtxEmail.SaveFile gcPath & "\email.txt", rtfText
    'salir
    Case 1: Unload Me
    '
End Select
End Sub

Private Sub Form_Load()
'variable locales
Dim numFichero As Integer
Dim strArchivo As String
Dim blnNuevo As Boolean
'
strArchivo = gcPath & "\email.txt"
'si no existe el archivo lo crea
If Dir(strArchivo) = "" Then 'si no existe lo crea
    
    numFichero = FreeFile
    Open strArchivo For Append As numFichero
    Write #numFichero, "----Escriba aquí el el texto del cuerpo del mensaje----"
    Close numFichero
    blnNuevo = True
End If
CenterForm frmEmail
rtxEmail.FileName = strArchivo
If blnNuevo Then
    rtxEmail.SelStart = 0
    rtxEmail.SelLength = Len(rtxEmail.Text)
End If
'

End Sub
