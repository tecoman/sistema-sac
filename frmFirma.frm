VERSION 5.00
Begin VB.Form frmFirma 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmd 
      Caption         =   "&Salir"
      Height          =   915
      Left            =   9510
      Picture         =   "frmFirma.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   6765
      Width           =   1215
   End
   Begin VB.Label lbl 
      Caption         =   "Firma Junta de Condominio 2503"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   315
      TabIndex        =   1
      Top             =   360
      Width           =   5730
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      Height          =   3615
      Left            =   315
      Picture         =   "frmFirma.frx":030A
      Stretch         =   -1  'True
      Top             =   1365
      Width           =   10455
   End
End
Attribute VB_Name = "frmFirma"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public strFichero As String
'
Private Sub cmd_Click()
Unload Me
Set frmFirma = Nothing
End Sub

Private Sub Form_Load()
'variables locales
lbl = "Firma Junta de Condominio " & gcCodInm
Me.Caption = lbl
Image1.Picture = LoadPicture(strFichero)
'
End Sub

Private Sub Form_Resize()
cmd.Top = Me.ScaleHeight - cmd.Height - 200
cmd.Left = Me.ScaleWidth - cmd.Width - 200
Image1.Left = 200
If LoadPicture(strFichero).Height > 5000 Then
    Image1.Top = lbl.Top + lbl.Height + 300
    Image1.Height = Me.ScaleHeight - Image1.Top - cmd.Height - 200
End If
End Sub
    
