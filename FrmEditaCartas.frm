VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form FrmEditaCartas 
   Caption         =   "Edición de Cartas y telegramas"
   ClientHeight    =   4020
   ClientLeft      =   60
   ClientTop       =   420
   ClientWidth     =   7245
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   4020
   ScaleWidth      =   7245
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command1 
      Caption         =   "&Salir"
      Height          =   615
      Index           =   1
      Left            =   4575
      TabIndex        =   2
      Top             =   3300
      Width           =   2535
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Guardar"
      Height          =   615
      Index           =   0
      Left            =   2040
      TabIndex        =   1
      Top             =   3300
      Width           =   2535
   End
   Begin RichTextLib.RichTextBox RichTextBox1 
      Height          =   2985
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   7005
      _ExtentX        =   12356
      _ExtentY        =   5265
      _Version        =   393217
      BorderStyle     =   0
      Enabled         =   -1  'True
      ScrollBars      =   3
      FileName        =   "F:\sac\Docs\Telegrama.txt"
      TextRTF         =   $"FrmEditaCartas.frx":0000
   End
End
Attribute VB_Name = "FrmEditaCartas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
    Dim strRutaCarta As String
    
    Private Sub Command1_Click(Index As Integer)
    '
    Select Case Index
    '
        Case 0  'Guardar
        '
            Open strRutaCarta For Output As 1
            Print #1, RichTextBox1.Text
            Close 1
            MsgBox "Archivo Actualizado......"
            
        Case 1 'Salir
        '
            Unload Me
            Set freditacartas = Nothing
    End Select
    '
    End Sub

    Private Sub Form_Load()
    RichTextBox1.FileName = "\\" & DataServer & "\APLI\sac\Docs\Telegrama.txt"
    strRutaCarta = RichTextBox1.FileName
    End Sub


    Private Sub Form_Resize()
    '
    If WindowState <> vbMinimized Then
        RichTextBox1.Width = ScaleWidth - RichTextBox1.Left - 100
        RichTextBox1.Height = ScaleHeight - (Command1(0).Width + ScaleTop + 100)
        Command1(0).Top = RichTextBox1.Height + RichTextBox1.Top + 200
        Command1(1).Top = RichTextBox1.Height + RichTextBox1.Top + 200
        Command1(0).Left = ScaleWidth - (Command1(0).Width + Command1(1).Width + 100)
        Command1(1).Left = Command1(0).Left + Command1(0).Width
    End If
    '
    End Sub


    Private Sub RichTextBox1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 And UCase(RichTextBox1.Text) Like "OBTENER*" Then
        KeyAscii = 0
        RichTextBox1.FileName = Left(gcPath, 7) & "docs\TelegramaCopy.txt"
        RichTextBox1.Refresh
    End If
    End Sub

