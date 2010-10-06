VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmBitacora 
   Caption         =   "Bitácora del Sistema"
   ClientHeight    =   60
   ClientLeft      =   75
   ClientTop       =   420
   ClientWidth     =   2685
   ControlBox      =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   60
   ScaleWidth      =   2685
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdBitacora 
      Caption         =   "Imprimir"
      Height          =   765
      Index           =   1
      Left            =   2520
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   3300
      Width           =   1065
   End
   Begin VB.CommandButton cmdBitacora 
      Caption         =   "Salir"
      Height          =   765
      Index           =   0
      Left            =   3735
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   3300
      Width           =   1065
   End
   Begin VB.Frame fraBitacora 
      Caption         =   "B&uscar:"
      Height          =   855
      Left            =   225
      TabIndex        =   8
      Top             =   3210
      Width           =   4560
      Begin VB.TextBox txt 
         Height          =   300
         Left            =   1965
         TabIndex        =   9
         Top             =   465
         Width           =   2205
      End
      Begin VB.Label lblBitacora 
         BackStyle       =   0  'Transparent
         Caption         =   "Introduzca el texto que desea buscar y presione la tecla <enter>:"
         Height          =   465
         Index           =   3
         Left            =   750
         TabIndex        =   10
         Top             =   225
         Width           =   3450
      End
   End
   Begin MSComDlg.CommonDialog ctldialog 
      Left            =   240
      Top             =   3495
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin RichTextLib.RichTextBox rtxBitacora 
      Height          =   1665
      Left            =   200
      TabIndex        =   5
      Top             =   1470
      Width           =   4620
      _ExtentX        =   8149
      _ExtentY        =   2937
      _Version        =   393217
      BorderStyle     =   0
      HideSelection   =   0   'False
      ReadOnly        =   -1  'True
      ScrollBars      =   3
      BulletIndent    =   10
      TextRTF         =   $"frmBitacora.frx":0000
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComCtl2.DTPicker dtpBitacora 
      Height          =   315
      Left            =   4515
      TabIndex        =   4
      Top             =   945
      Width           =   1290
      _ExtentX        =   2275
      _ExtentY        =   556
      _Version        =   393216
      CalendarTitleBackColor=   -2147483646
      CalendarTitleForeColor=   -2147483639
      Format          =   21233665
      CurrentDate     =   37608
   End
   Begin VB.ComboBox cmbBitacora 
      Height          =   315
      Left            =   1050
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   945
      Width           =   2295
   End
   Begin VB.Label lblBitacora 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmBitacora.frx":0077
      Height          =   705
      Index           =   2
      Left            =   195
      TabIndex        =   3
      Top             =   210
      Width           =   4680
   End
   Begin VB.Label lblBitacora 
      BackStyle       =   0  'Transparent
      Caption         =   "Fecha:"
      Height          =   285
      Index           =   1
      Left            =   3840
      TabIndex        =   2
      Top             =   960
      Width           =   735
   End
   Begin VB.Label lblBitacora 
      BackStyle       =   0  'Transparent
      Caption         =   "Usuario:"
      Height          =   285
      Index           =   0
      Left            =   195
      TabIndex        =   1
      Top             =   960
      Width           =   765
   End
End
Attribute VB_Name = "frmBitacora"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
    
    'Módulo de Utiloidades del sistema. Registro diario de transacciones de usuarios
    '18/12/2002
    Dim strUbica$  'Variable global a nivel de módulo
    

    Private Sub cmbBitacora_Click()
    Me.MousePointer = vbHourglass
    If cmbBitacora.Text = "TODOS" Then  'ver los registros de todos los usuarios
        Call Ver_Reg(cmbBitacora.Text, Format(dtpBitacora, "ddmyy") & ".txt")
    Else    'ver registros usuario determinado
        Call Ver_Reg(cmbBitacora.Text, Format(dtpBitacora, "ddmyy") & ".txt")
    End If
    Me.MousePointer = vbDefault
    End Sub


    Private Sub cmdBitacora_Click(Index As Integer)
    '
    Dim errLocal#, strT$
    
    Select Case Index
        Case 0  'Salir
        '--------------------
            Unload Me: Set frmBitacora = Nothing
        Case 1  'Imprimir   'Imprime el resultado del rtxBitacora
        '--------------------
            MousePointer = vbHourglass
            Printer.PaintPicture LoadPicture("F:\sac\iconos\logo.bmp"), 200, 200, 4290, 1365, _
            , , , , vbSrcCopy
            Printer.CurrentY = 300
            strT = "SAC - BITACORA"
            FontSize = 24: FontName = "Times New Roman"
            Printer.FontSize = FontSize: Printer.FontName = FontName
            Printer.CurrentX = Printer.ScaleWidth - TextWidth(strT) - 50
            Printer.Print strT
            FontSize = 14: Printer.FontSize = FontSize
            strT = "Usuario: " & cmbBitacora & Space(3) & "Fecha: " & dtpBitacora
            Printer.CurrentX = Printer.ScaleWidth - TextWidth(strT) - 150
            Printer.Print strT
            FontSize = 8: Printer.FontSize = FontSize
            strT = "Impreso por: " & gcUsuario & Space(3) & "Fecha Impresión: " & Date
            Printer.CurrentX = Printer.ScaleWidth - TextWidth(strT) + 150
            Printer.Print strT
            Printer.FontSize = 8: Printer.FontName = "arial"
            Printer.CurrentY = 1600
            Printer.Print rtxBitacora.Text
            Printer.EndDoc
            MousePointer = vbDefault
    End Select
    '
    End Sub

    Private Sub dtpBitacora_Change()
    Me.MousePointer = vbHourglass
    dtpBitacora.Refresh
    If dtpBitacora.Value <= Date Then
        Call Ver_Reg(cmbBitacora.Text, Format(dtpBitacora, "ddmyy") & ".txt")
    Else
        MsgBox "La fecha seleccionada es superior a la fecha actual..", vbInformation, _
        App.ProductName
    End If
    Me.MousePointer = vbDefault
    End Sub

    Private Sub Form_Load() '
    '----------------------------------------------------------------
    Dim cnnBitacora As New ADODB.Connection   'Variables locales
    Dim rstBitacora As New ADODB.Recordset
    Dim strSql$
    '
    strUbica = gcPath & "\Bitacora\"    'ubicacion de la carpeta
    cnnBitacora.Open cnnOLEDB & gcPath & "\Tablas.mdb"
    'If gcNivel >= 1 Then    'si
      strSql = "SELECT * FROM Usuarios WHERE Nivel>=" & gcNivel
    'Else
     '   strSql = "SELECT * FROM Usuarios;"
    'End If
    cmdBitacora(0).Picture = LoadResPicture("Salir", vbResIcon)
    cmdBitacora(1).Picture = LoadResPicture("Print", vbResIcon)
    '
    rstBitacora.CursorLocation = adUseClient
    rstBitacora.Open strSql, cnnBitacora, adOpenStatic, adLockReadOnly, adCmdText
    rstBitacora.Sort = "NombreUsuario"
    '
    If Not rstBitacora.EOF Or rstBitacora.BOF Then
        
        rstBitacora.MoveFirst
        Do
            cmbBitacora.AddItem (rstBitacora!NombreUsuario)
            rstBitacora.MoveNext
        Loop Until rstBitacora.EOF
        dtpBitacora.Value = Date
        cmbBitacora.AddItem ("TODOS")
        'cmbBitacora.Text = "TODOS"
    End If
    Set rstBitacora = Nothing
    Set cnnBitacora = Nothing
    '
    End Sub

    Private Sub Form_Resize()   '
    '----------------------------------------------------------------
    On Error Resume Next
    If WindowState <> vbMinimized Then
        lblBitacora(2).Width = ScaleWidth - lblBitacora(2).Left
        cmdBitacora(0).Top = ScaleHeight - (cmbBitacora.Height * 2) - 200
        cmdBitacora(1).Top = ScaleHeight - (cmbBitacora.Height * 2) - 200
        cmdBitacora(1).Left = (ScaleWidth - rtxBitacora.Width - (rtxBitacora.Left * 2)) + rtxBitacora.Width - (cmdBitacora(1).Width * 2)
        cmdBitacora(0).Left = cmdBitacora(1).Left + cmdBitacora(1).Width + 100
        With rtxBitacora
            .Width = ScaleWidth - (.Left * 2)
            .Height = ScaleHeight - (.Top + 200) - cmdBitacora(0).Height
            fraBitacora.Left = .Left
        End With
        fraBitacora.Top = cmdBitacora(0).Top - 90
    End If
    
    End Sub


    '---------------------------------------------------------------------------------------------
    '   Rutina:     Ver_Reg
    '
    '   Rutina que muestra la información correspondiente a un usuario específico,
    '   todas las transacciones efectuadas en una fecha determinada.
    '---------------------------------------------------------------------------------------------
    Public Sub Ver_Reg(strUser As String, strArchivo As String)
    'Variables locales
    Dim numFichero%
    Dim Hora, Maquina As String * 10, Usuario As String, Accion$
    '
    On Error Resume Next
    'Maquina = String(10, "_")
    'Usuario = String(10, "_")
    rtxBitacora = ""
    numFichero = FreeFile
    Open strUbica & strArchivo For Input As numFichero
    'rtxBitacora.Visible = False
    rtxBitacora.Text = "Buscando Transacciones..." & vbCrLf
    Do
        Input #numFichero, Hora, Maquina, Usuario, Accion
        If Not strUser = "TODOS" Then
            If Trim(Usuario) = strUser Then
                If txt = "" Then
                    rtxBitacora.Text = rtxBitacora.Text & Hora & vbTab & UCase(Maquina) & _
                    Space(3) & Usuario & Space(3) & Accion & vbCrLf
                Else
                    If UCase(Accion) Like "*" & UCase(txt) & "*" Or Maquina = txt Or Usuario = txt Then
                        rtxBitacora.Text = rtxBitacora.Text & Hora & vbTab & UCase(Maquina) & _
                        Space(3) & Usuario & Space(3) & Accion & vbCrLf
                    End If
                End If
            End If
        Else
            If txt = "" Then
                rtxBitacora.Text = rtxBitacora.Text & Hora & vbTab & UCase(Maquina) & Space(3) & _
                Usuario & Space(3) & Accion & vbCrLf
            Else
                If UCase(Accion) Like "*" & UCase(txt) & "*" Or UCase(Trim(Maquina)) = txt Or Trim(Usuario) = txt Then
                    rtxBitacora.Text = rtxBitacora.Text & Hora & vbTab & UCase(Maquina) & _
                    Space(3) & Usuario & Space(3) & Accion & vbCrLf
                End If
            End If
        End If
        'rtxBitacora.Refresh
    Loop Until EOF(numFichero)
    rtxBitacora.Text = rtxBitacora.Text & "***FIN ARCHIVO***"
    rtxBitacora.SetFocus
    Close numFichero
    '
    End Sub

    


    Private Sub rtxBitacora_Change()
    rtxBitacora.SelStart = Len(rtxBitacora)
    End Sub

    Private Sub rtxBitacora_MouseDown(Button%, Shift%, X!, Y!)
    If Button = 2 Then PopupMenu FrmAdmin.mnuBit
    End Sub

    Private Sub txt_GotFocus()
    txt.SelStart = 0
    txt.SelLength = Len(txt)
    End Sub

    '
    Private Sub txt_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    'variables locales
    Dim strMsg As String
    Static FoundPos As Long
    Static Coin As Integer
    '
    If KeyAscii = 13 Then
       ' Busca el texto especificado en el control TextBox.
       FoundPos = rtxBitacora.Find(txt, FoundPos, , tfWholeWord)
       ' Muestra un mensaje si ha encontrado el texto.
       If FoundPos <> -1 Then
            Coin = Coin + 1
       Else
            If Coin = 0 Then
                strMsg = "No existen coincidencias...."
            Else
                strMsg = "No existen más coincidencias..."
                Coin = 0
            End If
            MsgBox strMsg, vbInformation, App.ProductName
       End If
       FoundPos = FoundPos + 1
       
    End If
    '
    End Sub

    
