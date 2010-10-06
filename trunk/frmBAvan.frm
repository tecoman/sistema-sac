VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form frmBAvan 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Busqueda...."
   ClientHeight    =   6150
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5760
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6150
   ScaleWidth      =   5760
   Begin VB.PictureBox pic 
      AutoRedraw      =   -1  'True
      Height          =   465
      Left            =   210
      ScaleHeight     =   405
      ScaleWidth      =   2565
      TabIndex        =   12
      Top             =   5460
      Visible         =   0   'False
      Width           =   2625
   End
   Begin VB.TextBox txt 
      Height          =   315
      Index           =   3
      Left            =   2295
      TabIndex        =   11
      Tag             =   "email"
      Top             =   2295
      Width           =   3000
   End
   Begin VB.TextBox txt 
      Height          =   315
      Index           =   2
      Left            =   2295
      TabIndex        =   10
      Tag             =   "Cedula"
      Top             =   1845
      Width           =   3000
   End
   Begin VB.TextBox txt 
      Height          =   315
      Index           =   1
      Left            =   2295
      TabIndex        =   9
      Tag             =   "Nombre"
      Top             =   1365
      Width           =   3000
   End
   Begin VB.TextBox txt 
      Height          =   315
      Index           =   0
      Left            =   2295
      TabIndex        =   8
      Tag             =   "Codigo"
      Top             =   885
      Width           =   3000
   End
   Begin VB.CommandButton cmd 
      Caption         =   "&Salir"
      Height          =   495
      Index           =   1
      Left            =   4290
      TabIndex        =   7
      Top             =   5430
      Width           =   1215
   End
   Begin VB.CommandButton cmd 
      Caption         =   "&Buscar"
      Height          =   495
      Index           =   0
      Left            =   2955
      TabIndex        =   6
      Tag             =   "0"
      Top             =   5445
      Width           =   1215
   End
   Begin MSFlexGridLib.MSFlexGrid Grid 
      Height          =   2430
      Left            =   210
      TabIndex        =   5
      Tag             =   "1000|1000|3000"
      Top             =   2760
      Width           =   5325
      _ExtentX        =   9393
      _ExtentY        =   4286
      _Version        =   393216
      Cols            =   3
      FixedCols       =   0
      BackColorFixed  =   -2147483646
      ForeColorFixed  =   -2147483639
      BackColorBkg    =   -2147483636
      FocusRect       =   2
      HighLight       =   2
      MergeCells      =   2
      BorderStyle     =   0
      Appearance      =   0
      FormatString    =   "^Cod.Inm|^Cod.Pro.|Nombre/Razón Social"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Lbl 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmBAvan.frx":0000
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   660
      Index           =   4
      Left            =   225
      TabIndex        =   4
      Top             =   75
      Width           =   5340
   End
   Begin VB.Label Lbl 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "e-mail:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   3
      Left            =   585
      TabIndex        =   3
      Top             =   2295
      Width           =   1470
   End
   Begin VB.Label Lbl 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Cédula / RIF.:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   2
      Left            =   585
      TabIndex        =   2
      Top             =   1845
      Width           =   1470
   End
   Begin VB.Label Lbl 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Nombre / Razón social:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   1
      Left            =   -405
      TabIndex        =   1
      Top             =   1395
      Width           =   2460
   End
   Begin VB.Label Lbl 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Cod.Apto:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   0
      Left            =   585
      TabIndex        =   0
      Top             =   945
      Width           =   1470
   End
End
Attribute VB_Name = "frmBAvan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Cmd_Click(index As Integer)
Select Case index
    Case 0  'buscar - cancelar
        cmd(0).Caption = IIf(cmd(0).Tag = 0, "&Cancelar", "&Buscar")
        cmd(0).Tag = IIf(cmd(0).Tag = 0, 1, 0)
        cmd(1).Enabled = Not cmd(1).Enabled
        If cmd(0).Tag = 1 Then BusquedaAvanzada
        
    Case 1  'cerrar ventana
        Unload Me
        Set frmBAvan = Nothing
End Select

End Sub

Private Sub Form_Load()
CenterForm frmBAvan
Call centra_titulo(Grid, True)

End Sub

Private Sub BusquedaAvanzada()
'variables locales
Dim rstInm As ADODB.Recordset
Dim rstPro As ADODB.Recordset
Dim strConex As String
Dim strCriterio As String
Dim i&
Static J&
'
For i = 0 To 3
    If Txt(i) <> "" Then
        strCriterio = IIf(strCriterio <> "", " AND ", "") + Txt(i).Tag + _
        " LIKE '*" + Txt(i) & "*'"
    End If
Next
If strCriterio = "" Then
    MsgBox "Debe introducir un criterio para la búsqueda", vbInformation, App.ProductName
    cmd(1).Enabled = Not cmd(1).Enabled
    cmd(0).Caption = "&Buscar"
    Exit Sub
End If
'
Set rstInm = New ADODB.Recordset
rstInm.CursorLocation = adUseClient
rstInm.Open "Inmueble", cnnConexion, adOpenKeyset, adLockOptimistic, adCmdTable
pic.Visible = True
'
If Not (rstInm.EOF And rstInm.BOF) Then
    rstInm.MoveFirst
    Set rstPro = New ADODB.Recordset
    J = 1
    rtnLimpiar_Grid Grid
    Do
        DoEvents
        If cmd(0).Tag = 0 Then Exit Sub
        Call UpdateStatus(pic, CLng(rstInm.AbsolutePosition * 100 / rstInm.RecordCount) + 1, 2)
        strConex = cnnOLEDB + gcPath + rstInm("Ubica") + "inm.mdb"
        rstPro.Open "Propietarios", strConex, adOpenKeyset, adLockOptimistic, adCmdTable
        rstPro.Filter = strCriterio
        '
        If Not (rstPro.EOF And rstPro.BOF) Then
            rstPro.MoveFirst
            Do
                DoEvents
                If cmd(0).Tag = 0 Then
                    cmd(0).Caption = IIf(cmd(0).Tag = 0, "&Cancelar", "&Buscar")
                    cmd(0).Tag = IIf(cmd(0).Tag = 0, 1, 0)

                    Exit Sub
                End If
                If J > 1 Then Grid.AddItem ""
                '--aqui muestra las coincidencias en el grid--'
                Grid.TextMatrix(J, 0) = rstInm("CodInm")
                Grid.TextMatrix(J, 1) = rstPro("Codigo")
                Grid.TextMatrix(J, 2) = rstPro("Nombre")
                rstPro.MoveNext
                J = J + 1
            Loop Until rstPro.EOF
        End If
        rstPro.Close
        '
        rstInm.MoveNext
    Loop Until rstInm.EOF
    
End If
pic.Visible = False
rstInm.Close
cmd(0).Caption = IIf(cmd(0).Tag = 0, "&Cancelar", "&Buscar")
cmd(0).Tag = IIf(cmd(0).Tag = 0, 1, 0)
cmd(1).Enabled = True
Set rstInm = Nothing
Set rstPro = Nothing
End Sub

Private Sub txt_KeyPress(index As Integer, KeyAscii As Integer)
If index < 3 Then KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

    Private Sub UpdateStatus(pic As PictureBox, ByVal sngPercent As Single, _
    Optional ByVal fBorderCase)
    Dim strPercent As String
    Dim intX As Integer
    Dim intY As Integer
    Dim intWidth As Integer
    Dim intHeight As Integer
    
    If sngPercent > 100 Then sngPercent = 100
    If IsMissing(fBorderCase) Then fBorderCase = False
    
    Const colBackground = &HFFFFFF ' white
    Const colForeground = &HFF00&     'verde manzana

    pic.ForeColor = vbBlack
    pic.BackColor = colBackground
    
    '
    Dim intPercent
    intPercent = sngPercent
    
    If intPercent = 0 Then
        If Not fBorderCase Then
            intPercent = 1
        End If
    ElseIf intPercent = 100 Then
        If Not fBorderCase Then
            intPercent = 99
        End If
    End If
    
    strPercent = Format$(intPercent) & "%"
    intWidth = pic.TextWidth(strPercent)
    intHeight = pic.TextHeight(strPercent)

    intX = pic.Width / 2 - intWidth / 2
    intY = (pic.Height / 2) - (intHeight / 2) - 35

    If sngPercent > 0 Then
        pic.Line (0, 0)-(sngPercent * pic.Width / 100, pic.Height), colForeground, BF
    Else
        pic.Line (0, 0)-(pic.Width, pic.Height), colForeground, BF
    End If
    pic.DrawMode = 13 ' Copy Pen
    '
    pic.CurrentX = intX
    pic.CurrentY = intY
    pic.Print strPercent

    pic.Refresh
End Sub

