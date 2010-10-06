VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmSMS 
   Caption         =   "Enviar Correo Electrónico"
   ClientHeight    =   9900
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15240
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9900
   ScaleWidth      =   15240
   WindowState     =   2  'Maximized
   Begin VB.Frame fra 
      Caption         =   "Seleccione las opciones del envío:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   9525
      Index           =   0
      Left            =   300
      TabIndex        =   0
      Top             =   330
      Width           =   14595
      Begin VB.CheckBox chk 
         Alignment       =   1  'Right Justify
         Caption         =   "Todos los Inmuebles"
         Height          =   405
         Left            =   855
         TabIndex        =   21
         Top             =   1200
         Width           =   2100
      End
      Begin VB.ListBox lst 
         Height          =   1425
         Left            =   12735
         TabIndex        =   19
         Top             =   1995
         Visible         =   0   'False
         Width           =   1440
      End
      Begin VB.TextBox txt 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Index           =   0
         Left            =   4170
         MaxLength       =   140
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   15
         Top             =   2265
         Width           =   6795
      End
      Begin VB.CommandButton Cmd 
         Cancel          =   -1  'True
         Caption         =   "&Salir"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Index           =   4
         Left            =   12810
         Picture         =   "frmSMS.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   600
         Width           =   1380
      End
      Begin VB.CommandButton Cmd 
         Caption         =   "&Ver SMS"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Index           =   2
         Left            =   9690
         Picture         =   "frmSMS.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   600
         Width           =   1380
      End
      Begin VB.CommandButton Cmd 
         Caption         =   "&Enviar SMS"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Index           =   1
         Left            =   11250
         Picture         =   "frmSMS.frx":0552
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   600
         Width           =   1380
      End
      Begin VB.Frame fra 
         Caption         =   "Visor de Mesajes:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   5205
         Index           =   2
         Left            =   360
         TabIndex        =   10
         Top             =   3825
         Width           =   13875
         Begin ComctlLib.ProgressBar pBar 
            Height          =   375
            Left            =   360
            TabIndex        =   14
            Top             =   4680
            Visible         =   0   'False
            Width           =   13155
            _ExtentX        =   23204
            _ExtentY        =   661
            _Version        =   327682
            Appearance      =   0
            Min             =   1e-4
         End
         Begin RichTextLib.RichTextBox rtb 
            Height          =   3990
            Left            =   360
            TabIndex        =   18
            Top             =   405
            Width           =   13125
            _ExtentX        =   23151
            _ExtentY        =   7038
            _Version        =   393217
            BorderStyle     =   0
            Enabled         =   -1  'True
            ScrollBars      =   2
            Appearance      =   0
            TextRTF         =   $"frmSMS.frx":0B68
         End
         Begin VB.Label lbl 
            Height          =   195
            Index           =   6
            Left            =   330
            TabIndex        =   22
            Top             =   4530
            Width           =   8925
         End
         Begin VB.Shape Shape1 
            BackColor       =   &H8000000C&
            BackStyle       =   1  'Opaque
            BorderColor     =   &H8000000C&
            Height          =   4020
            Index           =   2
            Left            =   345
            Top             =   390
            Width           =   13155
         End
      End
      Begin VB.Frame fra 
         Caption         =   "Tipo de SMS:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1710
         Index           =   1
         Left            =   360
         TabIndex        =   4
         Top             =   1740
         Width           =   2685
         Begin VB.OptionButton Opt 
            Caption         =   "&Personalizado"
            Height          =   285
            Index           =   2
            Left            =   420
            TabIndex        =   7
            Top             =   1215
            Width           =   1950
         End
         Begin VB.OptionButton Opt 
            Caption         =   "Última &Facturación"
            Height          =   285
            Index           =   1
            Left            =   420
            TabIndex        =   6
            Top             =   810
            Width           =   1950
         End
         Begin VB.OptionButton Opt 
            Caption         =   "&Deuda"
            Height          =   285
            Index           =   0
            Left            =   420
            TabIndex        =   5
            Top             =   405
            Value           =   -1  'True
            Width           =   1950
         End
      End
      Begin MSDataListLib.DataCombo dtc 
         Height          =   315
         Index           =   0
         Left            =   1455
         TabIndex        =   8
         Top             =   750
         Width           =   1500
         _ExtentX        =   2646
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo dtc 
         Height          =   315
         Index           =   1
         Left            =   3930
         TabIndex        =   9
         Top             =   750
         Width           =   5070
         _ExtentX        =   8943
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
      End
      Begin VB.Label lbl 
         Caption         =   "Variables:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   165
         Index           =   5
         Left            =   12720
         TabIndex        =   20
         Top             =   1770
         Visible         =   0   'False
         Width           =   3195
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "(0 / 140)"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   225
         Index           =   4
         Left            =   4110
         TabIndex        =   17
         Top             =   2955
         Width           =   525
      End
      Begin VB.Image Image1 
         Height          =   1740
         Left            =   3300
         Picture         =   "frmSMS.frx":0BEA
         Stretch         =   -1  'True
         Top             =   1815
         Width           =   8595
      End
      Begin VB.Label lbl 
         Caption         =   "Mensaje: (Máx 140 Caracteres)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   165
         Index           =   3
         Left            =   3465
         TabIndex        =   16
         Top             =   1620
         Width           =   3195
      End
      Begin VB.Label lbl 
         Caption         =   "Nombre:"
         Height          =   285
         Index           =   1
         Left            =   3270
         TabIndex        =   3
         Top             =   780
         Width           =   795
      End
      Begin VB.Label lbl 
         Caption         =   "Codigo:"
         Height          =   285
         Index           =   0
         Left            =   825
         TabIndex        =   2
         Top             =   780
         Width           =   705
      End
      Begin VB.Label lbl 
         Caption         =   "Información sobre el inmueble:"
         Height          =   285
         Index           =   2
         Left            =   300
         TabIndex        =   1
         Top             =   360
         Width           =   2670
      End
   End
End
Attribute VB_Name = "frmSMS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim sRTF As String
Dim n As Long
Const movilnet = 0
Const movistar = 1
Const digitel = 2
Dim operadora(3) As Integer

Private Sub chk_Click()
If chk.Value = vbChecked Then
    dtc(0).Enabled = False
    dtc(1).Enabled = False
Else
    dtc(0).Enabled = True
    dtc(1).Enabled = True
End If
End Sub

Private Sub cmd_Click(Index As Integer)

Select Case Index
    Case 4  'salir
        Unload Me
        Set frmSMS = Nothing
    
    Case 1  ' enviar sms
        
    Case 2  'ver sms
        Cmd(4).Enabled = False
        Cmd(1).Enabled = False
        Cmd(2).Enabled = False
        rtb.Text = ""
        lbl(6) = ""
        operadora(movilnet) = 0
        operadora(movistar) = 0
        operadora(digitel) = 0
        sRTF = "{\rtf1\ansi "
        n = 0
        If chk.Value = vbChecked Then
            
            With FrmAdmin.objRst
                .Filter = "Codinm <> '7777' and codinm<>'8888' and codinm<>'9999' and Inactivo = false "
                pBar.Max = .RecordCount
                pBar.Visible = True
                If Not (.EOF And .BOF) Then
                    .MoveFirst
                    Do
                        Call ver_sms(!CodInm, !Nombre)
                        pBar.Value = .AbsolutePosition
                        DoEvents
                        .MoveNext
                    Loop Until .EOF
                End If
                .Filter = ""
                pBar.Visible = False
            End With
        Else
            If dtc(0) = "" And dtc(1) = "" Then Exit Sub
            If dtc(0) = "" Then dtc(0) = dtc(1).BoundText
            If dtc(1) = "" Then dtc(1) = dtc(0).BoundText

            Call ver_sms(dtc(0), dtc(1))
        End If
        sRTF = sRTF & "}"
        rtb.TextRTF = sRTF
        lbl(6) = "Mensajes Total: " & n & " - Movilnet: " & operadora(movilnet) & _
        " - Movistar: " & operadora(movistar) & " - Digitel: " & operadora(digitel)
        Cmd(4).Enabled = True
        Cmd(1).Enabled = True
        Cmd(2).Enabled = True
End Select

End Sub

Private Sub dtc_Click(Index As Integer, Area As Integer)
If Area = 2 Then
    dtc(IIf(Index = 0, 1, 0)) = dtc(Index).BoundText
    If Not dtc(IIf(Index = 0, 1, 0)).MatchedWithList Then
        dtc(IIf(Index = 0, 1, 0)) = ""
    End If
End If
End Sub

Private Sub dtc_KeyPress(Index As Integer, KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
If KeyAscii = 13 Then Call dtc_Click(Index, 2)
End Sub

Private Sub Form_Load()
Set dtc(0).RowSource = FrmAdmin.objRst
dtc(0).ListField = "CodInm"
dtc(0).BoundColumn = "Nombre"
Set dtc(1).RowSource = FrmAdmin.ObjRstNom
dtc(1).ListField = "Nombre"
dtc(1).BoundColumn = "CodInm"

End Sub

Private Sub Opt_Click(Index As Integer)
Select Case Index
    Case 0  'deuda
        txt(0).Enabled = False
        
    Case 1  'ultima facturacion
        txt(0).Enabled = False
        
    Case 2  'personalizado
        txt(0).Enabled = True
        
End Select
End Sub


Private Sub txt_Change(Index As Integer)
If Index = 0 Then
    lbl(4) = "(" & Len(txt(0)) & "/140)"
End If
End Sub

Private Function numero_celular(operadora As String) As Boolean
Dim compa() As String
compa = Split("0424,0426,0412,0414,0416", ",")
For I = 0 To UBound(compa)
    numero_celular = operadora = compa(I)
    If numero_celular Then Exit Function
Next
End Function


Private Sub ver_sms(cod_inm As String, Inmueble As String)
Dim rst As ADODB.Recordset
Dim cnn As ADODB.Connection
Dim sql As String, msg As String
Dim sCelular As String, sOperadora As String

Set cnn = New ADODB.Connection
Set rst = New ADODB.Recordset


cnn.Open cnnOLEDB & gcPath & "\" & cod_inm & "\inm.mdb"

Select Case True
    Case Opt(0)
        sql = "SELECT Codigo,Deuda, UltFact,celular FROM " & _
        "Propietarios where Celular <>'' and celular is not null ORDER BY Codigo"
    Case Opt(1)
        sql = "select factura.codprop, factura.facturado, factura.fechafactura, factura.periodo, " & _
        "propietarios.celular from factura inner join propietarios " & _
        "on propietarios.codigo = factura.codprop " & _
        "where Periodo in (select max(periodo) from factura) and celular<>'' and celular is not null ORDER BY Codigo"
    Case Opt(2)
        sql = Trim(txt(0))
End Select



rst.Open sql, cnn, adOpenKeyset, adLockOptimistic, adCmdText


If Not (rst.EOF And rst.BOF) Then
    
    rst.MoveFirst
    Do
        sCelular = rst("Celular")
        sCelular = Replace(sCelular, "(", "")
        sCelular = Replace(sCelular, ")", "")
        sCelular = Replace(sCelular, "-", "")
        rst.Update "celular", sCelular
        sOperadora = Left(sCelular, 4)
        If numero_celular(sOperadora) Then
            Select Case sOperadora
                Case "0426", "0416"
                    operadora(movilnet) = operadora(movilnet) + 1
                Case "0424", "0414"
                    operadora(movistar) = operadora(movistar) + 1
                Case "0412"
                    operadora(digitel) = operadora(digitel) + 1
            End Select
            If Opt(0) Then
                'SMS deuda
                msg = "Adm. Sac le informa que el apto. " & rst("Codigo") & _
                " de " & Inmueble & " presenta una deuda, a la fecha, de Bs." & _
                Format(rst("deuda"), "###.00")
                
            ElseIf Opt(1) Then
                'SMS ultima facturacion
                msg = "En fecha " & Format(rst("FechaFactura"), "dd/mm/yyyy") & _
                    ", se ha facturado el perído " & UCase(Format(rst("periodo"), "mmm-yyyy")) & _
                    ", correspondiente al Apto: " & rst("codprop") & " de " & _
                    Inmueble
            Else
        
            End If
            'escribimos en el ritch texbox
            n = n + 1
            sRTF = sRTF & n & ".- \b Para: \b0 "
            sRTF = sRTF & rst("Celular") & "\par "
            sRTF = sRTF & "\tab " & msg & "\par "
            sRTF = sRTF & String(100, "-") & " \par "
        End If
        rst.MoveNext
        
    Loop Until rst.EOF
    
End If
rst.Close
Set rst = Nothing
cnn.Close
Set cnn = Nothing

End Sub
