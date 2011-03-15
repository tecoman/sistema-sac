VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmMail 
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
      Caption         =   "Seleccione las opciones del caso:"
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
      Top             =   315
      Width           =   14595
      Begin VB.ListBox List 
         Appearance      =   0  'Flat
         Height          =   1395
         Index           =   2
         Left            =   9630
         TabIndex        =   22
         Top             =   1980
         Width           =   4485
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
         Picture         =   "frmMail.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   600
         Width           =   1380
      End
      Begin VB.CommandButton Cmd 
         Caption         =   "&Quitar Archivo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   3
         Left            =   10275
         TabIndex        =   20
         Top             =   3480
         Width           =   3345
      End
      Begin VB.CommandButton Cmd 
         Caption         =   "&Enviar Mensaje"
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
         Picture         =   "frmMail.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   600
         Width           =   1380
      End
      Begin VB.CommandButton Cmd 
         Caption         =   "&Adjuntar Archivo"
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
         Picture         =   "frmMail.frx":074C
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   600
         Width           =   1380
      End
      Begin VB.ListBox List 
         Appearance      =   0  'Flat
         Height          =   2565
         Index           =   1
         Left            =   825
         TabIndex        =   11
         Top             =   3225
         Visible         =   0   'False
         Width           =   4755
      End
      Begin VB.Frame fra 
         Caption         =   "Detalle del Mensaje:"
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
         TabIndex        =   13
         Top             =   3825
         Width           =   13875
         Begin ComctlLib.ProgressBar pBar 
            Height          =   375
            Left            =   360
            TabIndex        =   25
            Top             =   4680
            Visible         =   0   'False
            Width           =   13155
            _ExtentX        =   23204
            _ExtentY        =   661
            _Version        =   327682
            Appearance      =   0
            Min             =   1e-4
         End
         Begin VB.TextBox txt 
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   3510
            Index           =   1
            Left            =   360
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   16
            Top             =   1005
            Width           =   13125
         End
         Begin VB.TextBox txt 
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   0
            Left            =   1170
            TabIndex        =   15
            Top             =   405
            Width           =   12330
         End
         Begin VB.Shape Shape1 
            BackColor       =   &H8000000C&
            BackStyle       =   1  'Opaque
            BorderColor     =   &H8000000C&
            Height          =   3540
            Index           =   2
            Left            =   345
            Top             =   990
            Width           =   13155
         End
         Begin VB.Shape Shape1 
            BackColor       =   &H8000000C&
            BackStyle       =   1  'Opaque
            BorderColor     =   &H8000000C&
            Height          =   345
            Index           =   1
            Left            =   1155
            Top             =   390
            Width           =   12360
         End
         Begin VB.Label lbl 
            Caption         =   "Asunto"
            Height          =   285
            Index           =   3
            Left            =   360
            TabIndex        =   14
            Top             =   465
            Width           =   705
         End
      End
      Begin VB.CommandButton Cmd 
         Caption         =   "&Cerrar Selección"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   0
         Left            =   5640
         TabIndex        =   12
         Top             =   3480
         Visible         =   0   'False
         Width           =   3315
      End
      Begin VB.Frame fra 
         Caption         =   "Seleccione Destinatario(s):"
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
         Width           =   8595
         Begin VB.OptionButton Opt 
            Caption         =   "Selección"
            Height          =   285
            Index           =   2
            Left            =   420
            TabIndex        =   7
            Top             =   1215
            Width           =   1950
         End
         Begin VB.OptionButton Opt 
            Caption         =   "Junta de Condominio."
            Height          =   285
            Index           =   1
            Left            =   420
            TabIndex        =   6
            Top             =   810
            Width           =   1950
         End
         Begin VB.OptionButton Opt 
            Caption         =   "Todos los propietarios."
            Height          =   285
            Index           =   0
            Left            =   420
            TabIndex        =   5
            Top             =   405
            Value           =   -1  'True
            Width           =   1950
         End
         Begin VB.ListBox List 
            Appearance      =   0  'Flat
            Height          =   1005
            Index           =   3
            Left            =   2790
            TabIndex        =   23
            Top             =   465
            Width           =   5505
         End
         Begin VB.ListBox List 
            Appearance      =   0  'Flat
            Height          =   1005
            Index           =   0
            Left            =   2790
            TabIndex        =   10
            Top             =   465
            Width           =   5505
         End
      End
      Begin MSDataListLib.DataCombo dtc 
         Height          =   315
         Index           =   0
         Left            =   1455
         TabIndex        =   8
         Top             =   975
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
         Top             =   975
         Width           =   5070
         _ExtentX        =   8943
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
      End
      Begin VB.ListBox List 
         Appearance      =   0  'Flat
         Height          =   2565
         Index           =   4
         Left            =   825
         TabIndex        =   24
         Top             =   3225
         Visible         =   0   'False
         Width           =   4755
      End
      Begin VB.Label lbl 
         Caption         =   "Archivo(s) Adjunto(s) Máx 3:"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   4
         Left            =   9630
         TabIndex        =   19
         Top             =   1755
         Width           =   2505
      End
      Begin VB.Label lbl 
         Caption         =   "Nombre:"
         Height          =   285
         Index           =   1
         Left            =   3270
         TabIndex        =   3
         Top             =   990
         Width           =   795
      End
      Begin VB.Label lbl 
         Caption         =   "Codigo:"
         Height          =   285
         Index           =   0
         Left            =   825
         TabIndex        =   2
         Top             =   990
         Width           =   705
      End
      Begin VB.Label lbl 
         Caption         =   "Información sobre el inmueble:"
         Height          =   285
         Index           =   2
         Left            =   300
         TabIndex        =   1
         Top             =   480
         Width           =   2670
      End
   End
End
Attribute VB_Name = "frmMail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public sArchivo As String
Dim WithEvents poSendMail As clsSendMail
Attribute poSendMail.VB_VarHelpID = -1


Private Sub cmd_Click(Index As Integer)
On Error Resume Next
Select Case Index
    Case 4
        Unload Me
        Set frmMail = Nothing
    
    Case 1  'adjutar archivo
        If list(2).ListCount <= 3 Then
            With FrmAdmin.CDlMain
            
                 .CancelError = True
                .Filter = "Todos los Archivos (*.*)|*.*"
                .FilterIndex = 1
                .DialogTitle = "Seleccione el archivo...."
                .ShowOpen
                
                If Err.Number = cdlCancel Then Exit Sub
                list(2).AddItem .FileName
            End With
        Else
            MsgBox "Solo puede adjuntar hasta 3 archivos", vbInformation, App.ProductName
        End If
    
    Case 0
        list(1).Clear
        list(4).Clear
        list(1).Visible = False
        cmd(0).Visible = False
    
    Case 2
        Call enviar_email
    
    Case 3
        list(2).RemoveItem list(2).ListIndex
End Select
End Sub

Private Sub dtc_Click(Index As Integer, Area As Integer)
Dim rstlocal As ADODB.Recordset

If Area = 2 Then
    
    list(0).Clear
    list(1).Clear
    list(2).Clear
    list(3).Clear
    list(4).Clear
    opt(0).Value = True
    dtc(IIf(Index = 0, 1, 0)) = dtc(Index).BoundText
    If Not dtc(IIf(Index = 0, 1, 0)).MatchedWithList Then
        dtc(IIf(Index = 0, 1, 0)) = ""
    Else
        Call Opt_Click(0)
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
'-----------------------------------------------------
If sArchivo <> "" Then 'envia un reporte por mail
    
    dtc(0) = gcCodInm
    Call dtc_Click(0, 2)
    opt(2).Value = True
    list(2).AddItem sArchivo
'Else
'    MsgBox "No se encuentra el archivo", vbInformation, App.ProductName
End If

End Sub

Private Sub List_Click(Index As Integer)
Select Case Index
    Case 1
        list(3).AddItem list(1).list(list(1).ListIndex)
        list(0).AddItem list(4).list(list(1).ListIndex)
        list(4).RemoveItem (list(1).ListIndex)
        list(1).RemoveItem (list(1).ListIndex)
        
        
    Case 3
        list(1).AddItem list(3).list(list(3).ListIndex)
        list(4).AddItem list(3).list(list(3).ListIndex)
        list(0).RemoveItem (list(3).ListIndex)
        list(3).RemoveItem (list(3).ListIndex)
        
        
End Select
End Sub

Private Sub Opt_Click(Index As Integer)
Call cargar_destinatarios(Index)
End Sub

Private Sub cargar_destinatarios(Index As Integer)
'variables locales
Dim rstlocal As ADODB.Recordset

list(3).Clear
list(0).Clear
list(1).Visible = False
cmd(0).Visible = False

Select Case Index
    
    Case 0, 2 'todos los propietarios
    
        If dtc(0).MatchedWithList Then
            Set rstlocal = New ADODB.Recordset
            With rstlocal
                .Open "SELECT * FROM Propietarios WHERE email<>''", cnnOLEDB + gcPath + "\" + dtc(0) + _
                "\" + "inm.mdb", adOpenKeyset, adLockOptimistic, adCmdText
                If Not .EOF And Not .BOF Then
                    .MoveFirst
                    If Index = 0 Then
                        list(3).Clear
                        list(0).Clear
                        Do
                            list(3).AddItem !Codigo + " - " + !Nombre
                            list(0).AddItem !email
                            .MoveNext
                        Loop Until .EOF
                    Else
                        list(1).Clear
                        list(4).Clear
                         Do
                            list(1).AddItem !Codigo + " - " + !Nombre
                            list(4).AddItem !email
                            .MoveNext
                        Loop Until .EOF
                        list(1).Visible = True
                        cmd(0).Visible = True
                    End If
                End If
                .Close
            End With
            Set rstlocal = Nothing
        End If
    Case 1  'Junta de condominio
        
        If dtc(0).MatchedWithList Then
            Set rstlocal = New ADODB.Recordset
            With rstlocal
                .Open "SELECT * FROM Propietarios WHERE email<>'' AND CarJunta<>''", cnnOLEDB + gcPath + "\" + _
                dtc(0) + "\" + "inm.mdb", adOpenKeyset, adLockOptimistic, adCmdText
                If Not .EOF And Not .BOF Then
                    .MoveFirst
                    Do
                        list(3).AddItem !Codigo + " - " + !Nombre
                        list(0).AddItem !email
                        .MoveNext
                    Loop Until .EOF
                End If
                .Close
            End With
            Set rstlocal = Nothing
        End If
        
End Select
End Sub



Private Sub enviar_email()
'variables locales
'Set poSendMail = New clsSendMail

Dim I%, Y%, n%, msg$, Dir1$, Dir2$, archivos$
'valida los campos necesarios para enviar el email
If Txt(0) = "" Then msg = "- Falta el sujeto del mensaje." & vbCrLf
If Txt(1) = "" Then msg = msg + "- Debe escribir en el cuerpo del mensaje." & vbCrLf
If list(0).ListCount = 0 Then msg = msg + "- Agregue destinatario(s) a su mensaje." & vbCrLf
If msg <> "" Then
    MsgBox "No se puede enviar el mensaje:" & vbCrLf & vbCrLf & msg, vbCritical, App.ProductName
    Exit Sub
End If
pBar.Visible = True
pBar.Max = list(0).ListCount
MousePointer = vbHourglass
cmd(4).Enabled = False
cmd(1).Enabled = False
cmd(2).Enabled = False
For I = 0 To list(2).ListCount - 1
    archivos = archivos & list(2).list(I) & IIf(I = list(2).ListCount - 1, "", ",")
Next

For I = 0 To list(0).ListCount - 1
    
    If InStr(list(0).list(I), ";") Then
        Dir1 = Left(list(0).list(I), InStr(list(0).list(I), ";") - 1)
    Else
        Dir1 = list(0).list(I)
    End If
    If ModGeneral.enviar_email(Dir1, "administracion@administradorasac.com", Txt(0), True, Txt(1), archivos) Then
        n = n + 1
        pBar.Value = n
    Else
        If Err = 2 Then
            MsgBox "No se encontró el servidor de correo electrónico. Revise su conexión a Internet." & _
            vbCrLf & "Si el problema persiste pongase en contacto con el administrador del sistema", _
            vbCritical, "No se envió el mensaje"
            Exit Sub
        Else
            MsgBox Err.Description, vbCritical, "Error al enviar: " & Dir1
        End If
    End If
Next
If n > 0 Then
    MsgBox pBar.Max & IIf(pBar.Max > 1, " Mensajes Enviados ", " Mensaje Enviado ") & "con éxito.", _
    vbInformation, App.ProductName
End If
'On Error Resume Next

'With poSendMail
'    .SMTPHostValidation = VALIDATE_HOST_DNS
'    .EmailAddressValidation = VALIDATE_SYNTAX
'    .Delimiter = ";"
'    .SMTPHost = "mail.cantv.net"
'    .from = IIf(gcNivel = nuADSYS, "sistemas@administradorasac.com", "administracion@administradorasac.com.ve")
'    .FromDisplayName = "Servicio de Administración de Condominios"
'
'    'mail.HOST = "mail.cantv.net"
'    'mail.from = IIf(gcNivel = nuADSYS, "sistemas@administradorasac.com", "administracion@administradorasac.com.ve")
'    'mail.FromName = "Servicio de Administración de Condominios"
'    pBar.Visible = True
'    pBar.Max = List(0).ListCount
'    MousePointer = vbHourglass
'    cmd(4).Enabled = False
'    cmd(1).Enabled = False
'    cmd(2).Enabled = False
'    Do
'        'mail.Reset
'        If InStr(List(0).List(I), ";") Then
'            Dir1 = Left(List(0).List(I), InStr(List(0).List(I), ";") - 1)
'            Dir2 = Mid(List(0).List(I), InStr(List(0).List(I), ";") + 1, Len(List(0).List(I)))
'            'mail.AddAddress Dir1, List(3).List(i)
'            'mail.AddAddress Dir2, List(3).List(i)
'            .Recipient = Dir1 & ";" & Dir2
'            .RecipientDisplayName = List(3).List(I)
'            '.Recipient = Dir2
'            '.RecipientDisplayName = List(3).List(i)
'
'        Else
'            'mail.AddAddress List(0).List(i), List(3).List(i)
'            .Recipient = List(0).List(I)
'            .RecipientDisplayName = List(3).List(I)
'
'        End If
'        'mail.AddAddress "ynfantes@cantv.net", "Edgar"
'        I = I + 1
'        pBar.Value = I
'        'mail.Subject = txt(0)
'        'mail.Body = txt(1)
'        .Subject = Txt(0)
'        .Message = Txt(1)
'
'        Y = 0
'        If List(2).ListCount > 0 Then
'            Do
'                'mail.AddAttachment List(2).List(Y)
'                .Attachment = List(2).List(Y)
'                Y = Y + 1
'            Loop Until Y = List(2).ListCount
'        End If
'        .Send
'        'mail.Send
'        If Err = 2 Then
'            MsgBox "No se encontró el servidor de correo electrónico. Revise su conexión a Internet." & _
'            vbCrLf & "Si el problema persiste pongase en contacto con el administrador del sistema", _
'            vbCritical, "No se envió el mensaje"
'            Exit Sub
'        ElseIf Err <> 0 Then
'            MsgBox Err.Description, vbCritical, "Error al enviar: " & List(3).List(I)
'        End If
'
'
'    Loop Until I = List(0).ListCount
'End With

MousePointer = vbDefault
pBar.Visible = False
cmd(4).Enabled = True
cmd(1).Enabled = True
cmd(2).Enabled = True

'If Err = 0 Then
'    MsgBox pBar.Max & IIf(pBar.Max > 1, " Mensajes Enviados ", " Mensaje Enviado ") & "con éxito.", _
'    vbInformation, App.ProductName
'End If
'
End Sub
