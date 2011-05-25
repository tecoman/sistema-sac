VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Object = "{8767A745-088E-4CA6-8594-073D6D2DE57A}#9.2#0"; "crviewer9.dll"
Object = "{20C62CAE-15DA-101B-B9A8-444553540000}#1.1#0"; "MSMAPI32.OCX"
Begin VB.Form frmAC 
   Caption         =   "Form1"
   ClientHeight    =   6675
   ClientLeft      =   960
   ClientTop       =   1290
   ClientWidth     =   11325
   ControlBox      =   0   'False
   Icon            =   "frmAC.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6675
   ScaleWidth      =   11325
   WindowState     =   2  'Maximized
   Begin CRVIEWER9LibCtl.CRViewer9 crView 
      Height          =   5550
      Left            =   165
      TabIndex        =   4
      Top             =   195
      Width           =   7920
      lastProp        =   500
      _cx             =   13970
      _cy             =   9790
      DisplayGroupTree=   -1  'True
      DisplayToolbar  =   -1  'True
      EnableGroupTree =   -1  'True
      EnableNavigationControls=   -1  'True
      EnableStopButton=   -1  'True
      EnablePrintButton=   -1  'True
      EnableZoomControl=   -1  'True
      EnableCloseButton=   -1  'True
      EnableProgressControl=   -1  'True
      EnableSearchControl=   -1  'True
      EnableRefreshButton=   -1  'True
      EnableDrillDown =   -1  'True
      EnableAnimationControl=   -1  'True
      EnableSelectExpertButton=   0   'False
      EnableToolbar   =   -1  'True
      DisplayBorder   =   -1  'True
      DisplayTabs     =   -1  'True
      DisplayBackgroundEdge=   -1  'True
      SelectionFormula=   ""
      EnablePopupMenu =   -1  'True
      EnableExportButton=   0   'False
      EnableSearchExpertButton=   0   'False
      EnableHelpButton=   0   'False
      LaunchHTTPHyperlinksInNewBrowser=   -1  'True
   End
   Begin VB.CommandButton cmd 
      Caption         =   "email"
      Height          =   720
      Index           =   2
      Left            =   6780
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   5850
      Width           =   1410
   End
   Begin VB.CommandButton cmd 
      Caption         =   "&Imprimir"
      Height          =   720
      Index           =   1
      Left            =   8280
      TabIndex        =   2
      Top             =   5850
      Width           =   1410
   End
   Begin VB.CommandButton cmd 
      Caption         =   "&Cerrar"
      Height          =   720
      Index           =   0
      Left            =   9750
      TabIndex        =   1
      Top             =   5850
      Width           =   1410
   End
   Begin MSMAPI.MAPIMessages MAPm 
      Left            =   630
      Top             =   5790
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      AddressEditFieldCount=   1
      AddressModifiable=   0   'False
      AddressResolveUI=   0   'False
      FetchSorted     =   0   'False
      FetchUnreadOnly =   0   'False
   End
   Begin MSMAPI.MAPISession MAPs 
      Left            =   0
      Top             =   5775
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DownloadMail    =   -1  'True
      LogonUI         =   -1  'True
      NewSession      =   0   'False
   End
   Begin SHDocVwCtl.WebBrowser wb 
      Height          =   5550
      Left            =   180
      TabIndex        =   0
      Top             =   210
      Width           =   10980
      ExtentX         =   19368
      ExtentY         =   9790
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   "http:///"
   End
End
Attribute VB_Name = "frmAC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'variables publicas
Public strArchivo As String
Public strEmail As String
Private WithEvents poSendMail As clsSendMail
Attribute poSendMail.VB_VarHelpID = -1

Private Sub cmd_Click(Index As Integer)
'variables locales
Dim errLocal As Long
Dim gsTEMPDIR As String
Dim lchar As Long
Dim m_report As CRAXDRT.Report
Dim strTitulo As String


Select Case Index
    
    Case 0  'cerrar
        Me.Hide
        Unload Me
        Set frmAC = Nothing
        
    Case 1  'imprimir
        If crView.Visible Then
            Set m_report = crView.ReportSource
            m_report.PrintOut False
            Set m_report = Nothing
        Else
            wb.ExecWB OLECMDID_PRINT, OLECMDEXECOPT_DONTPROMPTUSER, 1, 1
        End If
    
    Case 2  'enviar por email
        MousePointer = vbHourglass
        On Error GoTo Salir:
        If strEmail <> "" Then  'envia el email
            strTitulo = Mid(Replace(Me.Caption, "-", ""), 5, Len(Me.Caption))
            
            If Me.crView.Visible Then
                gsTEMPDIR = String$(255, 0)
                lchar = GetTempPath(255, gsTEMPDIR)
                gsTEMPDIR = Left(gsTEMPDIR, lchar)
                
                Set m_report = crView.ReportSource
                m_report.DisplayProgressDialog = False
                m_report.ExportOptions.DestinationType = crEDTDiskFile
                m_report.ExportOptions.FormatType = crEFTPortableDocFormat
                strArchivo = gsTEMPDIR & "AC" & strTitulo & ".pdf"
                m_report.ExportOptions.DiskFileName = strArchivo
                m_report.Export (False)
                
            End If
'            For I = 0 To 100000
'                'retardo
'            Next
            Set m_report = Nothing
            Dim Dir1$, Dir2$, Subject$
            If InStr(strEmail, ";") Then
                Dir1 = Left(strEmail, InStr(strEmail, ";") - 1)
                Dir2 = Mid(strEmail, InStr(strEmail, ";") + 1, Len(strEmail))
            Else
                Dir1 = strEmail
            End If
            Subject = "Estimado cliente, adjunto le estamos enviando su aviso de cobro, correpondiente al " & Me.Caption
            If ModGeneral.enviar_email(Dir1, "pagoscondominio@administradorasac.com", "Aviso de Cobro " & Me.Caption, _
            True, Subject, strArchivo) Then
                MsgBox "Mensaje enviado con éxito", vbInformation, strEmail
            Else
                MsgBox "Error al enviar mensaje." & vbCrLf & Err.Description, vbCritical, strEmail
            End If
            
            'envia el archivo via mail
'            Set poSendMail = New clsSendMail
'            Dim Dir1$, Dir2$
'            With poSendMail
'                .SMTPHostValidation = VALIDATE_HOST_DNS
'                .EmailAddressValidation = VALIDATE_SYNTAX
'                .Delimiter = ";"
'                .SMTPHost = "mail.cantv.net"
'                .from = IIf(gcNivel = nuADSYS, "sistemas@administradorasac.com", "info@administradorasac.com")
'                .FromDisplayName = "Servicio Administración de Condominio"
'
'                If InStr(strEmail, ";") Then
'
'                    Dir1 = Left(strEmail, InStr(strEmail, ";") - 1)
'                    Dir2 = Mid(strEmail, InStr(strEmail, ";") + 1, Len(strEmail))
'
'                    'mail.AddAddress Dir1, FrmConsultaCxC.Dat(3)
'                    'mail.AddAddress Dir2, FrmConsultaCxC.Dat(3)
'                    .Recipient = Dir1 & ";" & Dir2
'                    .RecipientDisplayName = FrmConsultaCxC.Dat(3)
'                    '.CcRecipient = Dir2
'                    '.CcDisplayName = FrmConsultaCxC.Dat(3)
'                Else
'                    .Recipient = strEmail
'                    .RecipientDisplayName = FrmConsultaCxC.Dat(3)
'
'                    'mail.AddAddress strEmail, FrmConsultaCxC.Dat(3)
'                End If
'                .Subject = "Aviso de Cobro " & Me.Caption
'                .Message = Subject
'                .Attachment = strArchivo
'                .Send
'
'                'mail.Subject = "Aviso de Cobro " & Me.Caption
'                'mail.Body = Subjet
'                'mail.AddAttachment strArchivo
'                'mail.Send
'
'            End With
'            MsgBox "Mail enviado OK", vbInformation, strEmail
'            Set mail = Nothing
            If Me.crView.Visible Then Kill (strArchivo)
        Else
            MsgBox "Este propietario no tiene email registrado", vbCritical, App.ProductName
        End If
        MousePointer = vbDefault
Salir:
    If Err.Number <> 0 Then
        MsgBox Err.Description, vbCritical, "Error " & Err.Number
        'If MAPs.SessionID <> 0 Then MAPs.SignOff
    End If
End Select
'
End Sub

'
Private Sub Form_Load()
'variables locales
wb.Navigate strArchivo
cmd(2).Picture = LoadPicture(gcUbiGraf & "email.ico", vbLPSmall)
'
End Sub

Private Sub Form_Resize()
On Error Resume Next
crView.Top = 100
crView.Left = 200
crView.Width = Me.ScaleWidth - 200
crView.Height = Me.ScaleHeight - cmd(1).Height - 400
crView.ViewReport
crView.Zoom (92)

wb.Top = crView.Top
wb.Left = crView.Left
wb.Height = crView.Height
wb.Width = crView.Width
For i = 0 To 2
    cmd(i).Left = Me.ScaleWidth - (cmd(i).Width * (i + 1)) - 200
    cmd(i).Top = crView.Height + crView.Top + 100
Next

End Sub
