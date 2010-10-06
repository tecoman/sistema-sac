VERSION 5.00
Object = "{8767A745-088E-4CA6-8594-073D6D2DE57A}#9.2#0"; "crviewer9.dll"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmView 
   Caption         =   "Form1"
   ClientHeight    =   4785
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8010
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   4785
   ScaleWidth      =   8010
   Visible         =   0   'False
   WindowState     =   2  'Maximized
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   3300
      Top             =   6450
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmView.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmView.frx":059A
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tlbBarra 
      Align           =   1  'Align Top
      Height          =   330
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   8010
      _ExtentX        =   14129
      _ExtentY        =   582
      ButtonWidth     =   1958
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Style           =   1
      TextAlignment   =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   3
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
            Object.Width           =   1e-4
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Cerrar   "
            Key             =   "CLOSE"
            Object.ToolTipText     =   "Cerrar Ventana"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&e-mail"
            Key             =   "EMAIL"
            Object.ToolTipText     =   "Enviar por eMail"
            ImageIndex      =   2
         EndProperty
      EndProperty
   End
   Begin CRVIEWER9LibCtl.CRViewer9 crView 
      Height          =   5550
      Left            =   45
      TabIndex        =   0
      Top             =   720
      Width           =   10995
      lastProp        =   600
      _cx             =   19394
      _cy             =   9790
      DisplayGroupTree=   0   'False
      DisplayToolbar  =   -1  'True
      EnableGroupTree =   0   'False
      EnableNavigationControls=   -1  'True
      EnableStopButton=   0   'False
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
      DisplayBorder   =   0   'False
      DisplayTabs     =   -1  'True
      DisplayBackgroundEdge=   -1  'True
      SelectionFormula=   ""
      EnablePopupMenu =   -1  'True
      EnableExportButton=   -1  'True
      EnableSearchExpertButton=   0   'False
      EnableHelpButton=   0   'False
      LaunchHTTPHyperlinksInNewBrowser=   -1  'True
      EnableLogonPrompts=   -1  'True
   End
End
Attribute VB_Name = "frmView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim l As Boolean
Private Sub Form_Load()
'cmd.Picture = LoadResPicture("Salir", vbResIcon)
'Call Form_Resize
End Sub

Private Sub Form_Resize()
On Error Resume Next
If l Then Exit Sub
With crView
    .Top = 10 + tlbBarra.Height
    .Left = 0
    .Width = Me.ScaleWidth - 100
    .Height = Me.ScaleHeight - 200 - tlbBarra.Height
    'Debug.Print Me.ScaleHeight & "/" & .Height
End With

End Sub

Private Sub Form_Unload(Cancel As Integer)
l = True
End Sub

Private Sub tlbBarra_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim gsTEMPDIR As String
Dim m_report As CRAXDRT.Report
Dim archivo As String
Select Case Button.Key
    Case "CLOSE"
        l = True
        Me.Hide
        Unload Me
    
    Case "EMAIL"
        gsTEMPDIR = String$(255, 0)
        lchar = GetTempPath(255, gsTEMPDIR)
        gsTEMPDIR = Left(gsTEMPDIR, lchar)
        Set m_report = crView.ReportSource
        m_report.DisplayProgressDialog = False
        m_report.ExportOptions.DestinationType = crEDTDiskFile
        m_report.ExportOptions.FormatType = crEFTPortableDocFormat
        archivo = gsTEMPDIR & Replace(Replace(Replace(Replace(Caption, "/", " "), ".", ""), ":", ""), " ", "_") & ".pdf"
        m_report.ExportOptions.DiskFileName = archivo
        m_report.Export (False)
        Set m_report = Nothing
        If Dir(archivo) <> "" Then
            frmMail.sArchivo = archivo
            frmMail.Show
        Else
            MsgBox "Imposible encontrar el archivo para exportar", vbInformation, App.ProductName
        End If
        
End Select
End Sub
