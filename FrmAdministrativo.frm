VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Begin VB.Form FrmAdministrativo 
   Caption         =   "Administrativo"
   ClientHeight    =   3000
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4980
   ControlBox      =   0   'False
   Icon            =   "FrmAdministrativo.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   3000
   ScaleWidth      =   4980
   Begin VB.Frame fraAdm 
      Height          =   2910
      Index           =   0
      Left            =   90
      TabIndex        =   0
      Top             =   -60
      Width           =   4785
      Begin ComctlLib.Toolbar Toolbar1 
         Height          =   450
         Left            =   60
         TabIndex        =   1
         Top             =   135
         Width           =   2130
         _ExtentX        =   3757
         _ExtentY        =   794
         ButtonWidth     =   714
         ButtonHeight    =   688
         ImageList       =   "ImageList1"
         _Version        =   327682
         BeginProperty Buttons {0713E452-850A-101B-AFC0-4210102A8DA7} 
            NumButtons      =   5
            BeginProperty Button1 {0713F354-850A-101B-AFC0-4210102A8DA7} 
               Key             =   "Edit"
               Object.ToolTipText     =   "Editar Registro"
               Object.Tag             =   ""
               ImageIndex      =   10
            EndProperty
            BeginProperty Button2 {0713F354-850A-101B-AFC0-4210102A8DA7} 
               Key             =   "Save"
               Object.ToolTipText     =   "Guardar Registro"
               Object.Tag             =   ""
               ImageIndex      =   6
            EndProperty
            BeginProperty Button3 {0713F354-850A-101B-AFC0-4210102A8DA7} 
               Key             =   "Cancel"
               Object.ToolTipText     =   "Cancelar Registro"
               Object.Tag             =   ""
               ImageIndex      =   8
            EndProperty
            BeginProperty Button4 {0713F354-850A-101B-AFC0-4210102A8DA7} 
               Key             =   "Print"
               Object.ToolTipText     =   "Imprimir"
               Object.Tag             =   ""
               ImageIndex      =   11
            EndProperty
            BeginProperty Button5 {0713F354-850A-101B-AFC0-4210102A8DA7} 
               Key             =   "Close"
               Object.ToolTipText     =   "Salir"
               Object.Tag             =   ""
               ImageIndex      =   12
            EndProperty
         EndProperty
      End
      Begin VB.Frame fraAdm 
         Enabled         =   0   'False
         Height          =   2085
         Index           =   1
         Left            =   210
         TabIndex        =   2
         Top             =   630
         Width           =   4350
         Begin VB.TextBox txtAdm 
            DataField       =   "ServidorH"
            BeginProperty DataFormat 
               Type            =   0
               Format          =   "#,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   8202
               SubFormatType   =   0
            EndProperty
            Height          =   315
            Index           =   4
            Left            =   1785
            TabIndex        =   11
            Tag             =   "RUTA"
            Top             =   1150
            Width           =   2340
         End
         Begin VB.TextBox txtAdm 
            Alignment       =   1  'Right Justify
            DataField       =   "IDMoneda"
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "#,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   8202
               SubFormatType   =   1
            EndProperty
            Height          =   315
            Index           =   3
            Left            =   3315
            TabIndex        =   10
            Tag             =   "IDMONEDA"
            Top             =   300
            Width           =   840
         End
         Begin VB.TextBox txtAdm 
            DataField       =   "Ruta"
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "#,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   8202
               SubFormatType   =   1
            EndProperty
            Height          =   315
            Index           =   2
            Left            =   1785
            TabIndex        =   8
            Tag             =   "RUTA"
            Top             =   1560
            Width           =   2340
         End
         Begin VB.TextBox txtAdm 
            Alignment       =   1  'Right Justify
            DataField       =   "IDB"
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "#,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   8202
               SubFormatType   =   1
            EndProperty
            Height          =   315
            Index           =   1
            Left            =   1785
            TabIndex        =   7
            Tag             =   "IDB"
            Top             =   750
            Width           =   840
         End
         Begin VB.TextBox txtAdm 
            Alignment       =   1  'Right Justify
            DataField       =   "IVA"
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "#,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   8202
               SubFormatType   =   0
            EndProperty
            Height          =   315
            Index           =   0
            Left            =   975
            TabIndex        =   6
            Tag             =   "IVA"
            Top             =   300
            Width           =   810
         End
         Begin VB.CheckBox Check1 
            Caption         =   "Débito Bancario"
            Height          =   300
            Index           =   1
            Left            =   240
            TabIndex        =   4
            Top             =   825
            Value           =   1  'Checked
            Width           =   1515
         End
         Begin VB.CheckBox Check1 
            Caption         =   "I.V.A."
            Height          =   300
            Index           =   0
            Left            =   240
            TabIndex        =   3
            Top             =   345
            Value           =   1  'Checked
            Width           =   1065
         End
         Begin VB.Label Label1 
            Caption         =   "Servidor de Hora:"
            Height          =   180
            Index           =   2
            Left            =   240
            TabIndex        =   12
            Top             =   1250
            Width           =   1275
         End
         Begin VB.Label Label1 
            Caption         =   "Símbolo Moneda:"
            Height          =   300
            Index           =   1
            Left            =   1995
            TabIndex        =   9
            Top             =   375
            Width           =   1275
         End
         Begin VB.Label Label1 
            Caption         =   "Ubicación Datos:"
            Height          =   180
            Index           =   0
            Left            =   240
            TabIndex        =   5
            Top             =   1620
            Width           =   1275
         End
      End
   End
   Begin ComctlLib.ImageList ImageList1 
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   20
      ImageHeight     =   20
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   12
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmAdministrativo.frx":0442
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmAdministrativo.frx":05C4
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmAdministrativo.frx":0746
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmAdministrativo.frx":08C8
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmAdministrativo.frx":0A4A
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmAdministrativo.frx":0BCC
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmAdministrativo.frx":0D4E
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmAdministrativo.frx":0ED0
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmAdministrativo.frx":1052
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmAdministrativo.frx":11D4
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmAdministrativo.frx":1356
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmAdministrativo.frx":14D8
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "FrmAdministrativo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
    Dim rstAmbiente As New ADODB.Recordset

    Private Sub Check1_Click(Index As Integer)
    txtAdm(Index).Enabled = Not txtAdm(Index).Enabled
    End Sub

    Private Sub Form_Load()
    CenterForm Me
    rstAmbiente.Open "SELECT * FROM Ambiente", cnnConexion, adOpenKeyset, adLockOptimistic, _
    adCmdText
    For i = txtAdm.LBound To txtAdm.UBound
        Set txtAdm(i).DataSource = rstAmbiente
    Next i
    End Sub

    Private Sub Form_Unload(Cancel As Integer)
    rstAmbiente.Close
    Set rstAmbiente = Nothing
    End Sub

    Private Sub Toolbar1_ButtonClick(ByVal Button As ComctlLib.Button)
    '
    With rstAmbiente
        
        Select Case Button.Key
          
            Case "Save" 'Actualizar el Registro
            
                If Dir(txtAdm(2) & "\sac.mdb") = "" Then
                    MsgBox "Ubicación incorrecta", vbInformation, App.ProductName
                    Exit Sub
                End If
                '
                Call rtnEstBar(Button.Key)
                Call rtnBitacora("Actualizar Parám.Adm.")
                !Ruta = txtAdm(2)
                !IVA = IIf(txtAdm(0).Enabled, CCur(txtAdm(0)), 0)
                !IDB = IIf(txtAdm(1).Enabled, CCur(txtAdm(1)), 0)
                !Usuario = gcUsuario
                !FecAct = Date
                .Update
                For i = txtAdm.LBound To txtAdm.UBound
                    SaveSetting App.EXEName, "Entorno", txtAdm(i).Tag, txtAdm(i)
                Next i
                MsgBox "Registro Actualizado...Para que los cambios surtan efecto todos los " _
                & vbCrLf & "usuarios deben cerrar y volver abrir el sistema", vbInformation, _
                App.ProductName
                fraAdm(1).Enabled = False
                
            Case "Cancel"  'Cancelar Registro
                '
                Call rtnEstBar(Button.Key)
                rtnBitacora ("Cancelar Editar Parám.Adm.")
                .CancelUpdate
                MsgBox " Registro Cancelado ... ", vbInformation, App.ProductName
                fraAdm(1).Enabled = False
                
            Case "Edit" 'Editar Registro
                Call rtnEstBar(Button.Key)
                Call rtnBitacora("Editar Parám.Adm.")
                fraAdm(1).Enabled = True
                txtAdm(0).SelStart = 0
                txtAdm(0).SelLength = Len(txtAdm(0))
                txtAdm(0).SetFocus
                
            Case "Print"    'Imprimir Registro
                MsgBox "Opción No Disponible", vbInformation, App.ProductName
                
            Case "Close": 'Descargar Formulario
                Unload Me
                Set FrmAdministrativo = Nothing
        End Select
        '
     End With
    
    End Sub

    
    Private Sub txtAdm_KeyPress(Index As Integer, KeyAscii As Integer)
    '
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If Index = 0 Or Index = 1 Then
        If KeyAscii = 46 Then KeyAscii = 44
        Call Validacion(KeyAscii, "1234567890,")
    End If
    End Sub

    Private Sub rtnEstBar(strBoton As String)
    '
    Select Case strBoton
    
        Case "Edit"
            With Toolbar1
                .Buttons("Edit").Enabled = False
                .Buttons("Save").Enabled = True
                .Buttons("Cancel").Enabled = True
                .Buttons("Print").Enabled = False
                .Buttons("Close").Enabled = False
            End With
            '
        Case "Save", "Cancel", "Print"
        
            With Toolbar1
                .Buttons("Edit").Enabled = True
                .Buttons("Save").Enabled = True
                .Buttons("Cancel").Enabled = True
                .Buttons("Print").Enabled = True
                .Buttons("Close").Enabled = True
            End With
            '
    End Select
    '
    End Sub
