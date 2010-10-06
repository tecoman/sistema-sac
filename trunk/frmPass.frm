VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmPass 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Cambiar Contraseña"
   ClientHeight    =   2790
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   4155
   Icon            =   "frmPass.frx":0000
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2790
   ScaleWidth      =   4155
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
   Begin MSComctlLib.ImageList imglist 
      Left            =   3600
      Top             =   1080
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   2  'Align Bottom
      Height          =   330
      Left            =   0
      TabIndex        =   5
      Top             =   2460
      Width           =   4155
      _ExtentX        =   7329
      _ExtentY        =   582
      ButtonWidth     =   2090
      ButtonHeight    =   582
      Wrappable       =   0   'False
      Style           =   1
      TextAlignment   =   1
      ImageList       =   "imglist"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   9
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Aceptar   "
            Key             =   "OK"
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
            Object.Width           =   100
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Salir     "
            Key             =   "Out"
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
      EndProperty
   End
   Begin VB.TextBox txtPass 
      Height          =   315
      IMEMode         =   3  'DISABLE
      Index           =   1
      Left            =   1950
      MaxLength       =   7
      PasswordChar    =   "*"
      TabIndex        =   1
      ToolTipText     =   "Confirmación nueva contraseña"
      Top             =   1530
      Width           =   1770
   End
   Begin VB.TextBox txtPass 
      Height          =   315
      IMEMode         =   3  'DISABLE
      Index           =   0
      Left            =   1950
      MaxLength       =   7
      PasswordChar    =   "*"
      TabIndex        =   0
      ToolTipText     =   "Nueva contraseña"
      Top             =   885
      Width           =   1770
   End
   Begin VB.Label lblPass 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Introduzca la nueva contraseña. Con un Mínimo de 3 dìgitos y máximo 7."
      Height          =   540
      Index           =   2
      Left            =   60
      TabIndex        =   4
      Top             =   225
      Width           =   3945
   End
   Begin VB.Label lblPass 
      BackStyle       =   0  'Transparent
      Caption         =   "Confirme la nueva Contraseña:"
      Height          =   495
      Index           =   1
      Left            =   225
      TabIndex        =   3
      Top             =   1530
      Width           =   1560
   End
   Begin VB.Label lblPass 
      BackStyle       =   0  'Transparent
      Caption         =   "Escriba la nueva contraseña:"
      Height          =   540
      Index           =   0
      Left            =   270
      TabIndex        =   2
      Top             =   885
      Width           =   1575
   End
End
Attribute VB_Name = "frmPass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
    


    Private Sub Form_Load()
    imglist.ListImages.Add 1, "", LoadResPicture("SALIR", vbResIcon)
    imglist.ListImages.Add 2, "", LoadResPicture("OK", vbResIcon)
    
    Toolbar1.Buttons("OK").Image = 2
    Toolbar1.Buttons("Out").Image = 1
    CenterForm Me
    End Sub


    Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Key
        Case "OK"
            For i = 0 To 1
                If txtPass(i) = "" Then
                    MsgBox "Introduzca " & txtPass(i).ToolTipText, vbInformation, App.ProductName
                    Exit Sub
                End If
            Next
            If txtPass(0) = txtPass(1) Then 'las contraseñas coinciden
                Dim cnnPass As New ADODB.Connection
                cnnPass.Open cnnOLEDB & gcPath & "/tablas.mdb"
                cnnPass.Execute "UPDATE Usuarios SET Contraseña='" & txtPass(0) & "' WHERE NombreUsuari" _
                & "o ='" & gcUsuario & "';"
                MsgBox "Contraseña Atualizada...", vbInformation, App.ProductName
                Unload frmPass: Set frmPass = Nothing
            Else
                MsgBox "Verifique ambas contraseñas, No Coinciden..", vbExclamation, App.ProductName
            End If
        Case "Out"
            Unload Me
    End Select
    End Sub

    Private Sub txtPass_KeyPress(Index As Integer, KeyAscii As Integer)
    Dim boton As MSComctlLib.Button
    KeyAscii = Asc(UCase(Chr(KeyAscii)))    'todo a mayúsculas
    If KeyAscii = 13 Then
        If Index = 0 Then
            txtPass(1).SetFocus
        Else
            Set boton = Toolbar1.Buttons("OK")
            Call Toolbar1_ButtonClick(boton)
        End If
    End If
    '
    End Sub
