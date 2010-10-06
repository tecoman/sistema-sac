VERSION 5.00
Begin VB.Form frmPausa 
   Caption         =   "Pausa ||"
   ClientHeight    =   1875
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   4950
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   Moveable        =   0   'False
   ScaleHeight     =   1875
   ScaleWidth      =   4950
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txt 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      IMEMode         =   3  'DISABLE
      Left            =   3165
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   225
      Width           =   1590
   End
   Begin VB.CommandButton cmd 
      Caption         =   "&Reanudar"
      Height          =   645
      Left            =   3150
      TabIndex        =   2
      Top             =   1020
      Width           =   1635
   End
   Begin VB.Label lbl 
      Caption         =   "&Ingrese su clave de acceso, y presione el botón Reanudar para continuar:"
      Height          =   675
      Left            =   300
      TabIndex        =   0
      Top             =   225
      Width           =   2745
   End
End
Attribute VB_Name = "frmPausa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
    Private Sub cmd_Click()
    '
    Static Intento%
    '
    If UCase(txt) = gcContraseña Then
        cnnConexion.Execute "UPDATE Usuarios in '" & gcPath & "\tablas.mdb' SET " _
        & "LogIn = True WHERE NombreUsuario ='" & gcUsuario & "'"
        Call rtnBitacora("Sac Reanudado...")
        Unload Me
    Else
    '
        Intento = Intento + 1
        If Intento = 3 Then
            Call rtnBitacora("Transgresión de seguridad al reanudar SAC...")
            MsgBox "Ha transgredido niveles de seguridad, SAC se cerrará...", vbExclamation, _
            App.ProductName
            cnnConexion.Execute "UPDATE Usuarios in '" & gcPath & "\tablas.mdb' SET " _
            & "LogIn = false WHERE NombreUsuario ='" & gcUsuario & "'"
            Unload FrmAdmin
            End
            '
        Else
        '
            If txt = "" Then
                MsgBox "Introduzca su contraseña por favor..", vbInformation, App.ProductName
                txt.SetFocus
                
            Else
                MsgBox "Contraseña Invalida", vbCritical, App.ProductName
                txt.SelStart = 0
                txt.SelLength = Len(txt)
                 
            End If
        End If
        '
    End If
    '
    End Sub

Private Sub Form_Load()
cnnConexion.Execute "UPDATE Usuarios in '" & gcPath & "\tablas.mdb' SET " _
& "LogIn = False  WHERE NombreUsuario ='" & gcUsuario & "'"

End Sub

Private Sub txt_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then cmd_Click
End Sub
