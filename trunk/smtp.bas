Attribute VB_Name = "smtp"
Option Explicit

Public Enum CONEXION
    CONECTED = 0
    MailFrom = 1
    RCPTTO = 2
    DataC = 3
    MESSAGGE = 4
    Quit = 5
End Enum
Public SendStatus As CONEXION
Public mRespuesta As String
Public Code As Integer
Public DServer As String
Public DHelo As String
Public DMailFrom As String
Public DRcptTo As String
Public DSubject As String
Public DMensaje As String
Public DFrom As String
Public exCaption As String

'conectar al servidor
Sub Conectar()
    With FrmAdmin
        '.cmdEnviar.Enabled = False
        '.CmdCancel.Visible = True
        '.Refresh
        '.Caption = "Enviando..."
        .wskMail.Close
        .wskMail.Connect DServer, 25
    End With
'    AddStatus ("Conectando a " & DServer & "... " & Now)
End Sub

'cerrar coneccion
Sub DesConectar()
    SendStatus = CONECTED
'    Call AddStatus("Desconectado")
    With FrmAdmin
'        .CmdCancel.Visible = False
'        .Caption = exCaption
'        .cmdEnviar.Enabled = True
        .wskMail.Close
    End With
End Sub

'agregar status
Sub AddStatus(Texto As String)
    frmMain.txtStatus = frmMain.txtStatus & vbCrLf & Texto
    frmMain.txtStatus.SelStart = Len(frmMain.txtStatus.Text)
    frmMain.txtStatus.Refresh
End Sub

'generador de codigos alfanumericos
Function GenerateCode(NumChar As Integer)
    Randomize Timer
    Dim Code As String
    Dim Chars As Integer
    Dim Alfa As Integer
    Code = ""
    For Chars = 1 To NumChar
        Alfa = Int(Rnd * 2 + 1)
        If Alfa = 2 Then
            Code = Chr(Int((Rnd * 25 + 1) + 97)) & Code
        Else
            Code = Int((Rnd * 9 + 1)) & Code
        End If
    Next
    GenerateCode = Code
End Function

Public Function Enviar(From As String, MailFrom As String, MailTo As String, Subject As String, Mensaje As String)
    If Not Conectado Then
        AddStatus "No conectado a Internet para mandar: " & From & " " & Now
        Exit Function
    End If
        
    DHelo = GenerateCode(8)
    DMailFrom = MailFrom
    DFrom = From
    DSubject = Subject
    DMensaje = Mensaje
    DRcptTo = MailTo
    Call Conectar

End Function
