VERSION 5.00
Begin VB.Form frmCalendario 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00400000&
   Caption         =   "Inicio"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   WindowState     =   2  'Maximized
   Begin VB.Timer Tempo 
      Left            =   240
      Top             =   210
   End
   Begin VB.Label lbl 
      Caption         =   "..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8880
      TabIndex        =   0
      Top             =   5490
      Visible         =   0   'False
      Width           =   5190
   End
End
Attribute VB_Name = "frmCalendario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const WEEKS As Integer = 6
Const DAYS_INWEEK As Integer = 7
Const REFORMATION_YEAR As Integer = 1582

Dim MonthAr(WEEKS - 1, DAYS_INWEEK - 1) As Integer
Private Declare Function GetTempFileName Lib "kernel32" Alias "GetTempFileNameA" (ByVal lpszPath As String, ByVal lpPrefixString As String, ByVal wUnique As Long, ByVal lpTempFileName As String) As Long
Dim Path As String

'--------------------------------------------
'   CONSTANTES SERVIDOR SMTP
'--------------------------------------------
Private Const SMTP_SERVER$ = "smtp.gmail.com"
Private Const SMTP_SERVER_PORT& = 465
Private Const SERVER_AUTH  As Boolean = True
Private Const USER_NAME$ = "pagoscondominio@administradorasac.com"
Private Const Password$ = "10537439"
Private Const SSL As Boolean = True

Public cPropietarios As Collection
Dim email As CDO.Message
Dim login As Boolean

Private Function Archivo_Temporal() As String
    Dim sSave As String, hOrgFile As Long, hNewFile As Long, bBytes() As Byte
    Dim sTemp As String, nSize As Long, Ret As Long
    
    sTemp = String(260, 0)

    GetTempFileName Environ("temp"), "TTT", 0, sTemp

    Archivo_Temporal = Left$(sTemp, InStr(1, sTemp, Chr$(0)) - 1)

End Function


Function Cargar(ID As Integer) As String
   
    Path = Archivo_Temporal
    
    Dim aDatos() As Byte
    
    ' lee los datos en el array de bytes
    aDatos = LoadResData(ID, "CUSTOM")
    
    ' abre un archivo para escribir los datos en modo binario
    Open Path For Binary Access Write As #1
    
    ' escribe el array de bytes para
    Put #1, , aDatos
    ' cierra el fichero
    Close
    
    Cargar = Path
    
End Function

Private Sub Form_Load()
Dim anchoPantalla As Integer

Set email = New CDO.Message

anchoPantalla = Screen.Width / Screen.TwipsPerPixelX

If (anchoPantalla >= 1280 And anchoPantalla < 1360) Then
    Me.Picture = LoadPicture(Cargar(IIf(Demo, 108, 106)))
ElseIf (anchoPantalla >= 1360) Then
    Me.Picture = LoadPicture(Cargar(IIf(Demo, 107, 107)))
Else
     Me.Picture = LoadPicture(Cargar(IIf(Demo, 107, 107)))
End If

Me.BackColor = RGB(5, 68, 106)
Call CalendarioActual

End Sub

Private Sub CalendarioActual()
Dim nYear, nMonth, nDay As Double
Dim Century As Double
Dim B, C, D, X As Double
Dim JulianDate As Double
Dim intNumDays As Integer
Dim blnBeforeReformat As Boolean


nYear = Year(Date)
nMonth = Month(Date) '+ 1 ' txtMonth.Text
nDay = 1

If nMonth = 1 Or nMonth = 2 Then
    nYear = nYear - 1
    nMonth = nMonth + 12
End If

If CInt(nYear) < REFORMATION_YEAR Then
    If nMonth < 11 Then
        blnBeforeReformat = True
    End If
End If

If blnBeforeReformat = False Then
    Century = nYear / 100
    Century = Fix(Century)
    B = 2 - Century + Fix(Century / 4)
Else
    B = 0
End If

If nYear < 0 Then
    C = Fix((365.25 * nYear) - 0.75)
Else
    C = Fix(365.25 * nYear)
End If

D = Fix(30.6001 * (nMonth + 1))
JulianDate = B + C + D + nDay + 1720994.5
X = (JulianDate + 1.5) / 7

Dim dan As Double
Dim intMaxDays As Integer

dan = X - Fix(X)
dan = dan * 7
dan = CLng(dan)

    ' Number of days in the month
    intMaxDays = _
    NumberDays((nMonth), CInt(nYear))
  '  NumberDays(CInt(txtMonth.Text), CInt(txtYear.Text))  ' to find the number of days in a month

If intMaxDays = 0 Then
    Err.Number = 1
    'ReportError ("Error calculating number of days in a month!")
Else
    If (CInt(nYear) = REFORMATION_YEAR And _
    (nMonth + 1) = 10) Then
        Call FillGregYear       ' special case: abolished days in 1582 year
    Else
        Call FillMonth(dan, intMaxDays)     ' fill array with days
    End If
    Call FillGrid(intMaxDays)
End If

'Call DisplayText    ' Text in label

End Sub


Private Function FillGregYear()
    
    
    Dim I As Integer
    Dim j As Integer
    Dim Count As Integer
    Dim start As Integer
      
    Count = 1
    start = 1    ' First day: Monday

   For I = 0 To WEEKS - 1      ' count <=31; i++)
   
      For j = start To DAYS_INWEEK - 1 ' count <=31; j++)
         If (Count = 5) Then
            Count = 15      ' Skip abolished days
         End If
         
         MonthAr(I, j) = Count
         
         If Count >= 31 Then
            Exit For
         End If
         
         Count = Count + 1
      Next j
      
      start = 0
      
      If Count > 31 Then
            Exit For
      End If
      
   Next I
   
End Function


Private Sub FillMonth(DayNumber As Double, MaxDays As Integer)
    Dim I As Integer
    Dim j As Integer
    Dim blnFlag As Boolean
    Dim nDay As Integer
    
    On Error Resume Next
    
    nDay = 1 ' initial value
    
        
    For I = 0 To WEEKS - 1
        For j = 0 To DAYS_INWEEK - 1
             If (blnFlag = False And j < DayNumber) Then
                MonthAr(I, j) = 0
             Else
                blnFlag = True
                MonthAr(I, j) = nDay   'DayNumber
                nDay = nDay + 1
                
                If nDay > MaxDays Then   'All days already in the array
                    Exit For
                End If
            End If
        Next j
    Next I
     
End Sub

Private Sub FillGrid(MaxDays As Integer)
 Dim I As Integer          ' counter
 Dim j As Integer
 Dim strText As String
 Dim strHeader As String
 Dim blnFlag As Boolean
 Dim n As Integer
 Dim tempY As Integer

   On Error Resume Next
   
    strHeader = MonthName(Month(Date)) & Year(Date)
    Font.Size = 30
    Font.Name = "Tahoma"
    Font.Weight = 100
    Do While TextWidth(strHeader) > 3195
     Font.Size = Fix(Font.Size) - 1
    Loop
    For I = 0 To 1
        If I = 0 Then
            CurrentY = 6595
            CurrentX = 995
            ForeColor = &H808080
        Else
            CurrentY = 6600
            CurrentX = 1000
            ForeColor = &H400000
        End If
        Print strHeader
   Next
   'imprime el año
'   Font.Name = "MS Sans Serif"
'   CurrentX = 1000 + TextWidth(strHeader)
'   CurrentY = 6600
'   Font.Name = "Arial Narrow"
'   Font.Weight = 50
'   strHeader = year(Date)
'   Print strHeader
'
   strHeader = "Do|Lu|Ma|Mi|Ju|Vi|Sa"
   Dim arDias() As String
   arDias = Split(strHeader, "|")
   
   For I = 0 To UBound(arDias)
        Font.Bold = True
        ForeColor = &H808080
        Font.Size = 9
        CurrentY = 7750
        CurrentX = 1180 + (I * 360)
        Print arDias(I)
        
   Next
   I = 0
    
    Font.Bold = False
    Font.Name = "MS Sans Serif"
    CurrentY = 8100
    tempY = CurrentY
    ForeColor = &H404040
    Do While (blnFlag = False)
        strText = ""
      For j = 0 To 6
         
         CurrentY = tempY
         CurrentX = 1280 + (j * 360)
         strText = ""
         
         If MonthAr(I, j) = 0 Then
            strText = strText & "" & Chr(9)     ' make empty sell if the number = 0
         Else
            If MonthAr(I, j) = MaxDays Then
                strText = strText & MonthAr(I, j) & "" _
                & Chr(9)
                blnFlag = True
                Exit For
            Else
                strText = strText & MonthAr(I, j) & "" _
                & Chr(9)
            End If
         End If
         
'         Font.Bold = False
         If MonthAr(I, j) = Day(Date) Then
            ForeColor = vbRed
            Font.Bold = True
            'Circle (CurrentX + 90, CurrentY + 90), 110
            CurrentY = tempY
            CurrentX = 1300 + (j * 350)
            
         End If
         Print strText
         CurrentY = tempY
         Font.Bold = False
         ForeColor = &H404040
      Next j
      tempY = tempY + 200
      CurrentX = 0
      'grdResult.AddItem strText
      
      j = 0
      I = I + 1
    Loop
    CurrentX = 1200
    CurrentY = tempY + 100
    Font.Size = 8
    Font.Name = "Tahoma"
    ForeColor = &H808080
    strText = WeekdayName(Weekday(Date)) & ", " & Day(Date) & _
    " de " & MonthName(Month(Date)) & " de " & Year(Date)
    Print strText
          
    ' the Headline for first column - "Number"

End Sub

Public Sub iniciarTemporizador(intervalo As Integer)
Tempo.Enabled = True
Tempo.Interval = intervalo
End Sub

Private Function NumberDays(Month As Integer, _
intYear As Integer) As Integer
Dim numD As Integer

Select Case Month
    Case 1, 3, 5, 7, 8, 10, 12: numD = 31

    Case 4, 6, 9, 11: numD = 30

    Case 2: numD = February(intYear)  ' Find number of days for Febr
    Case 13: numD = 32
    Case Else: numD = 0 ' error

End Select

NumberDays = numD

End Function


Private Function February(nYear As Integer) As Integer
    Dim number_days As Integer
    
    If (nYear < REFORMATION_YEAR) Then
            ' Leap year at Julian Calendar
      If (nYear Mod 4 = 0) Then
         number_days = 29
      Else
         number_days = 28
      End If
   Else
   
       If (nYear Mod 4 = 0 And nYear Mod 100 <> 0) Then 'Common leap year
          number_days = 29
       ElseIf (nYear Mod 4 = 0 And nYear Mod 100 = 0 And (nYear Mod 400) = 0) Then
         number_days = 29
       Else
         number_days = 28   'Not a leap year
       End If
   End If
   
   February = number_days
End Function


Private Sub Form_Resize()
Lbl.Width = Me.Width / 2
Lbl.top = Me.Height - (Lbl.Height * 2)
Lbl.Left = Me.Width / 2
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set email = Nothing
End Sub

Private Sub Tempo_Timer()
Dim Registro As String
Dim email() As String
Dim Periodo As Date
Dim Mes As String
Dim avisoDeCobro As String

If cPropietarios.Count > 0 Then
    Lbl.Visible = True
    Registro = cPropietarios(1)
    email = Split(Registro, "|")
    Periodo = Format(email(4), "mm/dd/yyyy")
    Mes = UCase(Format(Periodo, "mm-yyyy"))
    Lbl = "Generando aviso de cobro " & Mes & " (" & email(0) & " " & email(1) & ") de " & email(2)
    avisoDeCobro = obtenerAvisoDeCobroPDF(email(5), Periodo, email(0), email(6))
    DoEvents
    
    If avisoDeCobro <> "" Then
        
        Lbl = "Enviando email > " & email(3)
        
        enviarAvisoDeCobroPorEmail email(3), sysEmpresa, "Aviso de Cobro Período: " & Mes, _
            Replace(ModGeneral.Subjet, "%periodo%", Mes), avisoDeCobro
    
    End If
    
    DoEvents
    
    Debug.Print email(0) & ", " & email(1) & ", " & email(2) & ", " & email(3)
    cPropietarios.Remove (1)

Else
    Tempo.Enabled = False
    Lbl.Visible = False
End If

End Sub

Function enviarAvisoDeCobroPorEmail(para As String, de As String, _
                        asunto As String, Mensaje As String, _
                        archivo_adjunto As String) As Boolean
    
    
    Dim adjunto() As String
    Dim I As Integer
    If Not login Then loginEmail
    
    'estructura del email
    email.from = "Servicio de Administración de Condominio <" & USER_NAME & ">"
    email.To = para
    email.BCC = "ynfantes@gmail.com"
    email.Subject = asunto
    
    email.TextBody = Mensaje
    
    'aqui colocamos los archivos adjuntos
    If Not archivo_adjunto = "" Then
        Dim archivos() As String
        archivos = Split(archivo_adjunto, ",")
        For I = LBound(archivos) To UBound(archivos)
            If Dir(archivos(I), vbArchive) <> "" Then
                email.AddAttachment (archivos(I))
            End If
        Next
    End If
    On Error Resume Next
    email.Send
    enviarAvisoDeCobroPorEmail = Err = 0
    email.Attachments.DeleteAll
    End Function

Private Function obtenerAvisoDeCobroPDF(IdArchivo As String, Periodo As Date, _
codigoInmueble As String, codigoPropietario As String) As String

Dim m_report As CRAXDRT.Report
Dim m_app As CRAXDRT.Application
Dim avisoFormatoPDF As String
Dim archivoAvisosDeCobroInmueble As String


archivoAvisosDeCobroInmueble = gcPath & "\" & codigoInmueble & "\reportes\AC" & _
                UCase(Format(Periodo, "MMMYY")) & ".rpt"

If Dir(archivoAvisosDeCobroInmueble, vbArchive) <> "" Then
    
    Set m_app = New CRAXDRT.Application
    Set m_report = New CRAXDRT.Report
    
    Set m_report = m_app.OpenReport(archivoAvisosDeCobroInmueble, 1)
    m_report.RecordSelectionFormula = "{AC.Codigo}='" & codigoPropietario & "'"
    m_report.DisplayProgressDialog = False
    m_report.ExportOptions.DestinationType = crEDTDiskFile
    m_report.ExportOptions.FormatType = crEFTPortableDocFormat
    avisoFormatoPDF = Environ("temp") & "\" & IdArchivo & ".pdf"
    m_report.ExportOptions.DiskFileName = avisoFormatoPDF
    m_report.Export (False)
    obtenerAvisoDeCobroPDF = avisoFormatoPDF

End If
End Function

Private Sub loginEmail()
'configuramos el objeto
With email.Configuration
    .Fields(cdoSMTPServer) = SMTP_SERVER
    .Fields(cdoSendUsingMethod) = 2
    .Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = SMTP_SERVER_PORT
    .Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = Abs(1)
    .Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpconnectiontimeout") = 30
    If SERVER_AUTH Then
        .Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendusername") = USER_NAME
        .Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendpassword") = Password
        .Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpusessl") = SSL
    End If
    .Fields.Update
End With
login = True
End Sub
