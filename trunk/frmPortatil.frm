VERSION 5.00
Begin VB.Form frmPortatil 
   AutoRedraw      =   -1  'True
   Caption         =   "SAC Portátil"
   ClientHeight    =   2415
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6765
   Icon            =   "frmPortatil.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   2415
   ScaleWidth      =   6765
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox pic1 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00800000&
      Height          =   330
      Left            =   240
      ScaleHeight     =   270
      ScaleWidth      =   6150
      TabIndex        =   10
      Top             =   285
      Visible         =   0   'False
      Width           =   6210
   End
   Begin VB.PictureBox pic 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00800000&
      Height          =   330
      Left            =   240
      ScaleHeight     =   270
      ScaleWidth      =   6150
      TabIndex        =   9
      Top             =   990
      Width           =   6210
   End
   Begin VB.CommandButton cmb 
      Caption         =   "&Cerrar"
      Height          =   375
      Index           =   2
      Left            =   5265
      TabIndex        =   2
      Top             =   1455
      Width           =   1215
   End
   Begin VB.CommandButton cmb 
      Height          =   285
      Index           =   1
      Left            =   4605
      Picture         =   "frmPortatil.frx":27A2
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   1935
      Width           =   240
   End
   Begin VB.TextBox txt 
      Height          =   315
      Left            =   1005
      TabIndex        =   1
      Top             =   1920
      Width           =   3600
   End
   Begin VB.CommandButton cmb 
      Caption         =   "&Inciar"
      Height          =   375
      Index           =   0
      Left            =   5265
      TabIndex        =   3
      Tag             =   "0"
      Top             =   1890
      Width           =   1215
   End
   Begin VB.Label lbl 
      Height          =   210
      Index           =   4
      Left            =   270
      TabIndex        =   8
      Top             =   750
      Width           =   3885
   End
   Begin VB.Label lbl 
      Height          =   315
      Index           =   3
      Left            =   1020
      TabIndex        =   6
      Top             =   1500
      Width           =   3885
   End
   Begin VB.Label lbl 
      Caption         =   "&Hasta:"
      Height          =   315
      Index           =   2
      Left            =   270
      TabIndex        =   0
      Top             =   1935
      Width           =   630
   End
   Begin VB.Label lbl 
      Caption         =   "Desde:"
      Height          =   315
      Index           =   1
      Left            =   270
      TabIndex        =   5
      Top             =   1500
      Width           =   630
   End
   Begin VB.Label lbl 
      Caption         =   $"frmPortatil.frx":28EC
      Height          =   540
      Index           =   0
      Left            =   285
      TabIndex        =   4
      Top             =   210
      Width           =   6300
   End
End
Attribute VB_Name = "frmPortatil"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Const MAX_PATH = 255

Private Enum eBIF
    BIF_RETURNONLYFSDIRS = &H1            'Sólo directorios del sistema
    BIF_DONTGOBELOWDOMAIN = &H2           'No incluir carpetas de red
    BIF_STATUSTEXT = &H4
    BIF_RETURNFSANCESTORS = &H8
    BIF_BROWSEFORCOMPUTER = &H1000        'Buscar PCs
    BIF_BROWSEFORPRINTER = &H2000         'Buscar impresoras
End Enum

Private Type BrowseInfo
    hwndOwner               As Long
    pIDLRoot                As Long             'Especifica dónde se empezará a mostrar
    pszDisplayName          As Long
    lpszTitle               As Long
    ulFlags                 As Long
    lpfnCallback            As Long
    lParam                  As Long
    iImage                  As Long
End Type

Private Declare Function SHBrowseForFolder Lib "shell32.dll" _
        (lpbi As BrowseInfo) As Long

Private Declare Sub CoTaskMemFree Lib "ole32.dll" _
        (ByVal hMem As Long)

Private Declare Function lstrcat Lib "kernel32.dll" Alias "lstrcatA" _
        (ByVal lpString1 As String, ByVal lpString2 As String) As Long

Private Declare Function SHGetPathFromIDList Lib "shell32.dll" _
        (ByVal pidList As Long, ByVal lpbuffer As String) As Long

Dim booDetener As Boolean


'Si se quiere usar en un form, cambiar el public por private
Private Function BrowseForFolder(ByVal hwndOwner As Long, ByVal sPrompt As String, Optional ByVal vFlags As eBIF) As String
    '
    Dim iNull As Integer
    Dim lpIDList As Long
    Dim lResult As Long
    Dim sPath As String
    Dim udtBI As BrowseInfo
    Dim lFlags As Long

    If Not IsMissing(vFlags) Then
        lFlags = CInt(vFlags)
    End If

    With udtBI
        .hwndOwner = hwndOwner
        .lpszTitle = lstrcat(sPrompt, "")
        .ulFlags = lFlags Or BIF_RETURNONLYFSDIRS
    End With

    lpIDList = SHBrowseForFolder(udtBI)
    If lpIDList Then
        sPath = String$(MAX_PATH, 0)
        lResult = SHGetPathFromIDList(lpIDList, sPath)
        Call CoTaskMemFree(lpIDList)
        iNull = InStr(sPath, vbNullChar)
        If iNull Then
            sPath = Left$(sPath, iNull - 1)
        End If
    Else
        'Se ha pulsado en cancelar
        sPath = ""
    End If

    BrowseForFolder = sPath
End Function


Private Sub cmb_Click(Index%)
Dim Direc, Carpe, archivos, Arc, errLocal, Carpes, Car, Car1, Carpes1
Dim strOrigen$, strDestino$, strMaquina$, strLinea$, strCarpera$, strClose$
Dim numArchivo As Integer
Dim booConec As Boolean
Dim I&, j&, Maxi& 'contadores
Dim INI, INI1 As Date
If Index = 0 Then

    If Cmb(0).Tag = 1 Then
        Fin = InputBox$("Introduzca la clave de cierre por favor." & vbCrLf & _
        "Si no la tiene consulte con el administrador del sistema", _
        "Confirme la finalización del proceso")
        If UCase(Fin) = "FIN" Then
            Cmb(0).Tag = 0
            booDetener = True
            Cmb(2).Enabled = True
            Lbl(4) = ""
            pic.Visible = False
            Lbl(0).Visible = True
        Else
            SetTimer hWnd, NV_CLOSEMSGBOX, 4000&, AddressOf TimerProc
            Call MessageBox(hWnd, "SAC continuará con este proceso", _
            App.ProductName, vbInformation)
        End If
    Else
        
        If Txt = "" Then
        
            MsgBox "Seleccione la carpeta de destino", vbCritical, App.ProductName
            Exit Sub
            
        End If
        'COMPRUBa la existencia del directorio
        If Dir(Txt & "\", vbDirectory) = "" Then
            MsgBox "Debe crear el directorio de destino antes de continuar", vbCritical, _
            App.ProductName
            Exit Sub
        End If
        If Dir(App.Path & "\sacportatil.log") <> "" Then Kill App.Path & "\sacportatil.log"
        pic.Visible = True
        pic1.Visible = True
        Cmb(2).Enabled = False
        Cmb(0).Caption = "Detener"
        Cmb(0).Tag = 1
        Lbl(0).Visible = False
        strOrigen = gcPath
        Lbl(3) = strOrigen
        strDestino = Txt
        
        On Error GoTo Salir:
        Set Direc = CreateObject("Scripting.FileSystemObject")
        Set Carpe = Direc.GetFolder(strOrigen)
        Set archivos = Carpe.Files
        Maxi = archivos.Count
        I = 1
        'Inicialmente copia todos los archivos dentro del directorio raiz
        UpdateStatus pic1, 1, True
        Call rtnBitacora("Inicianco el proceso de copiado...")
        rtnBitacora ("Fecha: " & Format(Date, "dd-mm-yyyy"))
        rtnBitacora ("Desde: " & strOrigen & vbTab & " Hasta: " & strDestino)
        INI1 = Time()
        rtnBitacora ("Hora Inicio: " & Format(INI, "hh:mm:ss"))
        For Each Arc In archivos
            UpdateStatus pic, I * 100 / Maxi, True
            If Not Arc.Name Like "*.ldb" Then
                Lbl(3) = "'" & Arc.Path & "'"
                Lbl(4) = "Copiando....'" & Arc.Name & "'"
                Me.rtnBitacora (Arc.Name)
                DoEvents
                If booDetener Then
                    Cmb(0).Caption = "Iniciar"
                    pic1.Visible = False
                    Exit Sub
                End If
                Direc.CopyFile Arc.Path, strDestino & "\" & Arc.Name, True
                If UCase(Arc.Name) = "SAC.MDB" Then
                    'actualiza el campo  ruta de la tabla ambiente en el equipo remoto
                    Dim cnn As New ADODB.Connection
                    cnn.Open cnnOLEDB & strDestino & "\sac.mdb"
                    cnn.Execute "UPDATE Ambiente SET Ruta = '" & Txt & "'"
                    cnn.Close
                    Set cnn = Nothing
                End If
            End If
            I = I + 1
        Next
        'ahora copia los archivos contenidos en cada carpeta dentro del
        'directorio raíz
        Set Carpes = Carpe.SubFolders
        j = 1
        I = 1
        Me.rtnBitacora ("Respaldando directorios...")
        For Each Car In Carpes
            If UCase(Carpe.Name) <> "NOMINA" Then
            Maxi = Carpes.Count
            strOrigen = Car.Path
            Set Carpe = Direc.GetFolder(strOrigen)
            Set archivos = Carpe.Files
            UpdateStatus pic, j * 100 / Maxi, True
            UpdateStatus pic1, j * 100 / Maxi, True
            Maxi = archivos.Count
            INI = Time()
            'Me.rtnBitacora (Car.Name & " -- " & Format(Time, "hh:mm:ss"))
            For Each Arc In archivos
                'pgb.Max = Archivos.Count
                UpdateStatus pic, I * 100 / Maxi, True
                If Not Arc.Name Like "*.ldb" Then
                    'crea el directorio si no existe
                    If Dir(strDestino & "\" & Car.Name & "\", vbDirectory) = "" Then
                        Direc.CreateFolder (strDestino & "\" & Car.Name)
                    End If
                        Lbl(3) = "'" & IIf(Len(Arc.Path) > 30, Left(Arc.Path, 24) & String(10, ".") & Arc.Name, Arc.Path) & "'"
                        Lbl(4) = "Copiando....'" & Car.Name & "\" & Arc.Name & "'"
                        Lbl(3).Refresh
                        Lbl(4).Refresh
                        'Me.Refresh
                        DoEvents
                        If booDetener Then
                            Cmb(0).Caption = "Iniciar"
                            pic1.Visible = False
                            Exit Sub
                        End If
                        Direc.CopyFile Arc.Path, strDestino & "\" & Car.Name & "\" & _
                        Arc.Name, True
                                
                End If
                I = I + 1
            Next
            
            I = 1
            j = j + 1
            
            Set Carpes1 = Carpe.SubFolders
            For Each Car1 In Carpes1
                'copia tod el contenido de la carpeta
                Set archivos = Car1.Files
                Maxi = archivos.Count
                If Maxi > 0 Then
                    If Dir$(strDestino & "\" & Carpe.Name & "\" & Car1.Name, vbDirectory) = "" Then Direc.CreateFolder (strDestino & "\" & Carpe.Name & "\" & Car1.Name)
                End If
                For Each Arc In archivos
                    UpdateStatus pic, I * 100 / Maxi, True
                    Lbl(3) = "'" & IIf(Len(Arc.Path) > 30, Left(Arc.Path, 24) & String(10, ".") & "\" & Arc.Name, Arc.Path) & "'"
                    Lbl(4) = "Copiando....'" & Car.Name & "\" & Car1.Name & "\" & Arc.Name & "'"
                    DoEvents
                    If booDetener Then
                        Cmb(0).Caption = "Iniciar"
                        pic1.Visible = False
                        Exit Sub
                    End If
                    Direc.CopyFile Arc.Path, strDestino & "\" & Car.Name & "\" & Car1.Name & "\" & _
                    Arc.Name, True
                    I = I + 1
                Next
            Next
            Me.rtnBitacora (Car.Name & " Inicia: " & Format(INI, "hh:mm:ss")) & " Fin: " & Format(Time, "hh:mm:ss") & " Duración: " & Format(Time() - INI, "hh:mm:ss")
            End If
        Next
        
Salir:
        Cmb(0).Tag = 0
        Cmb(0).Caption = "&Iniciar"
        Cmb(2).Enabled = True
        booDetener = False
        Lbl(0).Visible = True
        If Err.Number = 70 Then
            MsgBox "El archivo " & Arc.Name & " está siendo utilizado por otro usuario. Verifique q" _
            & "ue todos los usuarios hayan finalizado la sesión en el sistema e intentelo nuevament" _
            & "e", vbInformation, App.ProductName
        ElseIf Err.Number <> 0 Then
            MsgBox Err.Description, vbCritical, App.ProductName
        Else
            rtnBitacora ("Proceso Finalizado " & Format(Time, "hh:mm:ss") & vbTab & "Duración:" & Format(Time() - INI1, "hh:mm:ss"))
            Lbl(4) = "Proceso completado con Éxito!!"
            pic.Visible = False
            pic1.Visible = False
        End If
        
        
    End If
ElseIf Index = 1 Then
    Txt = BrowseForFolder(Me.hWnd, "Selecciona un directorio")
Else
    Unload Me
End If
'
End Sub

Private Sub Form_Load()
Lbl(3) = gcPath
CenterForm Me
pic.Visible = False
If Dir(App.Path & "\sacportatil.log") <> "" Then
    If Respuesta("Existe un registro del último respaldo" & vbCrLf & _
    "¿Desea abir este registro?") Then
        frmBitacora.rtxBitacora.FileName = App.Path & "\sacportatil.log"
    Else
        Kill App.Path & "\sacportatil.log"
    End If
End If
booDetener = False
End Sub

Private Function LoadResCustom(ID, Optional Tipo) As Variant
Dim bytes() As Byte, idf As Integer
    
idf = FreeFile
Open App.Path & "ressac.tmp" For Binary As #idf
bytes = LoadResData(ID, Tipo)
Put #idf, , bytes
Close #idf
LoadResCustom = App.Path & "\ressac1.tmp"
End Function


Public Sub rtnBitacora(strCadena$) '
Dim numFichero%, strArchivo$
'------------------------
numFichero = FreeFile
strArchivo = App.Path & "\sacportatil.log"
Open strArchivo For Append As numFichero
Print #numFichero, strCadena
Close numFichero
'-------------------------
End Sub

