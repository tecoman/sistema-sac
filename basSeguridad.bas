Attribute VB_Name = "basSeguridad"
    '---------------------------------------------------------------------------------------------
    '   Sinai Tech- Módulo de seguridad y auditoria del sistema
    '   Creado: 01/11/2002
    '
    '   Entradas:   Nombre Usuario y Contraseña
    '   Salidas:   True si usuario registrado, sino False
    '               Archivo *.log del seguimiento del usuario
    '---------------------------------------------------------------------------------------------
    Option Explicit
    '
    Private Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" _
    (ByVal lpbuffer As String, nSize As Long) As Long
    Private Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" ( _
    ByVal lpbuffer As String, nSize As Long) As Long
    Public IntTaquilla%             'variable que guarda la taquilla abierta por el usuario
    Public strFCaja$                'Fecha de apertura de la taquilla
    Public CurSaldoOpen@            'Guarda el monto de apertura de Caja
    Public blnCaja As Boolean       'Bandera formulario de Caja
    Public gcNivel As NivelUsuario  'Nivel de Seguridad Usuario
    Public gcMAC$                   'Nombre Hardware en el servidor
    Public gcContraseña$            'Almacena la contraseña del usuario
    
    '---------------------------------------------------------------------------------------------
    '   FUNCTION:   ftnSegur
    '
    '   DEVUELVE:   True si el usuario cumple con los requerimientos
    '               de seguridad, sino false
    '
    '   ENTRADAS:   Nombre Usuario, Contraseña
    '
    '   SALIDA:     Archivo *.Log (Bitacora o Auditoria del sistema)
    '               Formato(fecha,ddmmaa).log
    '               Ejem:011102.log archivo generado el 01/11/2002
    '---------------------------------------------------------------------------------------------
    Public Function ftnSegur(strUID$, strPWD$, Txt1, Txt2 As TextBox, strIP As String) As Boolean
    '---------------------------------------------------------------------------------------------
    'variables locales
    Dim cnnSeg As New ADODB.Connection    'conexion al origen de los registros
    Dim rstSeg As New ADODB.Recordset     'Conjunto de registros solicitados
    Dim rstPerfil As ADODB.Recordset
    Dim ctlPerfil As Object
    Dim intPA%, intPC%, intInd%
    Dim booCajero As Boolean
    Dim sBuffer$, lSize&, hKey&
    '
    On Error GoTo SalirSistema:
    'Obtiene el nombre de la maquina
    sBuffer = Space$(255)
    lSize = Len(sBuffer)
    Call GetComputerName(sBuffer, lSize)
    If lSize > 0 Then
        gcMAC = Left$(sBuffer, lSize)
    Else
        gcMAC = vbNullString
    End If
    '
    cnnSeg.Open cnnOLEDB & gcPath & "\tablas.mdb" ';Jet OLEDB:Database Password=" & _
    strPSWD(gcPath & "\tablas.mdb")
    rstSeg.Open "SELECT * FROM Usuarios WHERE NombreUsuario ='" & strUID & "'", cnnSeg, _
    adOpenDynamic, adLockPessimistic, adCmdText
    '
    If Not rstSeg.EOF Then
        
        If strPWD = rstSeg!Contraseña Then  'verifica la contraseña
        
            If rstSeg!Nivel = nuINACTIVO Then
                ftnSegur = MsgBox("Su Login está inactivo. Póngase en contacto con el administra" _
                & "dor del sistema.", vbInformation, App.ProductName)
                Call rtnBitacora("Login Inactivo. " & strUID)
            Else
                gcUsuario = strUID
                gcNivel = rstSeg!Nivel
                gcContraseña = strPWD
                gcNombreCompleto = IIf(IsNull(rstSeg!NombreCompleto), "", rstSeg!NombreCompleto)
                'If Not rstSeg!Login Then
                    Set rstPerfil = New ADODB.Recordset
                    rstSeg.Update "Login", True
                    rstSeg.Update "IP", strIP
                    rstSeg.Update "Maquina", gcMAC
                    Call rtnBitacora("Log In, Iniciando sesión SAC...")
                    
                    If gcNivel > nuADSYS Then
                        rstPerfil.Open "SELECT * FROM Perfiles WHERE Usuario ='" & strUID & "'", cnnSeg, _
                        adOpenKeyset, adLockOptimistic, adCmdText
                        If rstPerfil.EOF Or rstPerfil.BOF Then
                            ftnSegur = MsgBox("Consulte al supervisor, No Tiene Ningún Acceso al Sistem" _
                            & "a...", vbInformation, App.ProductName)
                            rstPerfil.Close
                            Set rstPerfil = Nothing
                            cnnSeg.Close
                            Set cnnSeg = Nothing
                            Exit Function
                        End If
                        rstPerfil.MoveFirst
                        On Error Resume Next
                        
                        Do Until rstPerfil.EOF
                            If rstPerfil!Acceso = "AC400(0)" Then booCajero = True
                            intPA = InStr(rstPerfil!Acceso, "(")
                            If Not intPA = 0 Then
                                intPC = InStr(rstPerfil!Acceso, ")")
                                intInd = Mid(rstPerfil!Acceso, intPA + 1, intPC - intPA - 1)
                                FrmAdmin.Controls(Left(rstPerfil!Acceso, intPA - 1)).Item(intInd).Enabled = True
                            Else
                                FrmAdmin.Controls(rstPerfil!Acceso).Enabled = True
                                
                            End If
                            rstPerfil.MoveNext
                        Loop
                        rstPerfil.Close ': Set rstPerfil = Nothing
                    End If
                    'coloca la hora de entrada al trabajo
                    If gcNivel > nuADSYS Then       '   si no es el administrador del sistema ó
                                                    '   no es administrador de la empresa
                        rstPerfil.Open "SELECT * FROM Emp_Asistencia WHERE Fecha=Date()", _
                        cnnConexion, adOpenKeyset, adCmdText
                        '
                        If rstPerfil.EOF And rstPerfil.BOF Then
                            'registras todos los usuarios activos
                            cnnConexion.Execute "INSERT INTO Emp_Asistencia (Usuario,Fecha,Entrada," _
                            & "Salida) SELECT NombreUsuario,Date(),'12:00:00 a.m.','12:00:00 a.m.' " _
                            & "FROM Usuarios IN '" & gcPath & "\Tablas.mdb' WHERE Nivel=2 or Nivel=3;"
                        End If
                        rstPerfil.Close
                        '
                        rstPerfil.Open "SELECT * FROM Emp_asistencia WHERE Usuario='" & gcUsuario & _
                        "' AND Fecha=Date() AND Entrada<>#12/30/1899#", cnnConexion, adOpenKeyset, _
                        adLockOptimistic, adCmdText
                        
                        If rstPerfil.EOF Or rstPerfil.BOF Then
                            'actualiza la hora de entrada del usuario
                            cnnConexion.Execute "UPDATE Emp_Asistencia SET Entrada=Time(),Salida=Ti" _
                            & "me() WHERE Usuario='" & gcUsuario & "' AND Fecha=Date()"
                            Call rtnBitacora("Hora de entrada Ingresada")
                        End If
                        rstPerfil.Close
                        Set rstPerfil = Nothing
                        
                    End If
                    If booCajero = True Then Call rtnCajero(cnnSeg)
    '            Else
    '                ftnSegur = MsgBox("Usuario: " & gcUsuario & " ya está conectado al sistema," & vbCrLf _
    '                & "Consulte al administrador del sistema", vbExclamation, App.EXEName)
    '                Call rtnBitacora(CadenaLog("Log In, sesión ya iniciada..."))
                'End If
            End If
        Else
            ftnSegur = MsgBox("Contraseña Invalida, verifique su contraseña e intentelo nuevame" _
            & "nte", vbCritical, App.ProductName)
            Call rtnBitacora("Log In , contraseña invalida")
            Call rtnFoco(Txt2)
        End If
    Else
        ftnSegur = MsgBox("Usuario '" & strUID & "' no está registrado, verifique que el nombre" _
        & vbCrLf & "este escrito correctamente y vuelva a intertarlo.", vbOKOnly + vbCritical, _
        App.ProductName)
        Call rtnBitacora("User no Registrado....")
        Call rtnFoco(Txt1)
    End If
    '
    rstSeg.Close
    Set rstSeg = Nothing
    cnnSeg.Close
    Set cnnSeg = Nothing
SalirSistema:
    If Hex(Err.Number) = 80004005 Then
        '
        MsgBox "Consulte al proveedor del sistema. Tome nota del error siguiente: '" & _
        Hex(Err.Number) & "' Formato DB no es correcto", vbCritical, "Error " & Hex(Err.Number)
        End
        '
    ElseIf Err.Number <> 0 Then
        '
        Call rtnBitacora("Error " & Err.Number & " " & Right(Err.Description, 50))
'        MsgBox "Ha ocurrido el siguiente error al tratar de iniciar el sistema:" & vbCrLf & _
'        Err.Description & vbCrLf & "Consulte al administrador del sistema", vbCritical, _
'        App.ProductName
'        End
        '
    End If
    '
    End Function
    
    '---------------------------------------------------------------------------------------------
    Private Sub rtnFoco(ByVal txt As TextBox)
    '---------------------------------------------------------------------------------------------
    '
    With txt
        .SelStart = 0
        .SelLength = Len(txt.Text)
        .SetFocus
    End With
    '
    End Sub
    
    '---------------------------------------------------------------------------------------------
    Private Sub rtnCajero(cnnCajero As ADODB.Connection) '
    '---------------------------------------------------------------------------------------------
    'Variables locales
    Dim rstCajero As New ADODB.Recordset
    Dim rstTaquilla As New ADODB.Recordset
    '
    rstCajero.Open "SELECT * FROM Taquillas WHERE USUARIO='" & gcUsuario & "'", cnnCajero, _
    adOpenKeyset, adLockOptimistic, adCmdText
    '
    If Not rstCajero.EOF Then
    '
        IntTaquilla = rstCajero!IDTaquilla
        '
        rstTaquilla.Open "SELECT * FROM Taquillas WHERE IDTaquilla=" & IntTaquilla, _
        cnnConexion, adOpenKeyset, adLockOptimistic, adCmdText
        CurSaldoOpen = rstTaquilla!OpenSaldo
    '
        If rstTaquilla!Estado Then
    '
            FrmAdmin.AC400(0).Checked = True
            strFCaja = CStr(rstTaquilla!fecha)
            If rstTaquilla!Cuadre Then
                blnCaja = True
           End If
    '
            If rstTaquilla!fecha < Date Then
                MsgBox "La Taquilla N° " & IntTaquilla & " permanece abierta...", vbCritical, _
                App.ProductName
                Call rtnBitacora("Caja Abierta desde " & rstTaquilla!fecha)
            End If
            Call rtnEstadoCaja("True")
        Else
            Call rtnEstadoCaja("False")
        End If
    '
    Else
        '
        MsgBox "Usuario: " & gcUsuario & " Ud. NO TIENE CAJA ASIGNADA", vbInformation, _
        App.ProductName
        '
    End If
    '
    rstCajero.Close
    Set rstCajero = Nothing
    '
    End Sub
    
    '---------------------------------------------------------------------------------------------
    Public Sub rtnBitacora(strCadena$) '
    '---------------------------------------------------------------------------------------------
    ''variables locales
    Dim numFichero%, strArchivo$
    '------------------------
    numFichero = FreeFile
    strArchivo = "\Bitacora\" & CStr(Format(Day(Date), "00") & Month(Date) _
    & Right(Year(Date), 2) & ".txt")
    Open gcPath & strArchivo For Append As numFichero
    Write #numFichero, Time(), gcMAC, gcUsuario, strCadena
    Close numFichero
    '-------------------------
    End Sub
    


    
