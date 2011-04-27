VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form FrmUsuarios 
   Caption         =   "Tablas del Sistema"
   ClientHeight    =   5775
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6600
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   5775
   ScaleWidth      =   6600
   Begin ComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   480
      Left            =   0
      TabIndex        =   8
      Top             =   0
      Width           =   6600
      _ExtentX        =   11642
      _ExtentY        =   847
      ButtonWidth     =   714
      ButtonHeight    =   688
      Appearance      =   1
      ImageList       =   "ImageList1"
      _Version        =   327682
      BeginProperty Buttons {0713E452-850A-101B-AFC0-4210102A8DA7} 
         NumButtons      =   12
         BeginProperty Button1 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "First"
            Object.ToolTipText     =   "Primer Registro"
            Object.Tag             =   ""
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Previous"
            Object.ToolTipText     =   "Registro Anterior"
            Object.Tag             =   ""
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Next"
            Object.ToolTipText     =   "Siguiente Registro"
            Object.Tag             =   ""
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "End"
            Object.ToolTipText     =   "Último Registro"
            Object.Tag             =   ""
            ImageIndex      =   4
         EndProperty
         BeginProperty Button5 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "New"
            Object.ToolTipText     =   "Nuevo Registro"
            Object.Tag             =   ""
            ImageIndex      =   5
         EndProperty
         BeginProperty Button6 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Save"
            Object.ToolTipText     =   "Guardar Registro"
            Object.Tag             =   ""
            ImageIndex      =   6
         EndProperty
         BeginProperty Button7 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Find"
            Object.ToolTipText     =   "Buscar Registro"
            Object.Tag             =   ""
            ImageIndex      =   7
         EndProperty
         BeginProperty Button8 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Undo"
            Object.ToolTipText     =   "Cancelar Registro"
            Object.Tag             =   ""
            ImageIndex      =   8
         EndProperty
         BeginProperty Button9 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Delete"
            Object.ToolTipText     =   "Eliminar Registro"
            Object.Tag             =   ""
            ImageIndex      =   9
         EndProperty
         BeginProperty Button10 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Edit1"
            Object.ToolTipText     =   "Editar Registro"
            Object.Tag             =   ""
            ImageIndex      =   10
         EndProperty
         BeginProperty Button11 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Print"
            Object.ToolTipText     =   "Imprimir"
            Object.Tag             =   ""
            ImageIndex      =   11
         EndProperty
         BeginProperty Button12 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Close"
            Object.ToolTipText     =   "Salir"
            Object.Tag             =   ""
            ImageIndex      =   12
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin VB.Frame Frame1 
      Height          =   5295
      Index           =   0
      Left            =   0
      TabIndex        =   6
      Top             =   480
      Width           =   6600
      Begin MSDataGridLib.DataGrid GridTablas 
         Bindings        =   "FrmUsuarios.frx":0000
         Height          =   3495
         Left            =   165
         TabIndex        =   7
         Top             =   240
         Width           =   6240
         _ExtentX        =   11007
         _ExtentY        =   6165
         _Version        =   393216
         AllowUpdate     =   0   'False
         AllowArrows     =   -1  'True
         BorderStyle     =   0
         Enabled         =   -1  'True
         HeadLines       =   1
         RowHeight       =   15
         RowDividerStyle =   4
         FormatLocked    =   -1  'True
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Usuarios Registrados"
         ColumnCount     =   4
         BeginProperty Column00 
            DataField       =   "NombreUsuario"
            Caption         =   "Usuario"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   "0,000E+00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   8202
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   "Contraseña"
            Caption         =   "Contraseña"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   8202
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column02 
            DataField       =   "Perfil"
            Caption         =   "Perfil"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   8202
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column03 
            DataField       =   "LogIn"
            Caption         =   "Conec."
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   5
               Format          =   ""
               HaveTrueFalseNull=   1
               TrueValue       =   "SI"
               FalseValue      =   "NO"
               NullValue       =   ""
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   8202
               SubFormatType   =   0
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            BeginProperty Column00 
               ColumnWidth     =   1890,142
            EndProperty
            BeginProperty Column01 
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   1425,26
            EndProperty
            BeginProperty Column03 
               Alignment       =   2
               ColumnWidth     =   675,213
            EndProperty
         EndProperty
      End
      Begin VB.Frame Frame1 
         Caption         =   "Aplicar Filtro:"
         Height          =   1215
         Index           =   1
         Left            =   3975
         TabIndex        =   9
         Top             =   3930
         Width           =   2550
         Begin VB.TextBox txtUsuario 
            Alignment       =   2  'Center
            Height          =   315
            Index           =   2
            Left            =   1380
            Locked          =   -1  'True
            TabIndex        =   15
            Top             =   240
            Width           =   840
         End
         Begin VB.OptionButton Opt 
            Caption         =   "Perfil"
            Height          =   285
            Index           =   4
            Left            =   1365
            TabIndex        =   14
            Top             =   855
            Visible         =   0   'False
            Width           =   1155
         End
         Begin VB.OptionButton Opt 
            Caption         =   "Activos"
            Height          =   285
            Index           =   3
            Left            =   1365
            TabIndex        =   13
            Top             =   570
            Width           =   960
         End
         Begin VB.OptionButton Opt 
            Caption         =   "Conectados"
            Height          =   285
            Index           =   2
            Left            =   90
            TabIndex        =   12
            Top             =   855
            Width           =   1215
         End
         Begin VB.OptionButton Opt 
            Caption         =   "Inactivos"
            Height          =   285
            Index           =   1
            Left            =   90
            TabIndex        =   11
            Top             =   570
            Width           =   1215
         End
         Begin VB.OptionButton Opt 
            Caption         =   "Ninguno"
            Height          =   285
            Index           =   0
            Left            =   90
            TabIndex        =   10
            Top             =   285
            Value           =   -1  'True
            Width           =   1215
         End
      End
      Begin MSAdodcLib.Adodc AdoTablas 
         Height          =   345
         Index           =   0
         Left            =   150
         Top             =   1320
         Visible         =   0   'False
         Width           =   2325
         _ExtentX        =   4101
         _ExtentY        =   609
         ConnectMode     =   3
         CursorLocation  =   3
         IsolationLevel  =   -1
         ConnectionTimeout=   15
         CommandTimeout  =   30
         CursorType      =   1
         LockType        =   3
         CommandType     =   1
         CursorOptions   =   0
         CacheSize       =   50
         MaxRecords      =   0
         BOFAction       =   0
         EOFAction       =   0
         ConnectStringType=   1
         Appearance      =   1
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Orientation     =   0
         Enabled         =   -1
         Connect         =   ""
         OLEDBString     =   ""
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   ""
         Caption         =   "AdoTablas"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _Version        =   393216
      End
      Begin VB.TextBox txtUsuario 
         Height          =   315
         Index           =   0
         Left            =   1290
         Locked          =   -1  'True
         TabIndex        =   1
         Top             =   4020
         Width           =   2535
      End
      Begin VB.TextBox txtUsuario 
         Height          =   315
         IMEMode         =   3  'DISABLE
         Index           =   1
         Left            =   1275
         Locked          =   -1  'True
         MaxLength       =   7
         PasswordChar    =   "*"
         TabIndex        =   3
         Top             =   4425
         Width           =   2535
      End
      Begin MSDataListLib.DataCombo dtcUsuario 
         Height          =   315
         Left            =   1275
         TabIndex        =   5
         Top             =   4830
         Width           =   2550
         _ExtentX        =   4498
         _ExtentY        =   556
         _Version        =   393216
         Locked          =   -1  'True
         Style           =   2
         ListField       =   "Perfil"
         Text            =   ""
      End
      Begin MSAdodcLib.Adodc AdoTablas 
         Height          =   345
         Index           =   1
         Left            =   2985
         Top             =   1185
         Visible         =   0   'False
         Width           =   2325
         _ExtentX        =   4101
         _ExtentY        =   609
         ConnectMode     =   3
         CursorLocation  =   3
         IsolationLevel  =   -1
         ConnectionTimeout=   15
         CommandTimeout  =   30
         CursorType      =   1
         LockType        =   3
         CommandType     =   1
         CursorOptions   =   0
         CacheSize       =   50
         MaxRecords      =   0
         BOFAction       =   0
         EOFAction       =   0
         ConnectStringType=   1
         Appearance      =   1
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Orientation     =   0
         Enabled         =   -1
         Connect         =   ""
         OLEDBString     =   ""
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   ""
         Caption         =   "AdoTablas"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _Version        =   393216
      End
      Begin VB.Label lblUsuario 
         Caption         =   "Contraseña:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   1
         Left            =   210
         TabIndex        =   2
         Top             =   4470
         Width           =   1200
      End
      Begin VB.Label lblUsuario 
         Caption         =   "Usuario:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   0
         Left            =   210
         TabIndex        =   0
         Top             =   4035
         Width           =   885
      End
      Begin VB.Label lblUsuario 
         AutoSize        =   -1  'True
         Caption         =   "Perfil:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   2
         Left            =   210
         TabIndex        =   4
         Top             =   4845
         Width           =   885
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
            Picture         =   "FrmUsuarios.frx":0018
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmUsuarios.frx":019A
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmUsuarios.frx":031C
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmUsuarios.frx":049E
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmUsuarios.frx":0620
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmUsuarios.frx":07A2
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmUsuarios.frx":0924
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmUsuarios.frx":0AA6
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmUsuarios.frx":0C28
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmUsuarios.frx":0DAA
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmUsuarios.frx":0F2C
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmUsuarios.frx":10AE
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "FrmUsuarios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
    Dim mlEdit As Boolean, mlNew As Boolean
    Dim cnnTabla As New ADODB.Connection
    

    Private Sub AdoTablas_MoveComplete(Index As Integer, _
    ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, _
    adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
    '
    If Index = 0 Then txtUsuario(2) = AdoTablas(0).Recordset.RecordCount
    '
    End Sub


    Private Sub Form_Load()
    CenterForm Me
    
    AdoTablas(0).ConnectionString = cnnOLEDB + gcPath + "\tablas.mdb"
    AdoTablas(1).ConnectionString = cnnOLEDB + gcPath + "\tablas.mdb"
    
    If gcNivel >= nuAdministrador Then
    '
        GridTablas.Columns(1).Visible = False
        '
        AdoTablas(0).RecordSource = "SELECT Usuarios.NombreUsuario, Usuarios.Contraseña, Nivele" _
        & "s.Perfil,Usuarios.LogIN,Usuarios.Nivel FROM Niveles INNER JOIN Usuarios ON Niveles.N" _
        & "ivel = Usuarios.Nivel WHERE Niveles.Nivel >" & nuADSYS & " ORDER BY Usuarios.NombreU" _
        & "suario;"
        AdoTablas(1).RecordSource = "SELECT * FROM Niveles WHERE Nivel>0;"
        
    Else
        AdoTablas(0).RecordSource = "SELECT Usuarios.NombreUsuario, Usuarios.Contraseña, Nivele" _
        & "s.Perfil,Usuarios.LogIn,Usuarios.Nivel FROM Niveles INNER JOIN Usuarios ON Niveles.N" _
        & "ivel = Usuarios.Nivel ORDER BY Usuarios.NombreUsuario;"
        AdoTablas(1).RecordSource = "SELECT * FROM Niveles"
        txtUsuario(1).Locked = False
        '
    End If
    '
    AdoTablas(0).Refresh
    AdoTablas(1).Refresh
    Set dtcUsuario.RowSource = AdoTablas(1)
    Set GridTablas.DataSource = AdoTablas(0)
    Set GridTablas.HeadFont = LetraTitulo(LoadResString(527), 7.5, True)
    Set GridTablas.Font = LetraTitulo(LoadResString(528), 8)
    cnnTabla.ConnectionString = AdoTablas(0).ConnectionString
    cnnTabla.Open
    '
    End Sub

    Private Sub Form_Unload(Cancel As Integer)
    cnnTabla.Close
    Set cnnTabla = Nothing
    End Sub

    Private Sub Opt_Click(Index As Integer)
    'variables locales
    Dim strFiltro As String
    Select Case Index
        Case 0
            strFiltro = ""
            Set GridTablas.DataSource = Nothing
        Case 1: strFiltro = "Nivel =" & nuINACTIVO
        Case 2: strFiltro = "LogIn =True"
        Case 3: strFiltro = "Nivel <>" & nuINACTIVO & " And Nivel <> " & nuADSYS
        Case 4: strFiltro = ""
    End Select
    AdoTablas(0).Recordset.Filter = strFiltro
    If Index = 0 Then Set GridTablas.DataSource = AdoTablas(0)
    '
    End Sub

    '---------------------------------------------------------------------------------------------
    Private Sub Toolbar1_ButtonClick(ByVal Button As ComctlLib.Button)  '
    '---------------------------------------------------------------------------------------------
    '
    With AdoTablas(0).Recordset
    '
        Select Case UCase(Button.Key)
        '
            Case "FIRST"    'Ir al Primer registro
            '
                If Not .EOF Or Not .BOF Then .MoveFirst
                '
            Case "PREVIOUS" 'Ir al registro anterior
            '
                If Not .EOF Or Not .BOF Then
                    .MovePrevious
                    If .BOF Then .MoveLast
                End If
                '
            Case "NEXT" 'Avanzar al siguiente registro
            '
                If Not .EOF Or Not .BOF Then
                    .MoveNext
                    If .EOF Then .MoveFirst
                End If
                '
            Case "END"  'Ir al último registro
            '
                If Not .EOF Or .BOF Then .MoveLast
                '
            Case "NEW"
            '
                Call RtnEstado(Button.Index, Toolbar1)
                txtUsuario(0).Locked = False
                txtUsuario(1) = "SAC"
                dtcUsuario.Locked = False
                txtUsuario(0) = ""
                dtcUsuario = ""
                txtUsuario(0).SetFocus
                mlNew = True
                mlEdit = Not mlNew
                '
            Case "SAVE"
            '
                On Error Resume Next
                txtUsuario(0).Locked = True
                txtUsuario(1).Locked = True
                dtcUsuario.Locked = True
                Dim strSQL, strBitacora As String
                If mlNew Then
                    strSQL = "INSERT INTO Usuarios(NombreUsuario,Contraseña,Nivel) VALUES('" _
                    & txtUsuario(0) & "','DEFAULT'," & Nivel & ");"
                    strBitacora = "Nuevo Reg.Usuarios: " & txtUsuario(0) & "; " & dtcUsuario
                ElseIf mlEdit Then
                    strSQL = "UPDATE Usuarios SET NombreUsuario='" & txtUsuario(0) & "',Nivel=" _
                    & Nivel & " WHERE NombreUsuario='" & !NombreUsuario & "';"
                    strBitacora = "Actualizar Reg.Usuarios: " & !NombreUsuario & "; " & dtcUsuario
                End If
                cnnTabla.Execute strSQL
'                If Nivel < 2 Then
'                    Dim ctlMenu As Control
'                    Dim strIndice As String
'                    For Each ctlMenu In FrmAdmin.Controls
'                        If TypeOf ctlMenu Is Menu Then
'                            strIndice = "(" & ctlMenu.Index & ")"
'                            If Err.Number = 343 Then strIndice = "": Err.Clear
'                            cnnTabla.Execute "INSERT INTO Perfiles (Usuario,Acceso) VALU" _
'                            & "ES('" & txtUsuario(0) & "','" & ctlMenu.Name & strIndice _
'                            & "');"
'                        End If
'                    Next
'                End If
                If Err.Number = 0 Then
                    Call rtnBitacora(strBitacora)
                    Call RtnEstado(Button.Index, Toolbar1)
                    MsgBox "Registro Actualizado...", vbInformation, App.ProductName
                    AdoTablas(0).Refresh
                    GridTablas.ReBind
                Else
                    Call rtnBitacora(strBitacora & "[FALlIDA]")
                    MsgBox "Operación Fallida..." & Err.Description, vbExclamation, _
                    App.ProductName
                End If
                '
            Case "FIND"
                
            Case "UNDO"
                Call RtnEstado(Button.Index, Toolbar1)
                txtUsuario(0).Locked = True: txtUsuario(0) = ""
                txtUsuario(1).Locked = True: txtUsuario(1) = ""
                dtcUsuario.Locked = True: dtcUsuario = ""
                '
            Case "DELETE"   'Eliminar Usuario
                Dim strMensaje$
                strMensaje = "¿Esta seguro que desea eliminar al usuario " & !NombreUsuario & "?"
                If Respuesta(strMensaje) Then
                    Call RtnEstado(Button.Index, Toolbar1)
                    cnnTabla.Execute "DELETE FROM Usuarios WHERE NombreUsuario='" _
                    & !NombreUsuario & "';"
                    Call rtnBitacora("Eliminar Usuario " & !NombreUsuario)
                   MsgBox " Registro Eliminado ... ", vbInformation, App.ProductName
                    AdoTablas(0).Refresh
                    GridTablas.ReBind
                End If
            
            Case "EDIT1"
            '
            mlEdit = True
            mlNew = Not mlEdit
            txtUsuario(0) = !NombreUsuario
            txtUsuario(1) = !Contraseña
            dtcUsuario = !Perfil
            txtUsuario(0).Locked = False
            dtcUsuario.Locked = False
            Call RtnEstado(Button.Index, Toolbar1)
            
            Case "PRINT"    'Imprimir Listado de Usuarios
            '
            Case "CLOSE"   'Cierra el formulario
            '
                Unload Me
                Set FrmUsuarios = Nothing
            
        End Select
        '
    End With
    '
    End Sub

    '---------------------------------------------------------------------------------------------
    Private Function Nivel() As Long    '
    '---------------------------------------------------------------------------------------------
    '
    With AdoTablas(1).Recordset
        .MoveFirst
        .Find "Perfil='" & dtcUsuario & "'"
        Nivel = !Nivel
    End With
    '
    End Function


    '---------------------------------------------------------------------------------------------
    Private Sub TxtUsuario_KeyPress(Index As Integer, KeyAscii As Integer)  '
    '---------------------------------------------------------------------------------------------
    KeyAscii = Asc(UCase(Chr(KeyAscii)))    'convierte a mayúsculas
    If Index = 0 And txtUsuario(0).Locked Then
        MsgBox "Presione el botón 'NUEVO' de la barra de herramientas para agregar un usuario", _
        vbInformation, "Usuarios y Perfiles"
    End If
    If KeyAscii = 13 Then
    '
        Select Case Index
        '
            Case 0  'Nombre Usuario
                txtUsuario(1).SetFocus
            Case 1
                dtcUsuario.SetFocus
                '
        End Select
        '
    End If
    '
    End Sub
