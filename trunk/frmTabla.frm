VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form FrmTabla 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tablas del Sistema"
   ClientHeight    =   5760
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   5460
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmTabla.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   5760
   ScaleWidth      =   5460
   Begin ComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   480
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   5460
      _ExtentX        =   9631
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
            Enabled         =   0   'False
            Object.Visible         =   0   'False
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
      MousePointer    =   99
      MouseIcon       =   "frmTabla.frx":000C
   End
   Begin VB.TextBox txtTabla 
      DataField       =   "Nombre"
      Height          =   330
      Left            =   1545
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   5025
      Width           =   3540
   End
   Begin MSAdodcLib.Adodc AdoTablas 
      Align           =   2  'Align Bottom
      Height          =   330
      Left            =   0
      Top             =   5430
      Visible         =   0   'False
      Width           =   5460
      _ExtentX        =   9631
      _ExtentY        =   582
      ConnectMode     =   16
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   1
      LockType        =   3
      CommandType     =   2
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
   Begin VB.Frame Frame1 
      Height          =   5265
      Left            =   15
      TabIndex        =   0
      Top             =   525
      Width           =   5400
      Begin MSDataGridLib.DataGrid GridTablas 
         Bindings        =   "frmTabla.frx":0326
         Height          =   4095
         Left            =   435
         TabIndex        =   1
         Top             =   240
         Width           =   4575
         _ExtentX        =   8070
         _ExtentY        =   7223
         _Version        =   393216
         AllowUpdate     =   0   'False
         BorderStyle     =   0
         Enabled         =   -1  'True
         HeadLines       =   1
         RowHeight       =   15
         FormatLocked    =   -1  'True
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
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
         Caption         =   "Tablas"
         ColumnCount     =   1
         BeginProperty Column00 
            DataField       =   "Nombre"
            Caption         =   "Descripcion"
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
         SplitCount      =   1
         BeginProperty Split0 
            BeginProperty Column00 
               ColumnWidth     =   2924,788
            EndProperty
         EndProperty
      End
      Begin VB.Label Label1 
         Caption         =   "Descripción:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   390
         TabIndex        =   4
         Top             =   4605
         Width           =   1515
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
               Picture         =   "frmTabla.frx":033E
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frmTabla.frx":04C0
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frmTabla.frx":0642
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frmTabla.frx":07C4
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frmTabla.frx":0946
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frmTabla.frx":0AC8
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frmTabla.frx":0C4A
               Key             =   ""
            EndProperty
            BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frmTabla.frx":0DCC
               Key             =   ""
            EndProperty
            BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frmTabla.frx":0F4E
               Key             =   ""
            EndProperty
            BeginProperty ListImage10 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frmTabla.frx":10D0
               Key             =   ""
            EndProperty
            BeginProperty ListImage11 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frmTabla.frx":1252
               Key             =   ""
            EndProperty
            BeginProperty ListImage12 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frmTabla.frx":13D4
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
End
Attribute VB_Name = "FrmTabla"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
    Dim mlEdit As Boolean
    
   
    Private Sub Form_Load()
    
    With AdoTablas
        .ConnectionString = cnnOLEDB + gcPath + "\Tablas.mdb"
        .RecordSource = mcTablas
        .Refresh
    End With
    CenterForm Me
    Set txtTabla.DataSource = AdoTablas
    GridTablas.Caption = mcTitulo
    Me.Caption = mcTitulo
    End Sub

    '---------------------------------------------------------------------------------------------
    '   Rutina: Toolbar1_ButtonClick
    '
    '   Procedimientos correspondientes a los eventos ButtonClick de la
    '   barra de herramientas
    '---------------------------------------------------------------------------------------------
    Private Sub Toolbar1_ButtonClick(ByVal Button As ComctlLib.Button)
    '
    With AdoTablas.Recordset
    
        Select Case UCase(Button.Key)
        '
            Case "FIRST"    'Primer Registro
            '-------------------------
                If Not .EOF Or Not .BOF Then .MoveFirst
                '
            Case "PREVIOUS" 'Registro Previo
            '-------------------------
                If Not .EOF Or Not .BOF Then .MovePrevious
                If .BOF Then .MoveLast
                '
            Case "NEXT" 'Registro siguiente
            '-------------------------
                If Not .EOF Or Not .BOF Then .MoveNext
                If .EOF Then .MoveFirst
                '
            Case "END"  'Ultimo registro
            '-------------------------
                If Not .EOF Or Not .BOF Then .MoveLast
                '
            Case "NEW"  'Nuevo Registro
            '-------------------------
                .AddNew
                txtTabla.Locked = False
                Call RtnEstado(Button.Index, Toolbar1)
            Case "SAVE" 'Guardar Regristro
            '-------------------------
                Call RtnEstado(Button.Index, Toolbar1)
                If Valida_guardar(txtTabla) Then Exit Sub
                If mcTablas = "Bancos" Then
                    If IsNull(!CodBanco) Then !CodBanco = ID
                    mlEdit = True
                End If
                !Usuario = gcUsuario
                !fecha = Date
                !Hora = Time()
                .Update
                txtTabla.Locked = True
                GridTablas.Rebind
                '
            Case "UNDO" 'Cancelar Regristro
            '-------------------------
                .CancelUpdate
                txtTabla = ""
                GridTablas.Refresh
                Call RtnEstado(Button.Index, Toolbar1)
                MsgBox "Operación Cancelada..."
                '
            Case "DELETE"   'Eliminar Registro
            '-------------------------
                
                If Respuesta(LoadResString(526)) Then
                    .Delete
                    mlEdit = True
                    .MoveNext
                    If .EOF Then .MoveFirst
                    Call RtnEstado(Button.Index, Toolbar1)
                    MsgBox " Registro Eliminado ... "
                End If
                        
            Case "PRINT"    'Imprimir
            '-------------------------
            Case "CLOSE"    'Descargar Formulario
            '-------------------------
                If mcTablas = "Bancos" And mlEdit Then
                    cnnConexion.Execute "DELETE FROM Bancos;"
                    cnnConexion.Execute "INSERT INTO BANCOS SELECT Bancos.codBanco, Bancos.Nomb" _
                    & "re FROM Bancos IN '" & gcPath & "\tablas.mdb';"
                End If
                Unload Me
                Set FrmTabla = Nothing
            '
        End Select
    '
    End With
    '
    End Sub


    Private Sub txtTabla_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    End Sub
    '---------------------------------------------------------------------------------------------
    '   Funcion:    ID
    '
    '   Devuelde un número único para la clave de la tabla bancos
    '---------------------------------------------------------------------------------------------
    Private Function ID() As Integer
    '
    Dim rstTabla As New ADODB.Recordset   'Variables locales
    Dim int1 As Integer
    '
    rstTabla.Open "SELECT * FROM " & mcTablas, AdoTablas.ConnectionString, _
    adOpenStatic, adLockOptimistic
    With rstTabla
        .MoveFirst
        int1 = !CodBanco
        .MoveNext
        Do Until .EOF
            If int1 + 1 <> !CodBanco Then
                ID = !CodBanco
                GoTo 10
            End If
            int1 = !CodBanco
            .MoveNext
        Loop
        ID = int1 + 1
10         .Close
    End With
    Set rstTabla = Nothing
    '
    End Function
    
    
    Private Function Valida_guardar(strNom) As Boolean
    '
    Dim rstValida As New ADODB.Recordset
    rstValida.Open AdoTablas.RecordSource, AdoTablas.ConnectionString, adOpenStatic, _
    adLockReadOnly
    rstValida.MoveFirst
    rstValida.Find "Nombre='" & strNom & "'"
    If Not rstValida.EOF Then
        Valida_guardar = MsgBox(strNom & " ya se encuetra registrado en la " & mcTitulo, _
        vbCritical)
    End If
    rstValida.Close: Set rstValida = Nothing
    '
    End Function
