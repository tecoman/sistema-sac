VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmDiario 
   AutoRedraw      =   -1  'True
   Caption         =   ".:::Diario:::."
   ClientHeight    =   6540
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8205
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   6540
   ScaleWidth      =   8205
   WindowState     =   2  'Maximized
   Begin ComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   480
      Left            =   0
      TabIndex        =   14
      Top             =   0
      Width           =   8205
      _ExtentX        =   14473
      _ExtentY        =   847
      ButtonWidth     =   714
      ButtonHeight    =   688
      AllowCustomize  =   0   'False
      Appearance      =   1
      ImageList       =   "ImageList1"
      _Version        =   327682
      BeginProperty Buttons {0713E452-850A-101B-AFC0-4210102A8DA7} 
         NumButtons      =   12
         BeginProperty Button1 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.Visible         =   0   'False
            Key             =   "FIRST"
            Object.ToolTipText     =   "Primer Registro"
            Object.Tag             =   ""
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.Visible         =   0   'False
            Key             =   "BACK"
            Object.ToolTipText     =   "Registro Anterior"
            Object.Tag             =   ""
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.Visible         =   0   'False
            Key             =   "NEXT"
            Object.ToolTipText     =   "Siguiente Registro"
            Object.Tag             =   ""
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.Visible         =   0   'False
            Key             =   "LAST"
            Object.ToolTipText     =   "Último Registro"
            Object.Tag             =   ""
            ImageIndex      =   4
         EndProperty
         BeginProperty Button5 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "NEW"
            Object.ToolTipText     =   "Nuevo Registro"
            Object.Tag             =   ""
            ImageIndex      =   5
         EndProperty
         BeginProperty Button6 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Enabled         =   0   'False
            Key             =   "SAVE"
            Object.ToolTipText     =   "Guardar Registro"
            Object.Tag             =   ""
            ImageIndex      =   6
         EndProperty
         BeginProperty Button7 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "FIND"
            Object.ToolTipText     =   "Buscar Registro"
            Object.Tag             =   ""
            ImageIndex      =   7
         EndProperty
         BeginProperty Button8 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Enabled         =   0   'False
            Key             =   "CANCEL"
            Object.ToolTipText     =   "Cancelar Registro"
            Object.Tag             =   ""
            ImageIndex      =   8
         EndProperty
         BeginProperty Button9 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Enabled         =   0   'False
            Key             =   "DELETE"
            Object.ToolTipText     =   "Eliminar Registro"
            Object.Tag             =   ""
            ImageIndex      =   9
         EndProperty
         BeginProperty Button10 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Enabled         =   0   'False
            Object.Visible         =   0   'False
            Key             =   "EDIT"
            Object.ToolTipText     =   "Editar Registro"
            Object.Tag             =   ""
            ImageIndex      =   10
         EndProperty
         BeginProperty Button11 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "PRINT"
            Object.ToolTipText     =   "Imprimir"
            Object.Tag             =   ""
            ImageIndex      =   11
         EndProperty
         BeginProperty Button12 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "CLOSE"
            Object.ToolTipText     =   "Salir"
            Object.Tag             =   ""
            ImageIndex      =   12
         EndProperty
      EndProperty
      BorderStyle     =   1
      MousePointer    =   99
      MouseIcon       =   "frmDiario.frx":0000
   End
   Begin MSComctlLib.ImageList iml 
      Left            =   1185
      Top             =   165
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
            Picture         =   "frmDiario.frx":031A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDiario.frx":076C
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   5415
      Left            =   135
      TabIndex        =   8
      Top             =   795
      Width           =   7920
      _ExtentX        =   13970
      _ExtentY        =   9551
      _Version        =   393216
      Tabs            =   2
      Tab             =   1
      TabsPerRow      =   2
      TabHeight       =   520
      ShowFocusRect   =   0   'False
      TabCaption(0)   =   "Diario"
      TabPicture(0)   =   "frmDiario.frx":0BBE
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "fra"
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Lista"
      TabPicture(1)   =   "frmDiario.frx":0BDA
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "shp"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "lbl(4)"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "lst"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "rtx(1)"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).ControlCount=   4
      Begin VB.Frame fra 
         Enabled         =   0   'False
         Height          =   4695
         Left            =   -74880
         TabIndex        =   9
         Top             =   570
         Width           =   7680
         Begin VB.TextBox txt 
            Height          =   315
            Index           =   1
            Left            =   1230
            TabIndex        =   7
            Top             =   1200
            Width           =   6060
         End
         Begin VB.TextBox txt 
            Height          =   315
            Index           =   0
            Left            =   4080
            TabIndex        =   3
            Top             =   315
            Width           =   1215
         End
         Begin RichTextLib.RichTextBox rtx 
            Height          =   2790
            Index           =   0
            Left            =   150
            TabIndex        =   10
            Top             =   1785
            Width           =   7365
            _ExtentX        =   12991
            _ExtentY        =   4921
            _Version        =   393217
            TextRTF         =   $"frmDiario.frx":0BF6
         End
         Begin MSDataListLib.DataCombo DataCombo1 
            Height          =   315
            Index           =   0
            Left            =   1230
            TabIndex        =   1
            Top             =   330
            Width           =   1500
            _ExtentX        =   2646
            _ExtentY        =   556
            _Version        =   393216
            Style           =   2
            ListField       =   "CodInm"
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo DataCombo1 
            Height          =   315
            Index           =   1
            Left            =   1245
            TabIndex        =   5
            Top             =   765
            Width           =   1500
            _ExtentX        =   2646
            _ExtentY        =   556
            _Version        =   393216
            Style           =   2
            ListField       =   "NombreUsuario"
            Text            =   ""
         End
         Begin VB.Label lbl 
            Caption         =   "Atención:"
            Height          =   285
            Index           =   3
            Left            =   225
            TabIndex        =   4
            Top             =   780
            Width           =   870
         End
         Begin VB.Label lbl 
            Caption         =   "Asunto:"
            Height          =   285
            Index           =   2
            Left            =   225
            TabIndex        =   6
            Top             =   1260
            Width           =   870
         End
         Begin VB.Label lbl 
            Caption         =   "Remitente:"
            Height          =   285
            Index           =   1
            Left            =   3045
            TabIndex        =   2
            Top             =   345
            Width           =   990
         End
         Begin VB.Label lbl 
            Caption         =   "Condominio:"
            Height          =   285
            Index           =   0
            Left            =   225
            TabIndex        =   0
            Top             =   375
            Width           =   990
         End
      End
      Begin RichTextLib.RichTextBox rtx 
         Height          =   2070
         Index           =   1
         Left            =   165
         TabIndex        =   11
         Top             =   3165
         Width           =   7440
         _ExtentX        =   13123
         _ExtentY        =   3651
         _Version        =   393217
         Enabled         =   -1  'True
         ReadOnly        =   -1  'True
         TextRTF         =   $"frmDiario.frx":0C78
      End
      Begin MSComctlLib.ListView lst 
         Height          =   1935
         Left            =   180
         TabIndex        =   13
         Top             =   945
         Width           =   7440
         _ExtentX        =   13123
         _ExtentY        =   3413
         View            =   3
         LabelEdit       =   1
         SortOrder       =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         SmallIcons      =   "iml"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   7
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Object.Tag             =   "14"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Object.Tag             =   "7"
            Text            =   "Inmueble"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Object.Tag             =   "12"
            Text            =   "Remite"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Object.Tag             =   "10"
            Text            =   "Para"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Object.Tag             =   "26"
            Text            =   "Asunto"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Object.Tag             =   "18"
            Text            =   "Fecha"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Object.Tag             =   "9"
            Text            =   "Usuario"
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.Label lbl 
         BackColor       =   &H8000000C&
         Caption         =   ".::::LISTADO DIARIO:::."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   4
         Left            =   225
         TabIndex        =   12
         Top             =   585
         Width           =   7350
      End
      Begin VB.Shape shp 
         BackColor       =   &H8000000C&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H8000000C&
         Height          =   495
         Left            =   180
         Top             =   450
         Width           =   7440
      End
   End
   Begin ComctlLib.ImageList ImageList1 
      Left            =   4785
      Top             =   450
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
            Picture         =   "frmDiario.frx":0CFA
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmDiario.frx":0E7C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmDiario.frx":0FFE
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmDiario.frx":1180
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmDiario.frx":1302
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmDiario.frx":161C
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmDiario.frx":179E
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmDiario.frx":1920
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmDiario.frx":1AA2
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmDiario.frx":1C24
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmDiario.frx":1DA6
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmDiario.frx":1F28
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmDiario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rst(2) As New ADODB.Recordset

Private Sub Form_Load()
'variables locales
Dim strSql(3) As String
Dim i As Integer
'
'selecciona todos los inmuebles activos
'selecciona todos los usuarios activos del sistema
strSql(0) = "SELECT * FROM Inmueble WHERE Inactivo=False ORDER BY CodInm"
strSql(1) = "SELECT * FROM Usuarios WHERE Nivel >0 and Nivel <4 ORDER BY NombreUsuario"
strSql(2) = "SELECT * FROM Diario"
strSql(3) = cnnOLEDB & gcPath & "\tablas.mdb"

For i = 0 To 2
    rst(i).Open strSql(i), IIf(i = 1, strSql(3), cnnConexion), adOpenKeyset, adLockReadOnly, _
    adCmdText
    If i <> 2 Then Set DataCombo1(i).RowSource = rst(i)
Next
'
Call Agregar_list("Fecha DESC, Hora DESC")
If Not rst(2).EOF And Not rst(2).BOF Then Toolbar1.Buttons("DELETE").Enabled = True
'Me.Show

End Sub

Private Sub Form_Resize()
'variables locales
Dim Ficha As Integer
Dim ancho As Long
Dim i%
'
On Error Resume Next
If WindowState = vbMaximized Then
'
With SSTab1
    '
    .Tab = 1
    .Left = ScaleLeft + 100
    .Width = ScaleWidth - (.Left * 2)
    .Height = ScaleHeight - .Top - 200

    lbl(4).Width = .Width - (lbl(4).Left * 2)
    lst.Width = .Width - (lst.Left * 2)
    lst.Height = .Height * 0.6
    shp.Width = lst.Width
    rtx(1).Top = lst.Top + lst.Height
    rtx(1).Width = lst.Width
    rtx(1).Height = .Height - rtx(1).Top - 200
    '
    'configura el ancho de las columnas
    ancho = lst.Width
    For i = 1 To 8
        lst.ColumnHeaders(i).Width = lst.ColumnHeaders(i).Tag * ancho / 100
    Next
    
    .Tab = 0
    fra.Height = .Height - .Top
    fra.Width = .Width - (fra.Left * 2)
    rtx(0).Width = fra.Width - (rtx(0).Left * 2)
    rtx(0).Height = fra.Height - rtx(0).Top - 200
    txt(0).Width = rtx(0).Width - txt(0).Left
    'txt(1).Width = rtx(0).Width - txt(1).Left
    '
    
    .Tab = 1
    
End With
'
End If

End Sub

Private Sub Form_Unload(Cancel As Integer)
'
Dim i%
On Error Resume Next
For i = 0 To 2
    rst(i).Close
    Set rst(i) = Nothing
Next
End Sub



Sub Agregar_list(Optional orden As String)
'variables locales
Dim X As Long

lst.ListItems.Clear
rtx(1).FileName = ""
lst.Sorted = False
'
With rst(2)
    .Requery
    If orden <> "" Then rst(2).Sort = orden
    
    If Not .EOF And Not .BOF Then
        .MoveFirst: X = 1
        Do
            'añade el elemento a la lista
            lst.ListItems.Add , , !msg, , IIf(!Leido, 2, 1)
            lst.ListItems(X).ListSubItems.Add , , Trim(!CodInm)
            lst.ListItems(X).ListSubItems.Add , , Trim(!Remite)
            lst.ListItems(X).ListSubItems.Add , , Trim(!Para)
            lst.ListItems(X).ListSubItems.Add , , Trim(!Asunto)
            lst.ListItems(X).ListSubItems.Add , , Trim(!fecha) & " " & Trim(!Hora)
            'lst.ListItems(X).ListSubItems.Add , ,
            lst.ListItems(X).ListSubItems.Add , , Trim(!Usuario)
            .MoveNext: X = X + 1
            
        Loop Until .EOF
        .MoveFirst
    End If
    
End With

End Sub


Function Error() As Boolean
If Not DataCombo1(0).MatchedWithList Then
    Error = MsgBox("Código de inmueble inválido", vbExclamation, App.ProductName)
End If
If txt(0) = "" Then
    Error = MsgBox("Falta Remitente", vbExclamation, App.ProductName)
End If
If txt(1) = "" Then
    Error = MsgBox("Falta Asunto", vbExclamation, App.ProductName)
End If
If Not DataCombo1(1).MatchedWithList Then
    Error = MsgBox("Destinatario no válido", vbExclamation, App.ProductName)
End If
If rtx(0).Text = "" Then
    Error = MsgBox("Falta el cuerpo del mensaje", vbExclamation, App.ProductName)
End If
End Function

Private Sub lst_Click()
'variables locales
If lst.ListItems.Count > 0 Then
'
    If lst.SelectedItem.SmallIcon = 1 Then
        lst.SelectedItem.SmallIcon = 2
        cnnConexion.Execute "UPDATE Diario SET Leido =True WHERE Msg = '" & _
        Trim(lst.SelectedItem) & "';"
    End If
    '
End If
'
End Sub

Private Sub lst_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
'variables locales
Dim ASCDES As String

If ColumnHeader = "Fecha" Then
    If lst.SortOrder = lvwAscending Then
        Call Agregar_list("Fecha DESC, Hora DESC")
    Else
        Call Agregar_list("Fecha ASC, Hora ASC")
    End If
Else
    lst.Sorted = True
    lst.SortKey = ColumnHeader.Index - 1
End If
lst.SortOrder = IIf(lst.SortOrder = lvwAscending, lvwDescending, lvwAscending)
'
End Sub

Private Sub lst_ItemClick(ByVal Item As MSComctlLib.ListItem)
'
If lst.ListItems.Count > 0 Then
    '
    If Dir(gcPath & "\msg\" & lst.SelectedItem & ".msg") <> "" Then
            rtx(1).FileName = gcPath & "\msg\" & lst.SelectedItem & ".msg"
    Else
        rtx(1).FileName = ""
    End If
    '
End If
'
End Sub

Private Sub lst_KeyUp(KeyCode As Integer, Shift As Integer)
Dim boton As ComctlLib.Button

If KeyCode = 46 And Shift = 0 Then
    Set boton = Toolbar1.Buttons("DELETE")
    Call Toolbar1_ButtonClick(boton)
    lst.SetFocus
End If
End Sub

Private Sub rtx_KeyPress(Index%, KeyAscii%): KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
'If SSTab1.Tab = 1 Then lst.Sorted = True
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As ComctlLib.Button)
'variables locales
Dim Nmsg As String
Dim Resp As Long
Dim rstMsg As ADODB.Recordset
Dim strSql As String
'
Select Case Button.Key
    
    
    Case "NEW"
        SSTab1.Tab = 0
        SSTab1.TabEnabled(1) = False
        fra.Enabled = True
        DataCombo1(0) = ""
        DataCombo1(1) = ""
        txt(0) = ""
        txt(1) = ""
        rtx(0) = ""
        DataCombo1(0).SetFocus
        Call RtnEstado(Button.Index, Toolbar1)
        
    Case "CANCEL"
        SSTab1.TabEnabled(1) = True
        DataCombo1(0) = ""
        DataCombo1(1) = ""
        txt(0) = ""
        txt(1) = ""
        rtx(0) = ""
        fra.Enabled = False
        Call RtnEstado(Button.Index, Toolbar1)
        
    Case "SAVE"
        
        If Not Error Then
        
            Nmsg = DataCombo1(0) & Format(Date, "ddmmyy") & Format(Time, "hhmmss")
            'guarda el archivo *.msg
            rtx(0).SaveFile gcPath & "\msg\" & Nmsg & ".msg", 1
            'guarda el registro
            cnnConexion.Execute "INSERT INTO Diario(CodInm,Remite,Para,Asunto,Fecha,Hora,Usuari" _
            & "o,Msg) VALUES ('" & DataCombo1(0) & "','" & txt(0) & "','" & DataCombo1(1) & "'," _
            & "'" & txt(1) & "',Date(),Time(),'" & gcUsuario & "','" & Nmsg & "')"
            SSTab1.TabEnabled(1) = True
            fra.Enabled = False
            Call Agregar_list
            Call RtnEstado(Button.Index, Toolbar1)
            MsgBox "Mensaje guardado", vbInformation, App.ProductName
            
        End If
    
    Case "DELETE"
    
        If lst.SelectedItem = "" Then
            MsgBox "Debe seleccionar un elemento de la lista", vbExclamation, App.ProductName
        Else
            '
            If gcNivel <= nuAdministrador Then
                strSql = "SELECT * FROM Diario WHERE MSG ='" & Trim(lst.SelectedItem) & "'"
            Else
                strSql = "SELECT * FROM Diario WHERE Usuario='" & gcUsuario & "' AND MSG ='" & _
                Trim(lst.SelectedItem) & "'"
            End If
            'crea una instancia del objeto ADODB.Recordset
            Set rstMsg = New ADODB.Recordset
            'selecciona el conjunto de registros
            rstMsg.Open strSql, cnnConexion, adOpenKeyset, adLockOptimistic, adCmdText
            '
            If Not rstMsg.EOF And Not rstMsg.BOF Then
            
                Resp = MsgBox("Seguro de eliminar el mensaje '" & Trim(lst.SelectedItem) & "'", _
                vbYesNo + vbQuestion + vbDefaultButton2, App.ProductName)
                If Resp = vbYes Then
                    If Dir(gcPath & "\msg\" & rstMsg("Msg") & ".msg") <> "" Then
                        Kill gcPath & "\msg\" & rstMsg("Msg") & ".msg"
                    End If
                    rstMsg.Delete
                    Call Agregar_list
                    MsgBox "Mensaje eliminado", vbInformation, App.ProductName
                End If
                
            Else
                MsgBox "Usted no puede eliminar este mensaje", vbInformation, App.ProductName
            End If
            rstMsg.Close
            Set rstMsg = Nothing
            '
            '
        End If
        
    Case "PRINT"
        MsgBox "Opción no disponible", vbInformation, App.ProductName
        
    Case "CLOSE"
        Unload Me
        
End Select

End Sub


Private Sub txt_KeyPress(Index%, KeyAscii%): KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub
