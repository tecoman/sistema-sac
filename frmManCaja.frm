VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmManCaja 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Utilidades Caja"
   ClientHeight    =   6390
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   8445
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6390
   ScaleWidth      =   8445
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmd 
      Caption         =   "&Salir"
      Height          =   900
      Index           =   2
      Left            =   6915
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   5220
      Width           =   1200
   End
   Begin VB.Frame fraManCaja 
      Caption         =   "Reasignar Transacciones"
      Height          =   3090
      Index           =   1
      Left            =   4395
      TabIndex        =   5
      Top             =   150
      Width           =   3705
      Begin VB.CommandButton cmd 
         Caption         =   "&Aceptar"
         Height          =   600
         Index           =   0
         Left            =   300
         TabIndex        =   12
         Top             =   2160
         Width           =   3060
      End
      Begin VB.TextBox txt 
         Height          =   315
         Index           =   2
         Left            =   2055
         MaxLength       =   1
         TabIndex        =   9
         Top             =   1575
         Width           =   525
      End
      Begin VB.TextBox txt 
         Height          =   315
         Index           =   1
         Left            =   2040
         MaxLength       =   1
         TabIndex        =   8
         Top             =   1140
         Width           =   525
      End
      Begin VB.TextBox txt 
         Height          =   315
         Index           =   0
         Left            =   2010
         Locked          =   -1  'True
         TabIndex        =   6
         Text            =   "Text1"
         Top             =   525
         Width           =   1215
      End
      Begin VB.Label lbl 
         Caption         =   "Hasta la taquilla:"
         Height          =   255
         Index           =   4
         Left            =   540
         TabIndex        =   11
         Top             =   1575
         Width           =   1275
      End
      Begin VB.Label lbl 
         Caption         =   "Desde la taquilla:"
         Height          =   255
         Index           =   3
         Left            =   540
         TabIndex        =   10
         Top             =   1200
         Width           =   1275
      End
      Begin VB.Label lbl 
         Caption         =   "Operaciones de fecha:"
         Height          =   300
         Index           =   2
         Left            =   240
         TabIndex        =   7
         Top             =   555
         Width           =   1740
      End
   End
   Begin VB.Frame fraManCaja 
      Caption         =   "Asignar Taquilla: "
      Height          =   5925
      Index           =   0
      Left            =   315
      TabIndex        =   0
      Top             =   165
      Width           =   3900
      Begin VB.CommandButton cmd 
         Height          =   390
         Index           =   3
         Left            =   465
         MaskColor       =   &H8000000F&
         Picture         =   "frmManCaja.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   15
         ToolTipText     =   "Asignar  taquilla "
         Top             =   5400
         Width           =   390
      End
      Begin VB.CommandButton cmd 
         Height          =   390
         Index           =   1
         Left            =   855
         Picture         =   "frmManCaja.frx":0172
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Elimina la taquilla seleccionada"
         Top             =   5400
         Width           =   390
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid grid 
         Height          =   2145
         Index           =   0
         Left            =   480
         TabIndex        =   1
         Tag             =   "2500"
         Top             =   720
         Width           =   2970
         _ExtentX        =   5239
         _ExtentY        =   3784
         _Version        =   393216
         Cols            =   1
         FixedCols       =   0
         BackColorFixed  =   -2147483646
         ForeColorFixed  =   -2147483639
         BackColorSel    =   65280
         ForeColorSel    =   0
         BackColorBkg    =   -2147483636
         GridColor       =   -2147483636
         GridColorFixed  =   -2147483636
         GridLinesFixed  =   1
         BorderStyle     =   0
         FormatString    =   "Nombre Usuario"
         _NumberOfBands  =   1
         _Band(0).Cols   =   1
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid grid 
         Height          =   2025
         Index           =   1
         Left            =   480
         TabIndex        =   2
         Tag             =   "600|2000"
         Top             =   3315
         Width           =   2970
         _ExtentX        =   5239
         _ExtentY        =   3572
         _Version        =   393216
         FixedCols       =   0
         BackColorFixed  =   -2147483646
         ForeColorFixed  =   -2147483639
         BackColorSel    =   65280
         ForeColorSel    =   0
         BackColorBkg    =   -2147483636
         GridColor       =   -2147483636
         GridColorFixed  =   -2147483636
         GridLinesFixed  =   1
         BorderStyle     =   0
         FormatString    =   "Taq.|Cajero"
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin VB.Label lbl 
         Alignment       =   2  'Center
         BackColor       =   &H80000002&
         Caption         =   "CAJAS ASIGNADAS"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000009&
         Height          =   240
         Index           =   1
         Left            =   465
         TabIndex        =   4
         Top             =   3060
         Width           =   2970
      End
      Begin VB.Label lbl 
         Alignment       =   2  'Center
         BackColor       =   &H80000002&
         Caption         =   "USUARIOS REGISTRADOS"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000009&
         Height          =   225
         Index           =   0
         Left            =   465
         TabIndex        =   3
         Top             =   405
         Width           =   2970
      End
   End
End
Attribute VB_Name = "frmManCaja"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

    Private Sub Cmd_Click(Index As Integer)
    'variables locales
    Dim Reg As Long
    Dim rstCaja As ADODB.Recordset
    '-----
    Select Case Index
        
        Case 0   'reasignar
            
            If txt(1) = "" Or txt(2) = "" Then
                MsgBox "Faltan parámtros para llevar a cabo esta petición", vbInformation, _
                App.ProductName
                Exit Sub
            ElseIf Not IsNumeric(txt(1)) Or Not IsNumeric(txt(2)) Then
                MsgBox "Introdujo Números de Caja no válidos", vbExclamation, App.ProductName
                Exit Sub
            End If
            On Error Resume Next
            cnnConexion.BeginTrans
            '
            'actualiza el movimiento de la caja
            cnnConexion.Execute "UPDATE MovimientoCaja SET IDTaquilla=" & txt(2) & " WHERE Fech" _
            & "aMovimientoCaja=Date() and IDTaquilla=" & txt(1), Reg
            '
            'actualiza la tabla de cheques
            cnnConexion.Execute "UPDATE TDFCheques SET IDTaquilla=" & txt(2) & " WHERE FechaMov" _
            & "=Date() AND IDTaquilla=" & txt(1)
            '
            'Actualiza la hora de apertura
            Set rstCaja = New ADODB.Recordset
            rstCaja.Open "SELECT Min(Hora) as H FROM MovimientoCaja WHERE IDTaquilla=" & txt(2) & _
            " or IDTaquilla=" & txt(1), cnnConexion, adOpenKeyset, adLockOptimistic, adCmdText
            
            cnnConexion.Execute "UPDATE Taquillas SET Hora='" & DateAdd("n", -5, rstCaja!H) & _
            "' WHERE IDTaquilla=" & txt(2)
            rstCaja.Close
            Set rstCaja = Nothing
            '
            If Err.Number = 0 Then
                '
                If Reg <= 0 Then
                    cnnConexion.RollbackTrans
                    MsgBox "La taquilla de origen no tiene transacciones para reasignar", _
                    vbInformation, App.ProductName
                    
                Else
                    cnnConexion.CommitTrans
                    MsgBox "Se reasignaron correctamente " & Reg & " transacciones", _
                    vbInformation, App.ProductName
                End If
                Call rtnBitacora("Reasignada " & Reg & " Operaciones " & txt(1) & "/" & txt(2))
            Else
                cnnConexion.RollbackTrans
                MsgBox Err.Description, vbExclamation, "Error " & Err.Number
                Call rtnBitacora("Error " & Err.Number & "al reasignar operaciones " & txt(1) _
                & "/" & txt(2))
            End If
    
        Case 1  'eliminar taquilla
            
            Set rstCaja = New ADODB.Recordset
            
            If IsNumeric(grid(1).TextMatrix(grid(1).RowSel, 0)) Then
                Reg = grid(1).TextMatrix(grid(1).RowSel, 0)
                'verifica que la taquilla no este abierta
                rstCaja.Open "SELECT * FROM Taquillas WHERE IDTaquilla=" & Reg, cnnConexion, _
                adOpenStatic, adLockReadOnly, adCmdText
                If Not rstCaja.EOF Or Not rstCaja.BOF Then
                    If rstCaja.Fields("Estado") = True Then
                        MsgBox "Imposible liberar la taquilla " & Reg & ", permanece abierta." & _
                        vbCrLf & "Cierre la taquilla y vuelva a intertarlo.", vbCritical, _
                        App.ProductName
                        rstCaja.Close
                        Set rstCaja = Nothing
                        Exit Sub
                    End If
                End If
                rstCaja.Close
                Set rstCaja = Nothing
                '
                If Respuesta("Desea dejar libre la taquilla " & Reg) Then
                    
                    cnnConexion.Execute "update  Taquillas IN '" & gcPath & "\tablas.mdb' set usuario=null,fecha=date(),asignado='" & gcUsuario & "'" _
                    & " WHERE IDTaquilla=" & Reg
                      Call Form_Load
                      'Grid(0).TextMatrix(Grid(0).RowSel, 0) = 0
                       
                      'grid(1) = ""
                     'Grid(1).TextMatrix(Grid(1).RowSel, 1) = ""
                    Call rtnBitacora("Liberada Taquilla " & Reg)
                    
                End If
            Else
                MsgBox "Seleccione una taquilla de la lista para eliminar", vbInformation, _
                App.ProductName
            End If
            
        Case 2  'salir
            Unload Me
            Set frmManCaja = Nothing
        
        Case 3  'AGREGAR UN NUEVO CAJERO
                Dim X As Integer
                Dim K As Integer
                Dim sql As String
                
                Set rstCaja = New ADODB.Recordset
        If (grid(0).TextMatrix(grid(0).RowSel, 0)) <> "" And grid(0).RowSel > 0 Then
          'si fue selccionado algo
                sql = "select *from taquillas in'" & gcPath & "/tablas.mdb'"
                Set rstCaja = New ADODB.Recordset
                rstCaja.Open sql, cnnConexion, adOpenStatic, adLockReadOnly, adCmdText
                 If Not rstCaja.EOF And Not rstCaja.BOF Then
                   rstCaja.MoveFirst
                     Do
                       X = rstCaja("idtaquilla")
                       If IsNull(rstCaja("usuario")) Then
                         K = X
                         Reg = X
                         sql = "Update  Taquillas IN '" & gcPath & "\tablas.mdb' set usuario='" & (grid(0).TextMatrix(grid(0).RowSel, 0)) & "',fecha=date(),asignado='" & gcUsuario & "'" _
                         & " WHERE IDTaquilla=" & Reg
                         Exit Do
                      End If
                        If rstCaja("usuario") = (grid(0).TextMatrix(grid(0).RowSel, 0)) Then
                              MsgBox "Usuario ya registaro", vbCritical, App.ProductName
                              Exit Sub
                        End If
                      rstCaja.MoveNext
                    Loop Until rstCaja.EOF
                      rstCaja.Close
                  Set rstCaja = Nothing
                      If K = 0 Then
                         K = X + 1
                         sql = "insert into taquillas(usuario,Asignado,IDtaquilla,fecha) in'" & gcPath & "/tablas.mdb' values('" & grid(0).TextMatrix(grid(0).RowSel, 0) & "','" & gcUsuario & "', '" & K & "',date())"
                      End If
                    End If
                   cnnConexion.Execute sql
                   Call Form_Load
                   MsgBox "Los datos fueron guardados", vbInformation, App.ProductName
                 Else
                   MsgBox "Seleccione un empleado de la lista empleado registrado", vbInformation, App.ProductName
                End If
      End Select
    End Sub

    Private Sub Form_Load()
    'variables locales
    Dim rstMe(1) As New ADODB.Recordset
    Dim strSql(1) As String
    Dim i As Integer
    '
    txt(0) = Format(Date, "dd/mm/yyyy")
    txt(0).Tag = Format(Date, "mm/dd/yyyy")
    strSql(0) = "SELECT NombreUsuario FROM Usuarios WHERE nivel < " & nuINACTIVO
    strSql(1) = "SELECT IDTaquilla,Usuario FROM Taquillas"
    
    For i = 0 To 1
        rstMe(i).Open strSql(i), cnnOLEDB + gcPath + "\tablas.mdb", adOpenKeyset, _
        adLockOptimistic, adCmdText
        Set grid(i).DataSource = rstMe(i)
        Set grid(i).FontFixed = LetraTitulo(LoadResString(527), 7.5, , True)
        Set grid(i).Font = LetraTitulo(LoadResString(528), 8)
        Call centra_titulo(grid(i), True)
    Next
    cmd(2).Picture = LoadResPicture("Salir", vbResIcon)
    'CenterForm frmManCaja
    End Sub

 

Private Sub txt_KeyPress(Index As Integer, KeyAscii As Integer)
Call Validacion(KeyAscii, "0123456789")
End Sub
