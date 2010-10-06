VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmCorrespondencia 
   Caption         =   "Correspondencia"
   ClientHeight    =   7380
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11940
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7380
   ScaleWidth      =   11940
   WindowState     =   2  'Maximized
   Begin VB.Frame fra 
      Height          =   7090
      Index           =   0
      Left            =   150
      TabIndex        =   0
      Top             =   200
      Width           =   11660
      Begin VB.CommandButton cmd 
         Caption         =   "&Refrescar Lista"
         Height          =   900
         Index           =   5
         Left            =   8430
         TabIndex        =   32
         Top             =   5730
         Width           =   975
      End
      Begin VB.TextBox txt 
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   315
         Index           =   3
         Left            =   8550
         TabIndex        =   30
         Top             =   420
         Width           =   825
      End
      Begin VB.CommandButton cmd 
         Caption         =   "&Imprimir"
         Height          =   900
         Index           =   4
         Left            =   9405
         TabIndex        =   29
         Top             =   5730
         Width           =   975
      End
      Begin VB.CommandButton cmd 
         Caption         =   "&Salir"
         Height          =   900
         Index           =   3
         Left            =   10395
         TabIndex        =   28
         Top             =   5730
         Width           =   975
      End
      Begin VB.Frame Frame2 
         Caption         =   "Otros:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   930
         Left            =   330
         TabIndex        =   22
         Top             =   3750
         Width           =   11070
         Begin VB.CommandButton cmd 
            Caption         =   "Agregar"
            Height          =   405
            Index           =   2
            Left            =   9885
            TabIndex        =   25
            Top             =   300
            Width           =   975
         End
         Begin VB.TextBox txt 
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   2
            Left            =   5175
            TabIndex        =   24
            Top             =   345
            Width           =   4605
         End
         Begin VB.Label lbl 
            Caption         =   "Si desea enviar algún otro documento, desde aqui puede registrarlo. Señale una pequeña descripción:"
            BeginProperty Font 
               Name            =   "Trebuchet MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   450
            Index           =   11
            Left            =   270
            TabIndex        =   23
            Top             =   270
            Width           =   4845
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Facturación y Caja:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1905
         Left            =   6630
         TabIndex        =   14
         Top             =   1680
         Width           =   4815
         Begin VB.CommandButton cmd 
            Caption         =   "Agregar"
            Height          =   405
            Index           =   0
            Left            =   3585
            TabIndex        =   17
            Top             =   1275
            Width           =   975
         End
         Begin VB.TextBox txt 
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   1
            Left            =   2670
            TabIndex        =   16
            Top             =   1335
            Width           =   825
         End
         Begin VB.ComboBox cmb 
            Height          =   315
            ItemData        =   "frmCorrespondencia.frx":0000
            Left            =   270
            List            =   "frmCorrespondencia.frx":000D
            TabIndex        =   15
            Top             =   1335
            Width           =   2325
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            Caption         =   "Período (MM/AAAA)"
            BeginProperty Font 
               Name            =   "Trebuchet MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   465
            Index           =   10
            Left            =   2625
            TabIndex        =   20
            Top             =   810
            Width           =   900
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            Caption         =   "Documento"
            BeginProperty Font 
               Name            =   "Trebuchet MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   9
            Left            =   225
            TabIndex        =   19
            Top             =   1050
            Width           =   2295
         End
         Begin VB.Label lbl 
            Caption         =   "Anexe documentación relacionada al proceso de facturación y caja."
            BeginProperty Font 
               Name            =   "Trebuchet MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   450
            Index           =   8
            Left            =   270
            TabIndex        =   18
            Top             =   300
            Width           =   4305
         End
      End
      Begin VB.Frame fra 
         Caption         =   "Cheques:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1905
         Index           =   2
         Left            =   300
         TabIndex        =   6
         Top             =   1680
         Width           =   6120
         Begin VB.CommandButton cmd 
            Caption         =   "Agregar"
            Height          =   405
            Index           =   1
            Left            =   4950
            TabIndex        =   21
            Top             =   1305
            Width           =   975
         End
         Begin VB.TextBox txt 
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   0
            Left            =   270
            TabIndex        =   8
            Top             =   1365
            Width           =   825
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Trebuchet MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   7
            Left            =   3615
            TabIndex        =   13
            Top             =   1365
            Width           =   1260
         End
         Begin VB.Label lbl 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Trebuchet MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   6
            Left            =   1200
            TabIndex        =   12
            Top             =   1365
            Width           =   2340
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            Caption         =   "Monto"
            BeginProperty Font 
               Name            =   "Trebuchet MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   5
            Left            =   3630
            TabIndex        =   11
            Top             =   1050
            Width           =   1215
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            Caption         =   "Beneficiario"
            BeginProperty Font 
               Name            =   "Trebuchet MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   4
            Left            =   1230
            TabIndex        =   10
            Top             =   1050
            Width           =   2310
         End
         Begin VB.Label lbl 
            Caption         =   "Introduzca el Nº cheque y presione la tecla <<enter>> para buscar la información del cheque"
            BeginProperty Font 
               Name            =   "Trebuchet MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   450
            Index           =   3
            Left            =   270
            TabIndex        =   9
            Top             =   285
            Width           =   5115
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            Caption         =   "Nº Cheque"
            BeginProperty Font 
               Name            =   "Trebuchet MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   2
            Left            =   270
            TabIndex        =   7
            Top             =   1050
            Width           =   825
         End
      End
      Begin VB.Frame fra 
         Caption         =   "Seleccione un Inmueble:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1230
         Index           =   1
         Left            =   300
         TabIndex        =   1
         Top             =   300
         Width           =   6990
         Begin MSDataListLib.DataCombo dtc 
            Height          =   315
            Index           =   0
            Left            =   285
            TabIndex        =   2
            Top             =   675
            Width           =   1125
            _ExtentX        =   1984
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo dtc 
            Height          =   315
            Index           =   1
            Left            =   1470
            TabIndex        =   3
            Top             =   660
            Width           =   4980
            _ExtentX        =   8784
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            Caption         =   "Nombre / Razón Social"
            BeginProperty Font 
               Name            =   "Trebuchet MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   1
            Left            =   2025
            TabIndex        =   5
            Top             =   375
            Width           =   3900
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            Caption         =   "Código"
            BeginProperty Font 
               Name            =   "Trebuchet MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   0
            Left            =   240
            TabIndex        =   4
            Top             =   375
            Width           =   1215
         End
      End
      Begin MSComctlLib.ListView lst 
         Height          =   1725
         Left            =   360
         TabIndex        =   26
         Top             =   5115
         Width           =   7440
         _ExtentX        =   13123
         _ExtentY        =   3043
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
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   4
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Cod.Inm"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Documento"
            Object.Width           =   7056
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Fecha"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Hora"
            Object.Width           =   1764
         EndProperty
      End
      Begin VB.Shape Shape1 
         Height          =   585
         Left            =   7905
         Top             =   4785
         Width           =   3495
      End
      Begin VB.Label lbl 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Para eliminar una linea de la lista, seleccionela y presione la tecla suprimir"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   525
         Index           =   14
         Left            =   7995
         TabIndex        =   33
         Top             =   4815
         Width           =   3375
      End
      Begin VB.Label lbl 
         Alignment       =   2  'Center
         Caption         =   "Reporte Nº:"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   13
         Left            =   7500
         TabIndex        =   31
         Top             =   450
         Width           =   1005
      End
      Begin VB.Label lbl 
         Alignment       =   2  'Center
         BackColor       =   &H80000002&
         Caption         =   ".::DETALLE REPORTE::."
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000009&
         Height          =   315
         Index           =   12
         Left            =   360
         TabIndex        =   27
         Top             =   4800
         Width           =   7440
      End
   End
End
Attribute VB_Name = "frmCorrespondencia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim mCuentas As String
Dim mBD As String
Dim Nreport As Long

Private Sub cmd_Click(Index As Integer)
Select Case Index
    Case 3  'cerrar formulario
        Unload Me
        Set frmCorrespondencia = Nothing
    
    Case 1  'agregar cheque
        If txt(0) <> "" And lbl(6) <> "" And lbl(7) <> "" Then
            agregar_item ("CHEQUE #" & txt(0) & " " & lbl(6) & " " & lbl(7))
            txt(0) = ""
            lbl(6) = ""
            lbl(7) = ""
        Else
            MsgBox "No se puede agregar la información. Complete los datos", _
            vbInformation, App.ProductName
        End If
        
    Case 0
        If IsDate("01/" & txt(1)) Then
            agregar_item (cmb & " PERÍODO: " & txt(1))
            cmb = ""
            txt(1) = ""
        Else
            MsgBox "Introdujo un período inválido", vbInformation, App.ProductName
        End If
    
    Case 2
        If txt(2) <> "" Then
            agregar_item (txt(2))
            txt(2) = ""
        End If
    Case 4
        Printer_Reporte
            
    Case 5
        Listar
        
End Select
End Sub

Private Sub dtc_Click(Index As Integer, Area As Integer)
If Area = 2 Then
    BuscaInm (Index)
End If
End Sub

Private Sub dtc_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii = 13 Then
    If Dtc(0).Text <> "" Or Dtc(1).Text <> "" Then BuscaInm (Index)
End If
End Sub

Private Sub Form_Load()
'configura el origen de los datos de los data combo Inmuebles
Set Dtc(0).RowSource = FrmAdmin.objRst
Dtc(0).ListField = "CodInm"
Set Dtc(1).RowSource = FrmAdmin.ObjRstNom
Dtc(1).ListField = "Nombre"
'
txt(3) = Format(ftnReport, "0000")

Call Listar

End Sub

'rutina que busca la información de un inmueble determinado
'para obtener las variables de la conexion
Private Sub BuscaInm(Index As Integer)
Dim strCriterio As String
Dim rstlocal As ADODB.Recordset
Dim J%

If Index = 0 Then
    strCriterio = "CodInm='" & Dtc(0) & "'"
Else
    strCriterio = "Nombre LIKE '*" & Dtc(1) & "*'"
End If
mCuentas = ""
With FrmAdmin.objRst
    If Not (.EOF And .BOF) Then
        .MoveFirst
        .Find strCriterio
        If Not .EOF Then
            Dtc(0) = !CodInm
            Dtc(1) = !Nombre
            Set rstlocal = New ADODB.Recordset
            If !Caja <> sysCodCaja Then
                mBD = !Ubica & "inm.mdb"
            Else
                mBD = "\" & sysCodInm & "\inm.mdb"
            End If
            rstlocal.Open "Cuentas", cnnOLEDB + gcPath + mBD, adOpenStatic, adLockReadOnly, adCmdTable
            If Not (rstlocal.EOF And rstlocal.BOF) Then
                Do
                    mCuentas = mCuentas & IIf(J = 0, "('" & rstlocal!NumCuenta & "'", ",'" _
                    & rstlocal!NumCuenta & "'")
                    J = J + 1
                    rstlocal.MoveNext
                Loop Until rstlocal.EOF
                mCuentas = mCuentas & ")"
            End If
            rstlocal.Close
            Set rstlocal = Nothing
        Else
            .MoveFirst
        End If
    End If
End With
End Sub


Private Sub lst_KeyUp(KeyCode As Integer, Shift As Integer)
Dim n As Integer
If KeyCode = 46 And Shift = 0 Then  'eliminar el item seleccionado
    If Respuesta("¿Desea eliminar esta línea?") Then
        cnnConexion.Execute "DELETE * FROM Correspondencia WHERE Cstr(IDReporte & CodInm " & _
        "& Documento & Fecha & Hora) ='" & CLng(txt(3)) & lst.SelectedItem & _
        lst.SelectedItem.ListSubItems(1) & lst.SelectedItem.ListSubItems(2) & _
        lst.SelectedItem.ListSubItems(3) & "'", n
        Listar
        MsgBox n & " registro(s) eliminado(s).", vbInformation, App.ProductName
        
    End If
End If
End Sub

Private Sub txt_KeyPress(Index As Integer, KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
Select Case Index
    Case 0
        Call Validacion(KeyAscii, "01234567890")
        If KeyAscii = 13 And mBD <> "" Then BuscaCheque
    Case 1
        If KeyAscii = Asc("-") Then KeyAscii = Asc("/")
        Call Validacion(KeyAscii, "0123456789/")
    Case 3
        Call Validacion(KeyAscii, "0123456789")
End Select
End Sub

Private Sub BuscaCheque()
'variables locales
Dim rstlocal As ADODB.Recordset
Dim strSql As String

Set rstlocal = New ADODB.Recordset

strSql = "SELECT Format(Cheque.IDcheque,'000000') as ID, Format(Sum(ChequeDetalle.Monto),'#" _
& ",##0.00') AS Monto, Format(Cheque.FechaCheque,'dd/mm/yy') as Fecha,Cheque.Beneficiario, " _
& "Cheque.Banco, Cheque.Cuenta, Cheque.Concepto,Cheque.Impreso FROM Cheque INNER JOIN Chequ" _
& "eDetalle ON Cheque.Clave = ChequeDetalle.Clave WHERE Cheque.IDcheque=" & txt(0) & " AND " _
& "Cheque.Cuenta IN " & mCuentas & " GROUP BY Cheque.IDCheque, Cheque.FechaCheque, " _
& "Cheque.Beneficiario, Cheque.Banco,Cheque.Cuenta,Cheque.Concepto,cheque.Impreso"

With rstlocal
    .Open strSql, cnnConexion, adOpenStatic, adLockReadOnly, adCmdText
    If Not (.EOF And .BOF) Then
        lbl(6) = Space(1) & !Beneficiario
        lbl(7) = !Monto
    Else
        MsgBox "No se encuentra información del cheque " & txt(0), vbInformation, _
        App.ProductName
    End If
    .Close
End With
Set rstlocal = Nothing
End Sub

Private Sub agregar_item(Documento As String)
Documento = UCase(Documento)
cnnConexion.Execute "INSERT INTO Correspondencia(IDReporte,CodInm,Documento,Usuario,Fecha,Hora) " & _
"VALUES (" & IIf(IsNumeric(txt(3)), txt(3), Nreport) & ",'" & Dtc(0) & "','" & Documento & "','" & gcUsuario & "',Date(),Time())"
lst.ListItems.Add , , Dtc(0)
lst.ListItems(lst.ListItems.Count).ListSubItems.Add , , Documento
lst.ListItems(lst.ListItems.Count).ListSubItems.Add , , Date
lst.ListItems(lst.ListItems.Count).ListSubItems.Add , , Time
End Sub


Private Sub Listar()
Dim rstlocal As ADODB.Recordset
Dim x%
Set rstlocal = New ADODB.Recordset
With rstlocal
    .CursorLocation = adUseClient
    .Open "Correspondencia", cnnConexion, adOpenStatic, adLockReadOnly, adCmdTable
    .Filter = "IDReporte=" & txt(3)
    lst.ListItems.Clear
    If Not (.EOF And .BOF) Then
        .MoveFirst
        
        Do
            x = x + 1
            lst.ListItems.Add , , !CodInm
            lst.ListItems(x).ListSubItems.Add , , !Documento
            lst.ListItems(x).ListSubItems.Add , , !fecha
            lst.ListItems(x).ListSubItems.Add , , !Hora
            .MoveNext
        Loop Until .EOF
    End If
    .Close
    
End With
    
Set rstlocal = Nothing
End Sub


Private Sub Printer_Reporte()
'variables locales
Dim strSql As String
Dim rstlocal As ADODB.Recordset
Dim StrCodInm As String
Dim n%
Set rstlocal = New ADODB.Recordset

strSql = "SELECT Correspondencia.*, Inmueble.Nombre FROM Inmueble " _
& "INNER JOIN Correspondencia ON Inmueble.CodInm = Correspondencia.CodInm;"


With rstlocal
    .CursorLocation = adUseClient
    .Open strSql, cnnConexion, adOpenStatic, adLockReadOnly, adCmdText
    .Filter = "IDReporte =" & IIf(IsNumeric(txt(3)), txt(3), Nreport)
    .Sort = "CodInm, Fecha, Hora"
    If Not (.EOF And .BOF) Then
        .MoveFirst
        Printer.FontBold = True
        Printer.FontSize = 14
        Printer.Print "Reporte Correspondencia"
        Printer.FontBold = False
        Printer.FontSize = 10
        Printer.Print "Fecha: " & Date
        Printer.Print "Hora: " & Time
        'Printer.FontSize = Printer.FontSize - 2
        '
        Printer.Print
        Do
            'If !CodInm = "2539" Then Stop
            
            If StrCodInm <> !CodInm Then    'encabezado condominio
                
                If Not StrCodInm = "" Then
                    Printer.CurrentY = Printer.CurrentY + 150
                    Printer.Print "COMENTARIOS"
                    For n = 1 To 4
                        Printer.CurrentY = Printer.CurrentY + 300
                        Printer.Line (Printer.ScaleLeft + 200, Printer.CurrentY)-(Printer.ScaleWidth, Printer.CurrentY + 2), vbBlack, B
                    Next
                    Printer.CurrentX = Printer.ScaleWidth - Printer.TextWidth("Recibí Conforme:________________________")
                    Printer.CurrentY = Printer.CurrentY + 400
                    Printer.Print "Recibí Conforme:________________________"
                    Printer.CurrentY = Printer.CurrentY + 600
                    Printer.CurrentX = 0
                End If
                If Printer.CurrentY > 13000 Then Printer.NewPage
                Printer.FontBold = True
                Printer.FontSize = 12
                Printer.Print !CodInm & " " & !Nombre
                Printer.FontBold = False
                Printer.FontSize = 9
                Printer.CurrentY = Printer.CurrentY + 100
                n = 0
            End If
            n = n + 1
            StrCodInm = !CodInm
            Printer.CurrentX = Printer.CurrentX + 200
            Printer.Print n & ".- " & !Documento
            .MoveNext
        Loop Until .EOF
        Printer.CurrentY = Printer.CurrentY + 150
        Printer.Print "COMENTARIOS"
        For n = 1 To 4
            Printer.CurrentY = Printer.CurrentY + 300
            Printer.Line (Printer.ScaleLeft, Printer.CurrentY)-(Printer.ScaleWidth, Printer.CurrentY + 2), vbBlack, B
        Next
        
        Printer.EndDoc
        MsgBox "Reporte impreso con éxito", vbInformation, App.ProductName
    End If
    .Close
End With
Set rstlocal = Nothing
End Sub

Private Sub txt_Validate(Index As Integer, Cancel As Boolean)
If Index = 3 Then If IsNumeric(txt(3)) Then txt(3) = Format(txt(3), "0000")
End Sub

Private Function ftnReport() As Long
Dim rstlocal As ADODB.Recordset
Set rstlocal = New ADODB.Recordset

rstlocal.Open "SELECT Max(IDReporte) FROM Correspondencia", cnnConexion, adOpenKeyset, _
adLockOptimistic, admcdtext
Nreport = IIf(IsNull(rstlocal.Fields(0)), 0, rstlocal.Fields(0)) + 1
rstlocal.Close
Set rstlocal = Nothing
ftnReport = Nreport
End Function
