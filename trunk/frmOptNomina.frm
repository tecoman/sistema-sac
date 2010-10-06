VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmOptNomina 
   Caption         =   "Parámtros de Nómina"
   ClientHeight    =   5925
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8760
   ControlBox      =   0   'False
   Icon            =   "frmOptNomina.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   11010
   ScaleWidth      =   15240
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmd 
      Caption         =   "<< Remover"
      Height          =   315
      Index           =   4
      Left            =   585
      TabIndex        =   46
      Top             =   8205
      Width           =   1335
   End
   Begin VB.ListBox Lst 
      Height          =   1815
      Index           =   1
      Left            =   4470
      TabIndex        =   44
      Top             =   7605
      Width           =   1350
   End
   Begin VB.ListBox Lst 
      Height          =   1815
      Index           =   0
      Left            =   2640
      TabIndex        =   43
      Top             =   7590
      Width           =   1350
   End
   Begin VB.CheckBox chk 
      Caption         =   "Fijo"
      Height          =   270
      Left            =   1725
      TabIndex        =   42
      Top             =   7230
      Width           =   720
   End
   Begin MSMask.MaskEdBox msk 
      Height          =   315
      Left            =   615
      TabIndex        =   41
      Top             =   7230
      Width           =   930
      _ExtentX        =   1640
      _ExtentY        =   556
      _Version        =   393216
      MaxLength       =   5
      Format          =   "dd/mm"
      Mask            =   "##/##"
      PromptChar      =   "_"
   End
   Begin VB.CommandButton cmd 
      Caption         =   "Registrar >>"
      Height          =   315
      Index           =   3
      Left            =   600
      TabIndex        =   40
      Top             =   7755
      Width           =   1335
   End
   Begin VB.CommandButton cmd 
      Cancel          =   -1  'True
      Caption         =   "Cancelar"
      Height          =   495
      Index           =   0
      Left            =   7365
      TabIndex        =   35
      Top             =   8760
      Width           =   1215
   End
   Begin VB.CommandButton cmd 
      Caption         =   "A&plicar"
      Height          =   495
      Index           =   1
      Left            =   7365
      TabIndex        =   34
      Top             =   8175
      Width           =   1215
   End
   Begin VB.CommandButton cmd 
      Caption         =   "Aceptar"
      Height          =   495
      Index           =   2
      Left            =   7365
      TabIndex        =   33
      Top             =   7605
      Width           =   1215
   End
   Begin VB.TextBox txt 
      Alignment       =   1  'Right Justify
      DataField       =   "codOtras_Asig"
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0000"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   8202
         SubFormatType   =   0
      EndProperty
      Height          =   315
      Index           =   10
      Left            =   3225
      TabIndex        =   32
      Text            =   "0"
      Top             =   4800
      Width           =   990
   End
   Begin VB.TextBox txt 
      Alignment       =   1  'Right Justify
      DataField       =   "codDia_Feriado"
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0000"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   8202
         SubFormatType   =   0
      EndProperty
      Height          =   315
      Index           =   9
      Left            =   6060
      TabIndex        =   31
      Text            =   "0"
      Top             =   4260
      Width           =   990
   End
   Begin VB.TextBox txt 
      Alignment       =   1  'Right Justify
      DataField       =   "codDia_Libre"
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0000"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   8202
         SubFormatType   =   0
      EndProperty
      Height          =   315
      Index           =   8
      Left            =   3225
      TabIndex        =   30
      Text            =   "0"
      Top             =   4290
      Width           =   990
   End
   Begin VB.TextBox txt 
      Alignment       =   1  'Right Justify
      DataField       =   "Dia_adi"
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "#,###"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   8202
         SubFormatType   =   0
      EndProperty
      Height          =   315
      Index           =   7
      Left            =   7155
      TabIndex        =   25
      Text            =   "0"
      Top             =   2970
      Width           =   525
   End
   Begin VB.TextBox txt 
      Alignment       =   1  'Right Justify
      DataField       =   "adi_ano"
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "#,##0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   8202
         SubFormatType   =   0
      EndProperty
      Height          =   315
      Index           =   6
      Left            =   5370
      TabIndex        =   23
      Text            =   "0"
      Top             =   2970
      Width           =   525
   End
   Begin VB.TextBox txt 
      Alignment       =   1  'Right Justify
      DataField       =   "dia_bono"
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "#,##0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   8202
         SubFormatType   =   0
      EndProperty
      Height          =   315
      Index           =   5
      Left            =   3495
      TabIndex        =   21
      Text            =   "0"
      Top             =   3510
      Width           =   705
   End
   Begin VB.TextBox txt 
      Alignment       =   1  'Right Justify
      DataField       =   "Dia_vacacion"
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "#,##0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   8202
         SubFormatType   =   0
      EndProperty
      Height          =   315
      Index           =   4
      Left            =   3495
      TabIndex        =   20
      Text            =   "0"
      Top             =   3015
      Width           =   705
   End
   Begin MSDataListLib.DataCombo dtc 
      Height          =   315
      Left            =   5370
      TabIndex        =   16
      Top             =   2070
      Width           =   3090
      _ExtentX        =   5450
      _ExtentY        =   556
      _Version        =   393216
      ListField       =   "Nombre"
      BoundColumn     =   "CodCargo"
      Text            =   "DataCombo1"
   End
   Begin VB.TextBox txt 
      Alignment       =   1  'Right Justify
      DataField       =   "Dia_Feriado"
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "#,##0.00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   8202
         SubFormatType   =   0
      EndProperty
      Height          =   315
      Index           =   3
      Left            =   3315
      TabIndex        =   12
      Text            =   "0,00"
      Top             =   1560
      Width           =   705
   End
   Begin VB.TextBox txt 
      Alignment       =   1  'Right Justify
      DataField       =   "LPH"
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "#,##0.00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   8202
         SubFormatType   =   0
      EndProperty
      Height          =   315
      Index           =   2
      Left            =   6825
      TabIndex        =   8
      Text            =   "0,00"
      Top             =   645
      Width           =   705
   End
   Begin VB.TextBox txt 
      Alignment       =   1  'Right Justify
      DataField       =   "SPF"
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "#,##0.00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   8202
         SubFormatType   =   0
      EndProperty
      Height          =   315
      Index           =   1
      Left            =   4530
      TabIndex        =   5
      Text            =   "0,00"
      Top             =   645
      Width           =   705
   End
   Begin VB.TextBox txt 
      Alignment       =   1  'Right Justify
      DataField       =   "SSO"
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "#,##0.00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   8202
         SubFormatType   =   0
      EndProperty
      Height          =   315
      Index           =   0
      Left            =   2370
      TabIndex        =   2
      Text            =   "0,00"
      Top             =   645
      Width           =   705
   End
   Begin VB.Label lbl 
      Caption         =   "Formato: (dd/mm)"
      BeginProperty Font 
         Name            =   "Agency FB"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   26
      Left            =   630
      TabIndex        =   47
      Top             =   6960
      Width           =   1110
   End
   Begin VB.Label lbl 
      Caption         =   "Para remover un elemento de la lista haga doble click sobre su selección o seleccione el elemento y presione el botón remover."
      Height          =   540
      Index           =   25
      Left            =   1965
      TabIndex        =   45
      Top             =   6450
      Width           =   6570
   End
   Begin VB.Label lbl 
      Caption         =   $"frmOptNomina.frx":27A2
      Height          =   540
      Index           =   24
      Left            =   1980
      TabIndex        =   39
      Top             =   5880
      Width           =   6570
   End
   Begin VB.Label lbl 
      Alignment       =   2  'Center
      Caption         =   "Variables"
      Height          =   270
      Index           =   23
      Left            =   4560
      TabIndex        =   38
      Top             =   7230
      Width           =   1350
   End
   Begin VB.Label lbl 
      Alignment       =   2  'Center
      Caption         =   "Fijos"
      Height          =   270
      Index           =   22
      Left            =   2670
      TabIndex        =   37
      Top             =   7230
      Width           =   1350
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      Index           =   6
      X1              =   270
      X2              =   8535
      Y1              =   5490
      Y2              =   5490
   End
   Begin VB.Label lbl 
      Caption         =   "Días Feriados:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   21
      Left            =   255
      TabIndex        =   36
      Top             =   5865
      Width           =   1605
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      Caption         =   "Otras Asignaciones:"
      Height          =   270
      Index           =   20
      Left            =   1185
      TabIndex        =   29
      Top             =   4875
      Width           =   1890
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      Caption         =   "Dias Feriados:"
      Height          =   270
      Index           =   19
      Left            =   4695
      TabIndex        =   28
      Top             =   4305
      Width           =   1245
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      Caption         =   "Dias Libres:"
      Height          =   270
      Index           =   18
      Left            =   1830
      TabIndex        =   27
      Top             =   4320
      Width           =   1245
   End
   Begin VB.Label lbl 
      Caption         =   "Códigos:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   17
      Left            =   225
      TabIndex        =   26
      Top             =   4320
      Width           =   1605
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      Index           =   4
      X1              =   240
      X2              =   8505
      Y1              =   4080
      Y2              =   4080
   End
   Begin VB.Label lbl 
      Caption         =   "año cancelar                adicional"
      Height          =   270
      Index           =   16
      Left            =   6030
      TabIndex        =   24
      Top             =   2985
      Width           =   2475
   End
   Begin VB.Label lbl 
      Caption         =   "A partir de "
      Height          =   270
      Index           =   15
      Left            =   4530
      TabIndex        =   22
      Top             =   2985
      Width           =   840
   End
   Begin VB.Label lbl 
      Caption         =   "Días Bono:"
      Height          =   270
      Index           =   14
      Left            =   1875
      TabIndex        =   19
      Top             =   3525
      Width           =   1245
   End
   Begin VB.Label lbl 
      Caption         =   "Días Vacaciones:"
      Height          =   270
      Index           =   13
      Left            =   1875
      TabIndex        =   18
      Top             =   3030
      Width           =   1680
   End
   Begin VB.Label lbl 
      Caption         =   "Vacaciones:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   12
      Left            =   225
      TabIndex        =   17
      Top             =   2895
      Width           =   1605
   End
   Begin VB.Label lbl 
      Caption         =   "Seleccione el cargo que devengará bono nocturno:"
      Height          =   270
      Index           =   11
      Left            =   1590
      TabIndex        =   15
      Top             =   2100
      Width           =   3900
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      Index           =   3
      X1              =   225
      X2              =   8490
      Y1              =   2670
      Y2              =   2670
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderWidth     =   2
      Index           =   2
      X1              =   210
      X2              =   8475
      Y1              =   2670
      Y2              =   2670
   End
   Begin VB.Label lbl 
      Caption         =   "de un día normal"
      Height          =   270
      Index           =   10
      Left            =   4305
      TabIndex        =   14
      Top             =   1605
      Width           =   1455
   End
   Begin VB.Label lbl 
      Caption         =   "%"
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
      Index           =   9
      Left            =   4080
      TabIndex        =   13
      Top             =   1605
      Width           =   255
   End
   Begin VB.Label lbl 
      Caption         =   "Día Feriado Trabajado."
      Height          =   270
      Index           =   8
      Left            =   1605
      TabIndex        =   11
      Top             =   1605
      Width           =   1725
   End
   Begin VB.Label lbl 
      Caption         =   "Asignaciones:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   7
      Left            =   225
      TabIndex        =   10
      Top             =   1425
      Width           =   1605
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      Index           =   1
      X1              =   240
      X2              =   8505
      Y1              =   1230
      Y2              =   1230
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderWidth     =   2
      Index           =   0
      X1              =   225
      X2              =   8490
      Y1              =   1230
      Y2              =   1230
   End
   Begin VB.Label lbl 
      Caption         =   "%"
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
      Index           =   6
      Left            =   7545
      TabIndex        =   9
      Top             =   690
      Width           =   285
   End
   Begin VB.Label lbl 
      Caption         =   "L.P.H."
      Height          =   270
      Index           =   5
      Left            =   6315
      TabIndex        =   7
      Top             =   690
      Width           =   825
   End
   Begin VB.Label lbl 
      Caption         =   "%"
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
      Index           =   4
      Left            =   5220
      TabIndex        =   6
      Top             =   690
      Width           =   285
   End
   Begin VB.Label lbl 
      Caption         =   "S.P.F."
      Height          =   270
      Index           =   3
      Left            =   3975
      TabIndex        =   4
      Top             =   690
      Width           =   825
   End
   Begin VB.Label lbl 
      Caption         =   "%"
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
      Index           =   2
      Left            =   3120
      TabIndex        =   3
      Top             =   690
      Width           =   285
   End
   Begin VB.Label lbl 
      Caption         =   "S.S.O:"
      Height          =   270
      Index           =   1
      Left            =   1725
      TabIndex        =   1
      Top             =   690
      Width           =   825
   End
   Begin VB.Label lbl 
      Caption         =   "Deducciones:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   0
      Left            =   270
      TabIndex        =   0
      Top             =   255
      Width           =   1605
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderWidth     =   2
      Index           =   5
      X1              =   225
      X2              =   8490
      Y1              =   4080
      Y2              =   4080
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderWidth     =   2
      Index           =   7
      X1              =   255
      X2              =   8520
      Y1              =   5490
      Y2              =   5490
   End
End
Attribute VB_Name = "frmOptNomina"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rstCal(2) As New ADODB.Recordset
Dim BN As Long


Private Sub Cmd_Click(Index As Integer)
'variables locales
Dim Frm As Form
Dim Diaf As Date
Select Case Index
    Case 0: Unload Me
    Case 1, 2
        If Not Error Then
            If Trim(Txt(8)) = "" Then Txt(8) = 0
            If Trim(Txt(9)) = "" Then Txt(9) = 0
            If Trim(Txt(10)) = "" Then Txt(10) = 0
            rstCal(0).Fields("Bono_Noc") = BN
            rstCal(0).Update
            
            For Each Frm In Forms
                If UCase(Frm.Name) = "FRMFICHAEMP" Then FrmFichaEmp.BN = BN
            Next
            
            If Index = 2 Then Unload Me
        End If
    
    Case 3  'añadir dia feriado
    
        With rstCal(2)
            Diaf = msk & "/" & Year(Date)
            If IsDate(Diaf) Then
                If Not (.EOF And .BOF) Then '
                    .MoveFirst
                    .Find "Fecha =#" & Format(msk, "mm/dd/yy") & "#"
                    If Not .EOF Then
                        MsgBox "Feriado ya registrado", vbInformation, App.ProductName
                        Exit Sub
                    End If
                End If
                .AddNew
                !fecha = Diaf
                !Fijo = IIf(CHK = vbChecked, True, False)
                Lst(IIf(!Fijo, 0, 1)).AddItem Diaf
                .Update
                msk.PromptInclude = False
                msk = ""
                msk.PromptInclude = True
                
                
            Else
                MsgBox "Fecha inválida", vbInformation, App.ProductName
            End If
            
        End With
        '
    Case 4  'remover dia feriado
        Diaf = msk & "/" & Year(Date)
        If IsDate(Diaf) Then
        With rstCal(2)
            If Not (.EOF And .BOF) Then
                .MoveFirst
                .Find "Fecha=#" & Format(Diaf, "dd/mm/yy") & "#"
                If Not .EOF Then
                    .Delete
                    msk.PromptInclude = False
                    msk = ""
                    msk.PromptInclude = True
                    cargar_diasferiados
                End If
            End If
        End With
        Else
            MsgBox "Introduzca una fecha válida", vbInformation, App.ProductName
        End If
        'lst(IIf(CHK, 1, 0)).RemoveItem lst(IIf(CHK, 1, 0)).Selected
        
End Select
End Sub

Private Sub dtc_Click(Area As Integer)
If Area = 2 Then
    BN = dtc.BoundText
End If
End Sub

Private Sub Form_Load()
'carga del formulario
rstCal(0).Open "Nom_Calc", cnnConexion, adOpenKeyset, adLockOptimistic, adCmdTable
rstCal(1).Open "SELECT * FROM Cargos ORDER BY Nombre", cnnOLEDB + gcPath & "\tablas.mdb", _
adOpenKeyset, adLockOptimistic, adCmdText
rstCal(2).Open "Nom_Feriados", cnnConexion, adOpenKeyset, adLockOptimistic, adCmdTable
Set dtc.RowSource = rstCal(1)
'
For Each C In Me.Controls
    If TypeName(C) = "TextBox" Then Set C.DataSource = rstCal(0)
Next
Call CenterForm(frmOptNomina)
With rstCal(1)
    If Not (.EOF And .BOF) Then
        .MoveFirst
        .Find "CodCargo=" & rstCal(0)!Bono_Noc
        If .EOF Then
            dtc = "INTRODUZCA EL CARGO"
        Else
            dtc.Text = .Fields("Nombre")
        End If
    End If
End With
'
cargar_diasferiados
'
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
For i = 0 To 2
    rstCal(i).Close
    Set rstCal(i) = Nothing
Next
End Sub

Private Sub txt_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii = Asc(".") Then KeyAscii = Asc(",")
Call Validacion(KeyAscii, "012345679,")
If KeyAscii = 13 Then SendKeys vbTab
End Sub

Private Function Error() As Boolean
For Each C In Me.Controls
    If TypeName(C) = "TextBox" Then
        If Not IsNumeric(C) Then
            Error = MsgBox("Introdujo un valor no válido", vbExclamation, App.ProductName)
            C.SelStart = 0
            C.SelLength = Len(C)
            C.SetFocus
            Exit For
        End If
    End If
Next
End Function

Private Sub cargar_diasferiados()
Lst(0).Clear
Lst(1).Clear
With rstCal(2)
    .Requery
    If Not .EOF And Not .BOF Then
        .MoveFirst
        Do
            Lst(IIf(!Fijo, 0, 1)).AddItem (!fecha)
            .MoveNext
        Loop Until .EOF
    End If
End With

End Sub
