VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmNomAgui 
   Caption         =   "Nómina - Cargar Aguinaldos"
   ClientHeight    =   7350
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4920
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7350
   ScaleWidth      =   4920
   WindowState     =   2  'Maximized
   Begin ComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   480
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4920
      _ExtentX        =   8678
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
      MousePointer    =   99
      MouseIcon       =   "frmNomAgui.frx":0000
   End
   Begin VB.ListBox List1 
      Height          =   3960
      Index           =   0
      ItemData        =   "frmNomAgui.frx":031A
      Left            =   270
      List            =   "frmNomAgui.frx":031C
      Sorted          =   -1  'True
      TabIndex        =   9
      Top             =   3135
      Width           =   1755
   End
   Begin VB.ListBox List1 
      Height          =   3960
      Index           =   1
      ItemData        =   "frmNomAgui.frx":031E
      Left            =   2640
      List            =   "frmNomAgui.frx":0320
      Sorted          =   -1  'True
      TabIndex        =   8
      Top             =   3165
      Width           =   1755
   End
   Begin VB.CommandButton cmd 
      Height          =   375
      Index           =   1
      Left            =   2130
      Picture         =   "frmNomAgui.frx":0322
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   4950
      Width           =   375
   End
   Begin VB.CommandButton cmd 
      Height          =   375
      Index           =   0
      Left            =   2130
      Picture         =   "frmNomAgui.frx":03F7
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   4515
      Width           =   375
   End
   Begin VB.Frame fra 
      Caption         =   "Seleccione el año y el período a cargar:"
      Height          =   1980
      Left            =   240
      TabIndex        =   1
      Top             =   750
      Width           =   4215
      Begin VB.ComboBox cmb 
         Height          =   315
         ItemData        =   "frmNomAgui.frx":053D
         Left            =   2175
         List            =   "frmNomAgui.frx":0544
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   585
         Width           =   1245
      End
      Begin MSMask.MaskEdBox Cargado 
         Height          =   315
         Left            =   2190
         TabIndex        =   4
         Top             =   1155
         Width           =   1245
         _ExtentX        =   2196
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   7
         Format          =   "mm-yyyy"
         Mask            =   "##-####"
         PromptChar      =   "_"
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         Caption         =   "Cargar a Período:"
         Height          =   255
         Index           =   1
         Left            =   600
         TabIndex        =   3
         Top             =   1215
         Width           =   1365
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         Caption         =   "Aguinaldos año:"
         Height          =   255
         Index           =   0
         Left            =   615
         TabIndex        =   2
         Top             =   645
         Width           =   1365
      End
   End
   Begin ComctlLib.ImageList ImageList1 
      Left            =   11340
      Top             =   675
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
            Picture         =   "frmNomAgui.frx":054E
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmNomAgui.frx":06D0
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmNomAgui.frx":0852
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmNomAgui.frx":09D4
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmNomAgui.frx":0B56
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmNomAgui.frx":0CD8
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmNomAgui.frx":0E5A
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmNomAgui.frx":0FDC
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmNomAgui.frx":115E
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmNomAgui.frx":12E0
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmNomAgui.frx":1462
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmNomAgui.frx":15E4
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmNomAgui"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmd_Click(Index As Integer)
'variables locales
Dim I%, X%

X = IIf(Index = 0, 1, 0)
For I = 0 To List1(Index).ListCount - 1
    List1(X).AddItem List1(Index).List(I)
Next
List1(Index).Clear

End Sub

Private Sub Form_Load()
'variables locales
Dim rstlocal As New ADODB.Recordset
'
rstlocal.Open "Inmueble", cnnConexion, adOpenKeyset, adLockOptimistic, adCmdTable
With rstlocal
    .Filter = "Inactivo = False"
    If Not .EOF Then
        Do
            List1(1).AddItem !CodInm
            .MoveNext
        Loop Until .EOF
    End If
    .Close
End With
Set rstlocal = Nothing
cmb = cmb.List(0)
End Sub

Private Sub Form_Resize()
List1(0).Height = Me.ScaleHeight - List1(0).Top
List1(1).Height = List1(0).Height

End Sub

Private Sub List1_Click(Index As Integer)
Dim X%
X = IIf(Index = 0, 1, 0)
List1(X).AddItem List1(Index).Text
List1(Index).RemoveItem List1(Index).ListIndex
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As ComctlLib.Button)
Select Case UCase(Button.Key)
    
    Case "SAVE"
        Call carga_aguinaldos
        
    Case "CLOSE"
        Unload Me
        Set frmNomAgui = Nothing
        
End Select
End Sub

Private Sub carga_aguinaldos()
'variables locales
Dim rstCalculo As ADODB.Recordset
Dim I%, strSQL$, strBD$
Dim curDias@, curAguinaldo@, curSueldo@

'valida el período cargado
If Not IsDate("01/" & Cargado) Then
    MsgBox "Verifique el período a cargar", vbInformation, App.ProductName
    Exit Sub
End If
If cmb = "" Then
    MsgBox "Seleccione el año al que corresponden los aguinaldos", vbInformation, App.ProductName
    Exit Sub
End If
If List1(1).ListCount <= 0 Then
    MsgBox "Debe seleccionar por lo menos un inmueble", vbInformation, App.ProductName
Else
    Set rstCalculo = New ADODB.Recordset
    rstCalculo.Open "qdfAquinaldos", cnnConexion, adOpenKeyset, adLockOptimistic, adCmdTable
    
    For I = 0 To List1(1).ListCount - 1
    On Error Resume Next
        With rstCalculo
            .Filter = "CodInm ='" & List1(1).List(I) & "'"
            If Not .EOF Then
                .MoveFirst
                Do
                    'efectua los calculos
                    If DateDiff("m", !Fingreso, "30/11/" & Year(Date)) >= 12 Then
                        curDias = 15
                    Else
                        curDias = 15 / 12 * DateDiff("m", !Fingreso, "30/11/" & Year(Date))
                    End If
                    curSueldo = !Sueldo + (!Sueldo * (!Porc_BonoNoc / 100))
                    curAguinaldo = curSueldo / 30 * curDias
                    strBD = gcPath & "\" & !CodInm & "\inm.mdb"
                    
                    strSQL = "INSERT INTO AsignaGasto (Ndoc,CodGasto,Cargado,Descripcion,Fijo,Co" _
                   & "mun,Alicuota,Monto,Usuario,Fecha,Hora) in '" & strBD & "'  SELECT 'AGU" & _
                   cmb & "',CodGasto,'01-" & Cargado & "',Titulo,Fijo,Comun,Alicuota,'" & _
                   curAguinaldo & "','" & gcUsuario & "',Date(),Time() FROM TGastos IN '" & strBD & "' WHERE CodGasto ='190001'"
                   cnnConexion.BeginTrans
                   cnnConexion.Execute strSQL
                    If Err.Number <> 0 Then
                        Call rtnReporte(I, !CodInm, !CodEmp, !Apellidos & ", " & !Nombres, !NombreCargo, curDias, curAguinaldo, "Rechazado")
                        cnnConexion.RollbackTrans
                        Err.Number = 0
                    Else
                        Call rtnReporte(I, !CodInm, !CodEmp, !Apellidos & ", " & !Nombres, !NombreCargo, curDias, curAguinaldo, "Ok")
                        cnnConexion.CommitTrans
                    End If
                    
                    .MoveNext
                Loop Until .EOF
                
            End If
        End With
    Next
    MsgBox "Proceso Finalizado con  éxito", vbInformation, App.ProductName
    
End If

End Sub


    Private Sub rtnReporte(ParamArray D()) '
    '---------------------------------------------------------------------------------------------
    '
    Dim numFichero%, strArchivo$  'variables locales
    '------------------------
    numFichero = FreeFile
    strArchivo = App.Path & "\aguinaldo.dat"
    Open strArchivo For Append As numFichero
    Print #numFichero, D(0), D(1), D(2), D(3), D(4), D(5), D(6), D(7)
    Close numFichero
    '-------------------------
    End Sub

