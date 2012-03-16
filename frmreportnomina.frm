VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmreportnomina 
   Caption         =   "Reporte de Nomina"
   ClientHeight    =   5190
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11895
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   5190
   ScaleWidth      =   11895
   WindowState     =   2  'Maximized
   Begin TabDlg.SSTab SSTab1 
      Height          =   7290
      Left            =   120
      TabIndex        =   0
      Top             =   60
      Width           =   11775
      _ExtentX        =   20770
      _ExtentY        =   12859
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Generales"
      TabPicture(0)   =   "frmreportnomina.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lbl(4)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "List1(0)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Frame1(0)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "cmd(0)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "cmd(2)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "cmd(3)"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "List1(1)"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "cmd(1)"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).ControlCount=   8
      TabCaption(1)   =   "Especificos"
      TabPicture(1)   =   "frmreportnomina.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "cmd(8)"
      Tab(1).Control(1)=   "cmd(9)"
      Tab(1).Control(2)=   "List1(4)"
      Tab(1).Control(3)=   "List1(5)"
      Tab(1).Control(4)=   "Frame1(1)"
      Tab(1).Control(5)=   "cmd(7)"
      Tab(1).Control(6)=   "List1(3)"
      Tab(1).Control(7)=   "List1(2)"
      Tab(1).Control(8)=   "cmd(6)"
      Tab(1).Control(9)=   "cmd(5)"
      Tab(1).Control(10)=   "cmd(4)"
      Tab(1).Control(11)=   "lbl(5)"
      Tab(1).Control(12)=   "lbl(6)"
      Tab(1).ControlCount=   13
      Begin VB.CommandButton cmd 
         Height          =   375
         Index           =   8
         Left            =   -72345
         Picture         =   "frmreportnomina.frx":0038
         Style           =   1  'Graphical
         TabIndex        =   40
         Top             =   4440
         Width           =   375
      End
      Begin VB.CommandButton cmd 
         Height          =   375
         Index           =   9
         Left            =   -72330
         Picture         =   "frmreportnomina.frx":017E
         Style           =   1  'Graphical
         TabIndex        =   39
         Top             =   5280
         Width           =   375
      End
      Begin VB.ListBox List1 
         Height          =   3960
         Index           =   4
         ItemData        =   "frmreportnomina.frx":0253
         Left            =   -71715
         List            =   "frmreportnomina.frx":0255
         Sorted          =   -1  'True
         TabIndex        =   38
         Top             =   3195
         Width           =   2190
      End
      Begin VB.ListBox List1 
         Height          =   3960
         Index           =   5
         ItemData        =   "frmreportnomina.frx":0257
         Left            =   -74745
         List            =   "frmreportnomina.frx":0259
         Sorted          =   -1  'True
         TabIndex        =   37
         Top             =   3150
         Width           =   2190
      End
      Begin VB.Frame Frame1 
         Caption         =   "Rangos"
         Height          =   2115
         Index           =   1
         Left            =   -68925
         TabIndex        =   24
         Top             =   675
         Width           =   5415
         Begin VB.OptionButton Option1 
            Caption         =   "Todos"
            Height          =   210
            Index           =   3
            Left            =   225
            TabIndex        =   32
            Top             =   435
            Width           =   960
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Rango( Ejemplo: Quincena - Mes - Año)"
            Height          =   195
            Index           =   1
            Left            =   240
            TabIndex        =   31
            Top             =   705
            Width           =   4935
         End
         Begin VB.ComboBox cmb 
            Height          =   315
            Index           =   11
            ItemData        =   "frmreportnomina.frx":025B
            Left            =   1230
            List            =   "frmreportnomina.frx":0265
            Style           =   2  'Dropdown List
            TabIndex        =   30
            Top             =   1065
            Width           =   735
         End
         Begin VB.ComboBox cmb 
            Height          =   315
            Index           =   10
            ItemData        =   "frmreportnomina.frx":026F
            Left            =   2160
            List            =   "frmreportnomina.frx":0297
            Style           =   2  'Dropdown List
            TabIndex        =   29
            Top             =   1065
            Width           =   1695
         End
         Begin VB.ComboBox cmb 
            Height          =   315
            Index           =   9
            ItemData        =   "frmreportnomina.frx":0300
            Left            =   4050
            List            =   "frmreportnomina.frx":0313
            Style           =   2  'Dropdown List
            TabIndex        =   28
            Top             =   1065
            Width           =   1095
         End
         Begin VB.ComboBox cmb 
            Height          =   315
            Index           =   8
            ItemData        =   "frmreportnomina.frx":0335
            Left            =   1230
            List            =   "frmreportnomina.frx":033F
            Style           =   2  'Dropdown List
            TabIndex        =   27
            Top             =   1425
            Width           =   735
         End
         Begin VB.ComboBox cmb 
            Height          =   315
            Index           =   7
            ItemData        =   "frmreportnomina.frx":0349
            Left            =   2160
            List            =   "frmreportnomina.frx":0371
            Style           =   2  'Dropdown List
            TabIndex        =   26
            Top             =   1425
            Width           =   1695
         End
         Begin VB.ComboBox cmb 
            Height          =   315
            Index           =   6
            ItemData        =   "frmreportnomina.frx":03DA
            Left            =   4050
            List            =   "frmreportnomina.frx":03ED
            Style           =   2  'Dropdown List
            TabIndex        =   25
            Top             =   1425
            Width           =   1095
         End
         Begin VB.Label lbl 
            Caption         =   "Desde:"
            Height          =   255
            Index           =   3
            Left            =   615
            TabIndex        =   34
            Top             =   1065
            Width           =   615
         End
         Begin VB.Label lbl 
            Caption         =   "Hasta:"
            Height          =   255
            Index           =   2
            Left            =   615
            TabIndex        =   33
            Top             =   1455
            Width           =   615
         End
         Begin VB.Line Line1 
            Index           =   1
            X1              =   1995
            X2              =   2115
            Y1              =   1230
            Y2              =   1230
         End
         Begin VB.Line Line2 
            Index           =   1
            X1              =   1995
            X2              =   2115
            Y1              =   1575
            Y2              =   1575
         End
         Begin VB.Line Line3 
            Index           =   1
            X1              =   3885
            X2              =   4005
            Y1              =   1215
            Y2              =   1215
         End
         Begin VB.Line Line4 
            Index           =   1
            X1              =   3885
            X2              =   4005
            Y1              =   1560
            Y2              =   1560
         End
      End
      Begin VB.CommandButton cmd 
         Height          =   375
         Index           =   7
         Left            =   -72345
         Picture         =   "frmreportnomina.frx":040F
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   1080
         Width           =   375
      End
      Begin VB.ListBox List1 
         Height          =   840
         Index           =   3
         ItemData        =   "frmreportnomina.frx":0555
         Left            =   -71745
         List            =   "frmreportnomina.frx":0557
         Sorted          =   -1  'True
         TabIndex        =   13
         Top             =   1080
         Width           =   2190
      End
      Begin VB.ListBox List1 
         Height          =   840
         Index           =   2
         ItemData        =   "frmreportnomina.frx":0559
         Left            =   -74715
         List            =   "frmreportnomina.frx":0569
         Sorted          =   -1  'True
         TabIndex        =   12
         Top             =   1095
         Width           =   2190
      End
      Begin VB.CommandButton cmd 
         Height          =   375
         Index           =   6
         Left            =   -72345
         Picture         =   "frmreportnomina.frx":05B4
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   1545
         Width           =   375
      End
      Begin VB.CommandButton cmd 
         Height          =   615
         Index           =   5
         Left            =   -64875
         Picture         =   "frmreportnomina.frx":0689
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   3735
         Width           =   735
      End
      Begin VB.CommandButton cmd 
         Height          =   615
         Index           =   4
         Left            =   -68250
         Picture         =   "frmreportnomina.frx":0993
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   3735
         Width           =   735
      End
      Begin VB.CommandButton cmd 
         Height          =   375
         Index           =   1
         Left            =   2670
         Picture         =   "frmreportnomina.frx":0C9D
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   1230
         Width           =   375
      End
      Begin VB.ListBox List1 
         Height          =   1425
         Index           =   1
         ItemData        =   "frmreportnomina.frx":0DE3
         Left            =   3285
         List            =   "frmreportnomina.frx":0DE5
         Sorted          =   -1  'True
         TabIndex        =   7
         Top             =   1035
         Width           =   2190
      End
      Begin VB.CommandButton cmd 
         Height          =   615
         Index           =   3
         Left            =   6750
         Picture         =   "frmreportnomina.frx":0DE7
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   3735
         Width           =   735
      End
      Begin VB.CommandButton cmd 
         Height          =   615
         Index           =   2
         Left            =   10125
         Picture         =   "frmreportnomina.frx":10F1
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   3735
         Width           =   735
      End
      Begin VB.CommandButton cmd 
         Height          =   375
         Index           =   0
         Left            =   2670
         Picture         =   "frmreportnomina.frx":13FB
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   1950
         Width           =   375
      End
      Begin VB.Frame Frame1 
         Caption         =   "Rangos"
         Height          =   2115
         Index           =   0
         Left            =   6075
         TabIndex        =   2
         Top             =   675
         Width           =   5415
         Begin VB.ComboBox cmb 
            Height          =   315
            Index           =   5
            ItemData        =   "frmreportnomina.frx":14D0
            Left            =   4050
            List            =   "frmreportnomina.frx":14E3
            Style           =   2  'Dropdown List
            TabIndex        =   23
            Top             =   1425
            Width           =   1095
         End
         Begin VB.ComboBox cmb 
            Height          =   315
            Index           =   4
            ItemData        =   "frmreportnomina.frx":1505
            Left            =   2175
            List            =   "frmreportnomina.frx":152D
            Style           =   2  'Dropdown List
            TabIndex        =   22
            Top             =   1425
            Width           =   1695
         End
         Begin VB.ComboBox cmb 
            Height          =   315
            Index           =   3
            ItemData        =   "frmreportnomina.frx":1596
            Left            =   1230
            List            =   "frmreportnomina.frx":15A0
            Style           =   2  'Dropdown List
            TabIndex        =   21
            Top             =   1425
            Width           =   735
         End
         Begin VB.ComboBox cmb 
            Height          =   315
            Index           =   2
            ItemData        =   "frmreportnomina.frx":15AA
            Left            =   4050
            List            =   "frmreportnomina.frx":15BD
            Style           =   2  'Dropdown List
            TabIndex        =   20
            Top             =   1065
            Width           =   1095
         End
         Begin VB.ComboBox cmb 
            Height          =   315
            Index           =   1
            ItemData        =   "frmreportnomina.frx":15DF
            Left            =   2160
            List            =   "frmreportnomina.frx":1607
            Style           =   2  'Dropdown List
            TabIndex        =   19
            Top             =   1065
            Width           =   1695
         End
         Begin VB.ComboBox cmb 
            Height          =   315
            Index           =   0
            ItemData        =   "frmreportnomina.frx":1670
            Left            =   1230
            List            =   "frmreportnomina.frx":167A
            Style           =   2  'Dropdown List
            TabIndex        =   18
            Top             =   1065
            Width           =   735
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Rango( Ejemplo: Quincena - Mes - Año)"
            Height          =   195
            Index           =   0
            Left            =   240
            TabIndex        =   17
            Top             =   705
            Width           =   4935
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Todos"
            Height          =   195
            Index           =   2
            Left            =   225
            TabIndex        =   3
            Top             =   435
            Width           =   975
         End
         Begin VB.Line Line4 
            Index           =   0
            X1              =   3885
            X2              =   4005
            Y1              =   1560
            Y2              =   1560
         End
         Begin VB.Line Line3 
            Index           =   0
            X1              =   3885
            X2              =   4005
            Y1              =   1215
            Y2              =   1215
         End
         Begin VB.Line Line2 
            Index           =   0
            X1              =   1995
            X2              =   2115
            Y1              =   1575
            Y2              =   1575
         End
         Begin VB.Line Line1 
            Index           =   0
            X1              =   1995
            X2              =   2115
            Y1              =   1230
            Y2              =   1230
         End
         Begin VB.Label lbl 
            Caption         =   "Hasta:"
            Height          =   255
            Index           =   1
            Left            =   615
            TabIndex        =   16
            Top             =   1455
            Width           =   615
         End
         Begin VB.Label lbl 
            Caption         =   "Desde:"
            Height          =   255
            Index           =   0
            Left            =   615
            TabIndex        =   15
            Top             =   1065
            Width           =   615
         End
      End
      Begin VB.ListBox List1 
         Height          =   1425
         Index           =   0
         ItemData        =   "frmreportnomina.frx":1684
         Left            =   285
         List            =   "frmreportnomina.frx":169D
         Sorted          =   -1  'True
         TabIndex        =   1
         Top             =   1065
         Width           =   2190
      End
      Begin VB.Label lbl 
         Caption         =   "Inmuebles"
         Height          =   255
         Index           =   5
         Left            =   -73980
         TabIndex        =   41
         Top             =   2835
         Width           =   765
      End
      Begin VB.Label lbl 
         Caption         =   "Reportes"
         Height          =   255
         Index           =   6
         Left            =   -73980
         TabIndex        =   36
         Top             =   690
         Width           =   690
      End
      Begin VB.Label lbl 
         Caption         =   "Reportes"
         Height          =   255
         Index           =   4
         Left            =   1020
         TabIndex        =   35
         Top             =   690
         Width           =   690
      End
   End
End
Attribute VB_Name = "frmreportnomina"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmd_Click(Index As Integer)
   Select Case Index
     Case 0   'subir
          Dim X1 As Integer
          For X1 = 0 To List1(1).ListCount - 1
          List1(0).AddItem List1(1).List(X1)
          Next
          List1(1).Clear
     Case 1
          Dim X2 As Integer
          For X2 = 0 To List1(0).ListCount - 1
          List1(1).AddItem List1(0).List(X2)
          Next
          List1(0).Clear
     Case 2
          Unload Me
     Case 3
          'Reporte General
          If Validar() = False Then Call printer_print
     Case 4
          'Reportes especifico
          If Validar() = False Then Call printer_print
     Case 5
          Unload Me
     Case 6
          Dim X3 As Integer
          For X3 = 0 To List1(3).ListCount - 1
          List1(2).AddItem List1(3).List(X3)
          Next
          List1(3).Clear
     Case 7
          Dim X4 As Integer
          For X4 = 0 To List1(2).ListCount - 1
          List1(3).AddItem List1(2).List(X4)
          Next
          List1(2).Clear
     Case 8
          Dim X6 As Integer
          For X6 = 0 To List1(5).ListCount - 1
          List1(4).AddItem List1(5).List(X6)
          Next
          List1(5).Clear
     Case 9
          Dim X5 As Integer
          For X5 = 0 To List1(4).ListCount - 1
          List1(5).AddItem List1(4).List(X5)
          Next
          List1(4).Clear
   End Select
End Sub
 


Private Sub Form_Load()

lbl(0).Visible = False
lbl(1).Visible = False
cmb(0).Visible = False
cmb(1).Visible = False
cmb(2).Visible = False
cmb(3).Visible = False
cmb(4).Visible = False
cmb(5).Visible = False
Line1(0).Visible = False
Line2(0).Visible = False
Line3(0).Visible = False
Line4(0).Visible = False
lbl(3).Visible = False
lbl(2).Visible = False
cmb(8).Visible = False
cmb(11).Visible = False
cmb(7).Visible = False
cmb(10).Visible = False
cmb(6).Visible = False
cmb(9).Visible = False
Line1(1).Visible = False
Line2(1).Visible = False
Line3(1).Visible = False
Line4(1).Visible = False

Set rstnomina = New ADODB.Recordset
Dim sql As String
Dim X As Integer
Set rstnomina = New ADODB.Recordset
sql = "SELECT * FROM INMUEBLE"
rstnomina.Open sql, cnnConexion, adOpenStatic, adLockReadOnly, adCmdText
If Not rstnomina.EOF And Not rstnomina.BOF Then
        rstnomina.MoveFirst
   Do
     List1(5).AddItem (rstnomina("codinm"))
     rstnomina.MoveNext
   Loop Until rstnomina.EOF
End If
    rstnomina.Close
Set rstnomina = Nothing
cmb(2).Clear
cmb(5).Clear
cmb(6).Clear
cmb(9).Clear
For I = 0 To 5
    cmb(2).AddItem Year(Date) - I
    cmb(5).AddItem Year(Date) - I
    cmb(6).AddItem Year(Date) - I
    cmb(9).AddItem Year(Date) - I
Next
End Sub



Private Sub List1_Click(Index As Integer)
  Select Case Index
    Case 0
         Dim X1 As String, Y1 As Integer
             X1 = List1(0).Text
             Y1 = List1(0).ListIndex
          If List1(0).ListIndex >= 0 Then
             List1(1).AddItem X1
             List1(0).RemoveItem Y1
          End If
    Case 1
         Dim X2 As String, Y2 As Integer
             X2 = List1(1).Text
             Y2 = List1(1).ListIndex
          If List1(1).ListIndex >= 0 Then
             List1(0).AddItem X2
             List1(1).RemoveItem Y2
          End If
    Case 2
         Dim X3 As String, Y3 As Integer
             X3 = List1(2).Text
             Y3 = List1(2).ListIndex
          If List1(2).ListIndex >= 0 Then
             List1(3).AddItem X3
             List1(2).RemoveItem Y3
          End If
    Case 3
         Dim X4 As String, Y4 As Integer
             X4 = List1(3).Text
             Y4 = List1(3).ListIndex
          If List1(3).ListIndex >= 0 Then
             List1(2).AddItem X4
             List1(3).RemoveItem Y4
          End If
    Case 4
         Dim X5 As String, Y5 As Integer
             X5 = List1(4).Text
             Y5 = List1(4).ListIndex
          If List1(4).ListIndex >= 0 Then
             List1(5).AddItem X5
             List1(4).RemoveItem Y5
          End If
          
    Case 5
         Dim X6 As String, Y6 As Integer
             X6 = List1(5).Text
             Y6 = List1(5).ListIndex
          If List1(5).ListIndex >= 0 Then
             List1(4).AddItem X6
             List1(5).RemoveItem Y6
          End If
          
  End Select
End Sub

Private Sub Option1_Click(Index As Integer)
  Select Case Index
    Case 0
         If Option1(0).Value = True Then
            lbl(0).Visible = True
            lbl(1).Visible = True
            cmb(0).Visible = True
            cmb(1).Visible = True
            cmb(2).Visible = True
            cmb(3).Visible = True
            cmb(4).Visible = True
            cmb(5).Visible = True
            Line1(0).Visible = True
            Line2(0).Visible = True
            Line3(0).Visible = True
            Line4(0).Visible = True
         End If
    
    Case 1
         If Option1(1).Value = True Then
            lbl(3).Visible = True
            lbl(2).Visible = True
            cmb(8).Visible = True
            cmb(11).Visible = True
            cmb(7).Visible = True
            cmb(10).Visible = True
            cmb(6).Visible = True
            cmb(9).Visible = True
            Line1(1).Visible = True
            Line2(1).Visible = True
            Line3(1).Visible = True
            Line4(1).Visible = True
         End If
    Case 2
         If Option1(2).Value = True Then
            lbl(0).Visible = False
            lbl(1).Visible = False
            cmb(0).Visible = False
            cmb(1).Visible = False
            cmb(2).Visible = False
            cmb(3).Visible = False
            cmb(4).Visible = False
            cmb(5).Visible = False
            Line1(0).Visible = False
            Line2(0).Visible = False
            Line3(0).Visible = False
            Line4(0).Visible = False
         End If
    Case 3
         If Option1(3).Value = True Then
            lbl(3).Visible = False
            lbl(2).Visible = False
            cmb(8).Visible = False
            cmb(11).Visible = False
            cmb(7).Visible = False
            cmb(10).Visible = False
            cmb(6).Visible = False
            cmb(9).Visible = False
            Line1(1).Visible = False
            Line2(1).Visible = False
            Line3(1).Visible = False
            Line4(1).Visible = False
         End If
  End Select
End Sub

 Private Sub printer_print()
 
 Dim Fecha1, j As Date
 Dim Fecha2 As Date
 Dim Nom() As Long, P As String, m As String, qui As String
 Dim Z As Integer, u As Integer, s As Integer, Mes As String
 Dim q As Integer, n As String, l As Integer, ano As String
 Dim X As Integer, K As String, r As String, msn As String
 Dim Nreport As String, T As Date, I As Integer
 Dim A As Integer, B As Integer, C As Integer, D As Integer
 Dim E As Integer, F As Integer, G As Integer, H As Integer
 Dim W As Integer, errLocal As Integer
 Dim rpReporte As ctlReport
 If SSTab1.tab = 0 Then
    'Reporte general
     A = 0
     B = 1
     C = 2
     D = 3
     E = 4
     F = 5
     G = 2
     H = 0
     Z = 1
     
 Else
    'Reporte especifico
     A = 11
     B = 10
     C = 9
     D = 8
     E = 7
     F = 6
     Z = 3
     G = 3
     H = 1
 End If
 If Option1(H).Value = True Then
    Fecha1 = cmb(A) & "/" & cmb(B) & "/" & cmb(C)
    Fecha2 = cmb(D) & "/" & cmb(E) & "/" & cmb(F)
    q = (DateDiff("m", Fecha1, Fecha2) + 1) * 2
 
    If cmb(A) = 2 Then q = q - 1
    If cmb(D) = 1 Then q = q - 1
    If Fecha1 > Fecha2 Then
       MsgBox "Chequee el Rago Seleccionado", vbCritical
       GoTo salto:
    End If
    ReDim Nom(q - 1)
    Nom(0) = Format(Fecha1, "dmmyyyy")
    X = 1
    If Fecha1 = Fecha2 Then
       GoTo paso
    End If
paso1:
 
    Do
      Nom(X) = IIf(Day(Fecha1) = 1, 2 & Format(Fecha1, "mmyyyy"), 1 & Format(DateAdd("m", 1, Fecha1), "mmyyyy"))
      X = X + 1:  Fecha1 = IIf(Day(Fecha1) = 1, 2 & Format(Fecha1, "/mm/yyyy"), 1 & Format(DateAdd("m", 1, Fecha1), "/mm/yyyy"))
    Loop Until Fecha1 = Fecha2
paso:
 
    For X = 0 To UBound(Nom)
        If SSTab1.tab = 0 Then
          'Reporte General
           For q = 0 To List1(Z).ListCount - 1
             Select Case List1(Z).List(q)
                Case "Cuadre de Nomina"
                     Nreport = "CN" & Nom(X) & ".RPT"
                
                Case "Nomina General"
                      Nreport = "RG" & Nom(X) & ".RPT"
                
                Case "Nomina General Mensual"
                      Nreport = "GE" & Nom(X) & ".RPT"
            
                Case "Novedades de Nomina"
                      Call Printer_Nom_Nov(crPantalla, Nom(X))
            
                Case "Pagos en Cheque"
                      Nreport = "CH" & Nom(X) & ".RPT"
            
                Case "Nomina Vs Facturacion"
                      Nreport = "FA" & Nom(X) & ".RPT"
                    
                Case "Recibo de Pago"
                        Nreport = "RP" & Nom(X) & ".RPT"
                        Call printer_recibo_pago(Nom(X), False, vbYes, crPantalla)
                        Exit For
             End Select
            Set rpReporte = New ctlReport
            With rpReporte
                  .Reporte = gcPath & "\nomina\" & Nreport
                  .Salida = crPantalla
                   errLocal = .Imprimir
                 If errLocal <> 0 Then
                    qui = Left(Nom(X), 1)
                    Mes = (MonthName(Mid(Nom(X), 2, 2)))
                    ano = Right(Nom(X), 4)
                    msn = "De la Quincena  " & qui & "  De  " & Mes & "  Del  " & ano & " "
                    If errLocal <> 0 Then
                        MsgBox Err.Description, _
                        vbCritical, App.ProductName
                    End If
                End If
            End With
            Set rpReporte = Nothing
    Next
    'Next
        Else
            'Reporte especifico
            'For X = 0 To UBound(Nom)
            'For q = 0 To List1(Z).ListCount - 1

          For W = 0 To List1(3).ListCount - 1
              For u = 0 To List1(4).ListCount - 1
                  n = Right(List1(4).List(u), 2)
                  Z = 3
                  '
                  Select Case List1(Z).List(W)
                  
                    Case "Nomina Vs Facturacion"
                        Nreport = "FA" & n & Nom(X) & ".rpt"
              
                    Case "Recibo de Pago"
                        Nreport = "RP" & n & Nom(X) & ".RPT"
             
                    Case "Carta al Banco"
                        Nreport = "CB" & n & Nom(X) & ".RPT"
             
                    Case "Nomina General"
                        Nreport = "RG" & List1(4).List(u) & Nom(X) & ".RPT"
           
                  End Select
                    '
                    '
                Set rpReporte = New ctlReport
                  With rpReporte
                      .Reporte = gcPath & "\nomina\" & Nreport
                      .Salida = crImpresora
                       .Imprimir
'                    If errLocal <> 0 Then
'                        errLocal = 0
'                       qui = Left(Nom(X), 1)
'                       Mes = (MonthName(Mid(Nom(X), 2, 2)))
'                       ano = Right(Nom(X), 4)
'                       msn = " De la Quincena  " & qui & "  De  " & Mes & "  Del  " & ano & " "
'                       MsgBox "El  reporte  " & List1(3).List(W) & "  Del Inmueble  " & List1(4).List(u) & " " & msn & "  no se encuentra   ", vbCritical, .LastErrorString
'                    End If
                  End With
                  Set Reporte = Nothing
              Next
          Next
         'Next
   
        End If
    Next
    'Next
'pasa:
'    Next
 End If
  If Option1(2).Value = True Then
       'General (cuando se escoje todo)
        Set rsttodo = New ADODB.Recordset
        Dim sql As String
        Set rsttodo = New ADODB.Recordset
        sql = "SELECT *FROM nom_inf"
        rsttodo.Open sql, cnnConexion, adOpenStatic, adLockReadOnly, adCmdText

    If Not rsttodo.EOF And Not rsttodo.BOF Then
        
        For W = 0 To List1(Z).ListCount - 1
            Select Case List1(Z).List(W)
                   Case "Cuadre de Nomina"
                         r = "CB"
                    
                   Case "Nomina General"
                         r = "RG"
                 
                   Case "Nomina General Mensual"
                         r = "GE"
                  
                   Case "Novedades de Nomina"
                         rsttodo.MoveFirst
                         K = (rsttodo("IDNomina"))
                         Call Printer_Nom_Nov(crPantalla, K)
            
                   Case "Pago en Cheque"
                         r = "CH"
            
                   Case "Nomina Vs Facturacion"
                         r = "FA"
                  
                   Case "Recibo de Pago"
                         r = "RP"
                  
            End Select
               ' For q = 0 To List1(4).ListCount - 1
            rsttodo.MoveFirst
            Do
               m = (rsttodo("fecha"))
               K = (rsttodo("IDnomina"))
               repotg = r & K & ".RPT"
               Set rpReporte = New ctlReport
               With rpReporte
                    .Reporte = gcPath & "\Nomina\" & repotg
                    .Salida = crPantalla
                     .Imprimir
'                  If errLocal <> 0 Then
'                     qui = Left(K, 1)
'                     Mes = (MonthName(Mid(K, 2, 2)))
'                     ano = Right(K, 4)
'                     msn = "De la Quincena  " & qui & "  De  " & Mes & "  Del  " & ano & " "
'                     MsgBox "El  reporte " & List1(Z).List(W) & "  " & msn & " no se encuentra   ", vbCritical, .LastErrorString
'                  End If
               End With
               Set rpReporte = Nothing
               rsttodo.MoveNext
            Loop Until rsttodo.EOF
      ' Next
        Next
    End If
        rsttodo.Close
        Set rsttodo = Nothing
 End If
 
 
   If Option1(3).Value = True Then
      'especifico(cuando se escoje todo)
       Set rsttodo = New ADODB.Recordset
       Dim sql1 As String
       Set rsttodo = New ADODB.Recordset
       sql1 = "SELECT *FROM nom_inf"
       rsttodo.Open sql1, cnnConexion, adOpenStatic, adLockReadOnly, adCmdText

     If Not rsttodo.EOF And Not rsttodo.BOF Then
          
        For W = 0 To List1(3).ListCount - 1
            For s = 0 To List1(4).ListCount - 1
               'p = List1(4).list(q)
                n = Right(List1(4).List(q), 2)
                    Select Case List1(3).List(W)
                           Case "Nomina Vs Facturacion"
                                 r = "FA"
            
                           Case "Recibo de Pago"
                                 r = "RP"
            
                           Case "Carta al Banco"
                                 r = "CB"
            
                           Case "Nomina General"
                                 r = "RG"
          
                    End Select
     
                    'For q = 0 To List1(4).ListCount - 1
                    rsttodo.MoveFirst
               Do
                  K = (rsttodo("IDnomina"))
                  repotg = r & n & K & ".RPT"
                  Set rpReporte = New ctlReport
                  With rpReporte
                       .Reporte = gcPath & "\Nomina\" & repotg
                       .Salida = crptToWindow
                        .Imprimir
'                        If errLocal <> 0 Then
'                           qui = Left(K, 1)
'                           Mes = (MonthName(Mid(K, 2, 2)))
'                           ano = Right(K, 4)
'                           msn = " De la Quincena  " & qui & "  De  " & Mes & "  Del  " & ano & " "
'                           MsgBox "El  reporte  " & List1(3).List(W) & "  Del  Inmueble  " & List1(4).List(s) & " " & msn & " no se encuentra   ", vbCritical, .LastErrorString
'                        End If
                  End With
                  Set rpReporte = Nothing
        
                   rsttodo.MoveNext
               Loop Until rsttodo.EOF
               'Next
            Next
        Next

     End If
     rsttodo.Close
     Set rsttodo = Nothing
     'Next
   End If
salto:
 End Sub
 
  
Private Function Validar() As Boolean
      If Fecha1 > Fecha2 Then
         MsgBox "Chequee las fecha"
      End If
      If SSTab1.tab = 0 Then
         'Reporte General
            If List1(1).List(0) = "" Then
               MsgBox "Falta seleccionar un tipo de reporte", vbCritical
            End If
         A = 0
         B = 1
         C = 2
         D = 3
         E = 4
         F = 5
         G = 2
         H = 0
         Z = 1

      Else
          'Reporte especifico
             If List1(3).List(0) = "" Then
                MsgBox "Falta seleccionar Un tipo de reporte", vbCritical
             End If
             If List1(4).List(0) = "" Then
                MsgBox "Falta seleccionar Inmueble", vbCritical
             End If
          A = 11
          B = 10
          C = 9
          D = 8
          E = 7
          F = 6
          G = 3
          H = 1
          Z = 3
      End If

       
 If Option1(G).Value = False Then
        If cmb(A) = "" Or cmb(B) = "" Or cmb(C) = "" Or cmb(D) = "" Or cmb(E) = "" Or cmb(F) = "" Then
           Validar = MsgBox("Falta algun dato, vea que las fecha este completa")
        End If
        If cmb(C) > cmb(F) Then
           Validar = MsgBox("Fecha incorrecto,Revise los años")
        Else
            If cmb(C) = cmb(F) Then
                  If cmb(B) > cmb(E) And cmb(C) > cmb(F) Then
                     Validar = MsgBox("Fecha incorecta,El mese de donde se quiere el reporte no puede ser mayor a el mese hasta donde se quiere")
                  End If
               If cmb(C) < cmb(F) Then
                  MsgBox "El peridodo selecionado es un año a otro¿Desea continuar?", vbYesNo
               End If
            End If
     
            If cmb(B) = cmb(E) Then
               If cmb(A) > cmb(D) And cmb(C) >= cmb(F) Then
                  Validar = MsgBox("Fecha incorrecta, Revise el periodo seleccionado")
               End If
            End If
        End If
 End If
End Function
