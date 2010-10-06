VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmVacacion 
   Caption         =   "Vacaciones"
   ClientHeight    =   8490
   ClientLeft      =   -5040
   ClientTop       =   555
   ClientWidth     =   11880
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8490
   ScaleWidth      =   11880
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmd 
      Caption         =   "&Suplencias"
      Height          =   495
      Index           =   3
      Left            =   5880
      TabIndex        =   73
      Tag             =   "0489013380004950121503660058800049501215"
      Top             =   3660
      Width           =   1215
   End
   Begin VB.TextBox Txt 
      Alignment       =   2  'Center
      Height          =   315
      Index           =   12
      Left            =   4440
      TabIndex        =   72
      Tag             =   "0415500525003150069003765004650031500690"
      Top             =   6960
      Width           =   690
   End
   Begin VB.CommandButton cmd 
      Caption         =   "&Procesar"
      Enabled         =   0   'False
      Height          =   495
      Index           =   2
      Left            =   8535
      TabIndex        =   37
      Tag             =   "0384013380004950121503660085350049501215"
      Top             =   3660
      Width           =   1215
   End
   Begin VB.CommandButton cmd 
      Caption         =   "&Cerrar"
      Height          =   495
      Index           =   1
      Left            =   9810
      TabIndex        =   36
      Tag             =   "0436513380004950121503660098100049501215"
      Top             =   3660
      Width           =   1215
   End
   Begin VB.CommandButton cmd 
      Caption         =   "C&alcular"
      Enabled         =   0   'False
      Height          =   495
      Index           =   0
      Left            =   7260
      TabIndex        =   35
      Tag             =   "0333013380004950121503660072600049501215"
      Top             =   3660
      Width           =   1215
   End
   Begin VB.TextBox Txt 
      Alignment       =   2  'Center
      Height          =   315
      Index           =   4
      Left            =   570
      Locked          =   -1  'True
      TabIndex        =   33
      Tag             =   "0348000668003150121503090005700031501095"
      Top             =   3090
      Width           =   1095
   End
   Begin VB.TextBox Txt 
      Alignment       =   2  'Center
      Height          =   315
      Index           =   7
      Left            =   5775
      Locked          =   -1  'True
      TabIndex        =   32
      Tag             =   "0348007395003150121503090057750031501215"
      Top             =   3090
      Width           =   1215
   End
   Begin VB.TextBox Txt 
      Alignment       =   2  'Center
      Height          =   315
      Index           =   11
      Left            =   3255
      Locked          =   -1  'True
      TabIndex        =   26
      Tag             =   "0348004200003150069003090032550031500690"
      Top             =   3090
      Width           =   690
   End
   Begin VB.TextBox Txt 
      Alignment       =   2  'Center
      Height          =   315
      Index           =   10
      Left            =   9690
      TabIndex        =   24
      Tag             =   "0348011858003150121503090096900031501215"
      Top             =   3090
      Width           =   1215
   End
   Begin VB.TextBox Txt 
      Alignment       =   2  'Center
      Height          =   315
      Index           =   9
      Left            =   8175
      TabIndex        =   23
      Tag             =   "0348010373003150121503090081750031501215"
      Top             =   3090
      Width           =   1215
   End
   Begin VB.TextBox Txt 
      Alignment       =   2  'Center
      Height          =   315
      Index           =   8
      Left            =   7200
      Locked          =   -1  'True
      TabIndex        =   22
      Tag             =   "0348008910003150121503090072000031500690"
      Top             =   3090
      Width           =   690
   End
   Begin VB.TextBox Txt 
      Alignment       =   2  'Center
      Height          =   315
      Index           =   6
      Left            =   4350
      Locked          =   -1  'True
      TabIndex        =   21
      Tag             =   "0348005895003150121503090043500031501215"
      Top             =   3090
      Width           =   1215
   End
   Begin VB.TextBox Txt 
      Alignment       =   2  'Center
      Height          =   315
      Index           =   5
      Left            =   1830
      Locked          =   -1  'True
      TabIndex        =   20
      Tag             =   "0348002160003150121503090018300031501095"
      Top             =   3090
      Width           =   1095
   End
   Begin VB.TextBox Txt 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   3
      Left            =   10110
      TabIndex        =   19
      Tag             =   "0160511460003150142501500101100031501425"
      Top             =   1500
      Width           =   1425
   End
   Begin VB.TextBox Txt 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   2
      Left            =   7200
      TabIndex        =   18
      Tag             =   "0160508925003150142501485072000031501425"
      Top             =   1485
      Width           =   1425
   End
   Begin VB.TextBox Txt 
      Height          =   315
      Index           =   1
      Left            =   1455
      TabIndex        =   17
      Tag             =   "0160503900003150379501530014550031503795"
      Top             =   1530
      Width           =   3795
   End
   Begin VB.TextBox Txt 
      Height          =   315
      Index           =   0
      Left            =   10110
      TabIndex        =   16
      Tag             =   "0160501455003150121500383101100031501215"
      Top             =   383
      Width           =   1215
   End
   Begin MSDataListLib.DataCombo Dtc 
      Bindings        =   "frmVacacion.frx":0000
      Height          =   315
      Index           =   0
      Left            =   1455
      TabIndex        =   12
      Tag             =   "0067501455003150121500368014550031501215"
      Top             =   368
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   556
      _Version        =   393216
      ListField       =   "CodInm"
      BoundColumn     =   "Nombre"
      Text            =   ""
   End
   Begin MSDataListLib.DataCombo Dtc 
      Bindings        =   "frmVacacion.frx":0015
      Height          =   315
      Index           =   1
      Left            =   1455
      TabIndex        =   13
      Tag             =   "0067503900003150379500915014550031503795"
      Top             =   915
      Width           =   3795
      _ExtentX        =   6694
      _ExtentY        =   556
      _Version        =   393216
      ListField       =   "Nombre"
      BoundColumn     =   "CodInm"
      Text            =   ""
   End
   Begin MSDataListLib.DataCombo Dtc 
      Bindings        =   "frmVacacion.frx":002A
      Height          =   315
      Index           =   2
      Left            =   7200
      TabIndex        =   14
      Tag             =   "0069008925003150121500368072000031501215"
      Top             =   368
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   556
      _Version        =   393216
      ListField       =   "CodEmp"
      BoundColumn     =   "Empleado"
      Text            =   ""
   End
   Begin MSDataListLib.DataCombo Dtc 
      Bindings        =   "frmVacacion.frx":003F
      Height          =   315
      Index           =   3
      Left            =   7200
      TabIndex        =   15
      Tag             =   "0069011460003150324000915072000031504125"
      Top             =   915
      Width           =   4125
      _ExtentX        =   7276
      _ExtentY        =   556
      _Version        =   393216
      ListField       =   "Empleado"
      BoundColumn     =   "CodEmp"
      Text            =   ""
   End
   Begin MSAdodcLib.Adodc Ado 
      Height          =   330
      Index           =   3
      Left            =   6210
      Top             =   3765
      Visible         =   0   'False
      Width           =   2025
      _ExtentX        =   3572
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
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
      Caption         =   "Emp"
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
   Begin MSAdodcLib.Adodc Ado 
      Height          =   330
      Index           =   2
      Left            =   4185
      Top             =   3765
      Visible         =   0   'False
      Width           =   2025
      _ExtentX        =   3572
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
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
      Caption         =   "Cod Emp"
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
   Begin MSAdodcLib.Adodc Ado 
      Height          =   330
      Index           =   1
      Left            =   2430
      Top             =   3765
      Visible         =   0   'False
      Width           =   1755
      _ExtentX        =   3096
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
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
      Caption         =   "Inm"
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
   Begin MSAdodcLib.Adodc Ado 
      Height          =   330
      Index           =   0
      Left            =   480
      Top             =   1875
      Visible         =   0   'False
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
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
      Caption         =   "Cod Inm"
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
   Begin VB.Label lbl 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Ingrese los dias feriados, y presione enter para calcular:"
      ForeColor       =   &H80000008&
      Height          =   285
      Index           =   26
      Left            =   480
      TabIndex        =   74
      Tag             =   "0313500525002850150002730004800028501275"
      Top             =   6975
      Width           =   3915
   End
   Begin VB.Label lbl 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "TOTAL"
      ForeColor       =   &H80000008&
      Height          =   285
      Index           =   56
      Left            =   10320
      TabIndex        =   69
      Tag             =   "0535511895002850123004815103200028501230"
      Top             =   4665
      Width           =   1230
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "NETO A PAGAR"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   285
      Index           =   55
      Left            =   7560
      TabIndex        =   68
      Tag             =   "0807009180002850234006975075600028502340"
      Top             =   6975
      Width           =   2340
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "0,00"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   285
      Index           =   54
      Left            =   10320
      TabIndex        =   67
      Tag             =   "0807011835002850123006975103200028501230"
      Top             =   6975
      Width           =   1230
   End
   Begin VB.Line Line1 
      Index           =   4
      Tag             =   "1171511715052350837010005100050459007260"
      X1              =   10005
      X2              =   10005
      Y1              =   4590
      Y2              =   7260
   End
   Begin VB.Shape sha 
      BorderColor     =   &H00404040&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   450
      Index           =   16
      Left            =   6450
      Tag             =   "0793507245004500595506840064500045005190"
      Top             =   6840
      Width           =   5190
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "0,00"
      ForeColor       =   &H80000008&
      Height          =   285
      Index           =   53
      Left            =   10260
      TabIndex        =   66
      Tag             =   "0763511835002850123006585102600028501230"
      Top             =   6585
      Width           =   1230
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "TOTAL DEDUCCIONES"
      ForeColor       =   &H80000008&
      Height          =   285
      Index           =   52
      Left            =   8100
      TabIndex        =   65
      Tag             =   "0763509180002850234006585081000028501830"
      Top             =   6585
      Width           =   1830
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "0,00"
      ForeColor       =   &H80000008&
      Height          =   285
      Index           =   51
      Left            =   10260
      TabIndex        =   64
      Tag             =   "0648011835002850123005640102600028501230"
      Top             =   5640
      Width           =   1230
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "0,00"
      ForeColor       =   &H80000008&
      Height          =   285
      Index           =   50
      Left            =   10260
      TabIndex        =   63
      Tag             =   "0609011835002850123005340102600028501230"
      Top             =   5340
      Width           =   1230
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "0,00"
      ForeColor       =   &H80000008&
      Height          =   285
      Index           =   49
      Left            =   10260
      TabIndex        =   62
      Tag             =   "0573011835002850123004995102600028501230"
      Top             =   4995
      Width           =   1230
   End
   Begin VB.Label lbl 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "L.P.H."
      ForeColor       =   &H80000008&
      Height          =   285
      Index           =   48
      Left            =   8160
      TabIndex        =   61
      Tag             =   "0646509120002850213005640081600028501560"
      Top             =   5640
      Width           =   1560
   End
   Begin VB.Label lbl 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "S.P.F."
      ForeColor       =   &H80000008&
      Height          =   285
      Index           =   47
      Left            =   8160
      TabIndex        =   60
      Tag             =   "0607509120002850213005340081600028501560"
      Top             =   5340
      Width           =   1560
   End
   Begin VB.Label lbl 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "S.S.O."
      ForeColor       =   &H80000008&
      Height          =   285
      Index           =   46
      Left            =   8160
      TabIndex        =   59
      Tag             =   "0571509120002850213004995081600028501560"
      Top             =   4995
      Width           =   1560
   End
   Begin VB.Label lbl 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "CONCEPTO"
      ForeColor       =   &H80000008&
      Height          =   285
      Index           =   45
      Left            =   8160
      TabIndex        =   58
      Tag             =   "0535508850002850264004665081600028501530"
      Top             =   4665
      Width           =   1530
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "0,00"
      ForeColor       =   &H80000008&
      Height          =   285
      Index           =   44
      Left            =   6600
      TabIndex        =   57
      Tag             =   "0763507440002850123006585066000028501230"
      Top             =   6585
      Width           =   1230
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "TOTAL ASIGNACIONES"
      ForeColor       =   &H80000008&
      Height          =   285
      Index           =   43
      Left            =   1050
      TabIndex        =   56
      Tag             =   "0763501050002850607506585010500028505220"
      Top             =   6585
      Width           =   5220
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "0,00"
      ForeColor       =   &H80000008&
      Height          =   285
      Index           =   41
      Left            =   6585
      TabIndex        =   55
      Tag             =   "0646507440002850123005640065850028501230"
      Top             =   5640
      Width           =   1230
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "0,00"
      ForeColor       =   &H80000008&
      Height          =   285
      Index           =   40
      Left            =   6585
      TabIndex        =   54
      Tag             =   "0607507440002850123005340065850028501230"
      Top             =   5340
      Width           =   1230
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "0,00"
      ForeColor       =   &H80000008&
      Height          =   285
      Index           =   39
      Left            =   6585
      TabIndex        =   53
      Tag             =   "0571507440002850123004995065850028501230"
      Top             =   4995
      Width           =   1230
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "0,00"
      ForeColor       =   &H80000008&
      Height          =   285
      Index           =   37
      Left            =   5100
      TabIndex        =   52
      Tag             =   "0646505880002850123005640051000028501230"
      Top             =   5640
      Width           =   1230
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "0,00"
      ForeColor       =   &H80000008&
      Height          =   285
      Index           =   36
      Left            =   5100
      TabIndex        =   51
      Tag             =   "0607505880002850123005340051000028501230"
      Top             =   5340
      Width           =   1230
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "0,00"
      ForeColor       =   &H80000008&
      Height          =   285
      Index           =   35
      Left            =   5100
      TabIndex        =   50
      Tag             =   "0571505880002850123004995051000028501230"
      Top             =   4995
      Width           =   1230
   End
   Begin VB.Label lbl 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "TOTAL"
      ForeColor       =   &H80000008&
      Height          =   285
      Index           =   34
      Left            =   6600
      TabIndex        =   49
      Tag             =   "0535507455002850123004665066000028501230"
      Top             =   4665
      Width           =   1230
   End
   Begin VB.Line Line1 
      Index           =   3
      Tag             =   "0724507245052650795006450064500459006840"
      X1              =   6450
      X2              =   6450
      Y1              =   4590
      Y2              =   6840
   End
   Begin VB.Label lbl 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "DIARIO"
      ForeColor       =   &H80000008&
      Height          =   285
      Index           =   33
      Left            =   5070
      TabIndex        =   48
      Tag             =   "0535505910002850123004665050700028501110"
      Top             =   4665
      Width           =   1110
   End
   Begin VB.Line Line1 
      Index           =   2
      Tag             =   "0571505715052500750004965049650459006465"
      X1              =   4965
      X2              =   4965
      Y1              =   4590
      Y2              =   6465
   End
   Begin VB.Label lbl 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H80000008&
      Height          =   285
      Index           =   31
      Left            =   4245
      TabIndex        =   47
      Tag             =   "0646504245002850123005640042450028500675"
      Top             =   5640
      Width           =   675
   End
   Begin VB.Label lbl 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H80000008&
      Height          =   285
      Index           =   30
      Left            =   4245
      TabIndex        =   46
      Tag             =   "0607504245002850123005340042450028500675"
      Top             =   5340
      Width           =   675
   End
   Begin VB.Label lbl 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H80000008&
      Height          =   285
      Index           =   29
      Left            =   4245
      TabIndex        =   45
      Tag             =   "0571504245002850123004995042450028500675"
      Top             =   4995
      Width           =   675
   End
   Begin VB.Label lbl 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "DIAS"
      ForeColor       =   &H80000008&
      Height          =   285
      Index           =   28
      Left            =   4245
      TabIndex        =   44
      Tag             =   "0535504245002850123004665042450028500660"
      Top             =   4665
      Width           =   660
   End
   Begin VB.Label lbl 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "CONCEPTO"
      ForeColor       =   &H80000008&
      Height          =   285
      Index           =   27
      Left            =   600
      TabIndex        =   43
      Tag             =   "0535500600002850340504665006000028503465"
      Top             =   4665
      Width           =   3405
   End
   Begin VB.Label lbl 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Bono Vacacional (L.O.T.)"
      ForeColor       =   &H80000008&
      Height          =   285
      Index           =   25
      Left            =   750
      TabIndex        =   42
      Tag             =   "0646500735002850330005670007500028503300"
      Top             =   5670
      Width           =   3300
   End
   Begin VB.Label lbl 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Días Feriados (L.O.T.)"
      ForeColor       =   &H80000008&
      Height          =   285
      Index           =   24
      Left            =   750
      TabIndex        =   41
      Tag             =   "0607500735002850330005355007500028503300"
      Top             =   5355
      Width           =   3300
   End
   Begin VB.Label lbl 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Días Hábiles (L.O.T.)"
      ForeColor       =   &H80000008&
      Height          =   285
      Index           =   23
      Left            =   750
      TabIndex        =   40
      Tag             =   "0571500735002850330005040007500028503300"
      Top             =   5040
      Width           =   3300
   End
   Begin VB.Line Line1 
      Index           =   1
      Tag             =   "0415504155052500750004170041700459006465"
      X1              =   4170
      X2              =   4170
      Y1              =   4590
      Y2              =   6465
   End
   Begin VB.Shape sha 
      DrawMode        =   9  'Not Mask Pen
      FillStyle       =   5  'Downward Diagonal
      Height          =   315
      Index           =   15
      Left            =   510
      Tag             =   "0712500510003901269006165005100031511130"
      Top             =   6165
      Width           =   11130
   End
   Begin VB.Shape sha 
      Height          =   315
      Index           =   14
      Left            =   510
      Tag             =   "0675000510003901269005865005100031511130"
      Top             =   5865
      Width           =   11130
   End
   Begin VB.Shape sha 
      Height          =   315
      Index           =   13
      Left            =   510
      Tag             =   "0637500510003901269005565005100031511130"
      Top             =   5565
      Width           =   11130
   End
   Begin VB.Shape sha 
      Height          =   315
      Index           =   12
      Left            =   510
      Tag             =   "0600000510003901269005265005100031511130"
      Top             =   5265
      Width           =   11130
   End
   Begin VB.Shape sha 
      Height          =   315
      Index           =   11
      Left            =   510
      Tag             =   "0562500510003901269004965005100031511130"
      Top             =   4965
      Width           =   11130
   End
   Begin VB.Shape sha 
      Height          =   390
      Index           =   10
      Left            =   510
      Tag             =   "0525000510003901269004590005100039011130"
      Top             =   4590
      Width           =   11130
   End
   Begin VB.Shape sha 
      Height          =   345
      Index           =   9
      Left            =   510
      Tag             =   "0480000510004651269004260005100034511130"
      Top             =   4260
      Width           =   11130
   End
   Begin VB.Line Line1 
      Index           =   0
      Tag             =   "0874508745048000795007980079800426006840"
      X1              =   7980
      X2              =   7980
      Y1              =   4260
      Y2              =   6840
   End
   Begin VB.Label lbl 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "DEDUCCIONES"
      ForeColor       =   &H80000008&
      Height          =   285
      Index           =   22
      Left            =   8790
      TabIndex        =   39
      Tag             =   "0493508790002850426004335087900028502700"
      Top             =   4335
      Width           =   2700
   End
   Begin VB.Label lbl 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "ASIGNACIONES"
      ForeColor       =   &H80000008&
      Height          =   285
      Index           =   21
      Left            =   675
      TabIndex        =   38
      Tag             =   "0493500675002850613504335006450028506135"
      Top             =   4335
      Width           =   6135
   End
   Begin VB.Shape sha 
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   2595
      Index           =   8
      Left            =   510
      Tag             =   "0480000510031501269004260005100259511130"
      Top             =   4260
      Width           =   11130
   End
   Begin VB.Label lbl 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "PERIODO CORRESPONDIENTE"
      ForeColor       =   &H80000008&
      Height          =   285
      Index           =   7
      Left            =   480
      TabIndex        =   71
      Tag             =   "0286500525002850300002460004800028502550"
      Top             =   2460
      Width           =   2550
   End
   Begin VB.Label lbl 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Desde"
      ForeColor       =   &H80000008&
      Height          =   285
      Index           =   8
      Left            =   480
      TabIndex        =   70
      Tag             =   "0313500525002850150002730004800028501275"
      Top             =   2730
      Width           =   1275
   End
   Begin VB.Shape sha 
      Height          =   495
      Index           =   7
      Left            =   9540
      Tag             =   "0340511715004950150003000095400049501500"
      Top             =   3000
      Width           =   1500
   End
   Begin VB.Shape sha 
      Height          =   495
      Index           =   6
      Left            =   8055
      Tag             =   "0340510230004950150003000080550049501500"
      Top             =   3000
      Width           =   1500
   End
   Begin VB.Shape sha 
      Height          =   495
      Index           =   5
      Left            =   7065
      Tag             =   "0340508730004950150003000070650049501005"
      Top             =   3000
      Width           =   1005
   End
   Begin VB.Shape sha 
      Height          =   495
      Index           =   4
      Left            =   5640
      Tag             =   "0340507245004950150003000056400049501440"
      Top             =   3000
      Width           =   1440
   End
   Begin VB.Shape sha 
      Height          =   495
      Index           =   3
      Left            =   4215
      Tag             =   "0340505760004950150003000042150049501440"
      Top             =   3000
      Width           =   1440
   End
   Begin VB.Shape sha 
      Height          =   495
      Index           =   2
      Left            =   3015
      Tag             =   "0340503525004950223503000030150049501215"
      Top             =   3000
      Width           =   1215
   End
   Begin VB.Shape sha 
      Height          =   495
      Index           =   1
      Left            =   1740
      Tag             =   "0340502010004950151503000017400049501290"
      Top             =   3000
      Width           =   1290
   End
   Begin VB.Shape sha 
      Height          =   495
      Index           =   0
      Left            =   480
      Tag             =   "0340500525004950150003000004800049501275"
      Top             =   3000
      Width           =   1275
   End
   Begin VB.Label lbl 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Años de Servicio"
      ForeColor       =   &H80000008&
      Height          =   285
      Index           =   13
      Left            =   5640
      TabIndex        =   31
      Tag             =   "0313507245002850150002730056400028501440"
      Top             =   2730
      Width           =   1440
   End
   Begin VB.Label lbl 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Inicia"
      ForeColor       =   &H80000008&
      Height          =   285
      Index           =   16
      Left            =   8055
      TabIndex        =   30
      Tag             =   "0313510230002850150002730080550028501500"
      Top             =   2730
      Width           =   1500
   End
   Begin VB.Label lbl 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Incorpora"
      ForeColor       =   &H80000008&
      Height          =   285
      Index           =   17
      Left            =   9540
      TabIndex        =   29
      Tag             =   "0313511715002850150002730095400028501500"
      Top             =   2730
      Width           =   1500
   End
   Begin VB.Label lbl 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "BONO VACACIONAL"
      ForeColor       =   &H80000008&
      Height          =   285
      Index           =   11
      Left            =   4215
      TabIndex        =   28
      Tag             =   "0286505760002850447002460042150028503855"
      Top             =   2460
      Width           =   3855
   End
   Begin VB.Label lbl 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Fecha Ingreso"
      ForeColor       =   &H80000008&
      Height          =   285
      Index           =   12
      Left            =   4215
      TabIndex        =   27
      Tag             =   "0313505760002850150002730042150028501440"
      Top             =   2730
      Width           =   1440
   End
   Begin VB.Label lbl 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "CALCULO DE TIEMPO"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   6
      Left            =   480
      TabIndex        =   25
      Tag             =   "0262500525002551411502220004800025511070"
      Top             =   2220
      Width           =   11070
   End
   Begin VB.Label lbl 
      Caption         =   "Sueldo Diario:"
      Height          =   375
      Index           =   19
      Left            =   9180
      TabIndex        =   11
      Tag             =   "0156010575003750087001470091800037500870"
      Top             =   1470
      Width           =   870
   End
   Begin VB.Label lbl 
      Caption         =   "Sueldo Mensual:"
      Height          =   405
      Index           =   18
      Left            =   5880
      TabIndex        =   10
      Tag             =   "0153007875004050106501440058800040501065"
      Top             =   1440
      Width           =   1065
   End
   Begin VB.Label lbl 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "FECHA DE DISFRUTE DE LAS VACACIONES"
      ForeColor       =   &H80000008&
      Height          =   285
      Index           =   15
      Left            =   8055
      TabIndex        =   9
      Tag             =   "0286510230002850441002460080550028503495"
      Top             =   2460
      Width           =   3495
   End
   Begin VB.Label lbl 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Días Bono"
      ForeColor       =   &H80000008&
      Height          =   285
      Index           =   14
      Left            =   7065
      TabIndex        =   8
      Tag             =   "0313508730002850150002730070650028501005"
      Top             =   2730
      Width           =   1005
   End
   Begin VB.Label lbl 
      Alignment       =   2  'Center
      Caption         =   "Dias a Pagar"
      Height          =   285
      Index           =   10
      Left            =   3210
      TabIndex        =   7
      Tag             =   "0300004095002850094002595032100028500940"
      Top             =   2595
      Width           =   940
   End
   Begin VB.Label lbl 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Hasta"
      ForeColor       =   &H80000008&
      Height          =   285
      Index           =   9
      Left            =   1740
      TabIndex        =   6
      Tag             =   "0313502010002850151502730017400028501290"
      Top             =   2730
      Width           =   1290
   End
   Begin VB.Label lbl 
      Caption         =   "Cargo:"
      Height          =   285
      Index           =   5
      Left            =   480
      TabIndex        =   5
      Tag             =   "0163502910002850121501545004800028501215"
      Top             =   1545
      Width           =   1215
   End
   Begin VB.Label lbl 
      Caption         =   "Cédula de Identidad:"
      Height          =   390
      Index           =   4
      Left            =   9180
      TabIndex        =   4
      Tag             =   "0153000525003900121500345091800039001215"
      Top             =   345
      Width           =   1215
   End
   Begin VB.Label lbl 
      Caption         =   "Empleado:"
      Height          =   285
      Index           =   3
      Left            =   5880
      TabIndex        =   3
      Tag             =   "0070510575002850093000945058800028500930"
      Top             =   945
      Width           =   930
   End
   Begin VB.Label lbl 
      Caption         =   "Código Empleado:"
      Height          =   420
      Index           =   2
      Left            =   5880
      TabIndex        =   2
      Tag             =   "0063007875004200097500315058800042000975"
      Top             =   315
      Width           =   975
   End
   Begin VB.Label lbl 
      Caption         =   "Inmueble:"
      Height          =   285
      Index           =   1
      Left            =   480
      TabIndex        =   1
      Tag             =   "0070502910002850093000945004800028500930"
      Top             =   945
      Width           =   930
   End
   Begin VB.Label lbl 
      Caption         =   "Código Inmueble:"
      Height          =   420
      Index           =   0
      Left            =   525
      TabIndex        =   0
      Tag             =   "0063000525004200091500315005250042000915"
      Top             =   315
      Width           =   915
   End
   Begin VB.Label lbl 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   555
      Index           =   20
      Left            =   3015
      TabIndex        =   34
      Tag             =   "0286503525005550223502460030150055501215"
      Top             =   2460
      Width           =   1215
   End
End
Attribute VB_Name = "frmVacacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmd_Click(Index%)
'variables locales
Select Case Index
    Case 0  'calcular vacaciones
        Call Calcula_Vacacion
    Case 1  'cerrar ventana
        Unload Me
        Set frmVacacion = Nothing
    Case 2  'procesar las vacaciones
        Call Procesar_Vacaciones
    Case 3 ' suplencias
        frmVacaSup.Show vbModal, FrmAdmin
End Select
'
End Sub

Private Sub dtc_Click(Index As Integer, Area As Integer)
'variables locales
If Area = 2 Then
    Select Case Index
        Case 0, 1   '//
            dtc(IIf(Index = 0, 1, 0)) = dtc(Index).BoundText
            Call Borrar_Entradas
            Call Listar_Empleados
            
            
        Case 2, 3   '//
            dtc(IIf(Index = 2, 3, 2)) = dtc(Index).BoundText
            Call Busca_emp
            
    End Select
End If
'
End Sub

Private Sub dtc_KeyPress(Index As Integer, KeyAscii As Integer)
'variables locales
KeyAscii = Asc(UCase(Chr(KeyAscii)))
If KeyAscii = 13 Then
    Select Case Index
        Case 0  '
            Call dtc_Click(0, 2)
            If dtc(1).MatchedWithList Then dtc(2).SetFocus
        Case 1  '
            Call Busca_Inmueble
        Case 2  '
            Call dtc_Click(2, 2)
        Case 3  '
            Call Busca_emp
    End Select
End If
'
End Sub

Private Sub Form_Load()
'variables locales
Call Configurar_ADO(Ado(0), adCmdTable, "Inmueble", gcPath & "\sac.mdb", "CodInm")
Call Configurar_ADO(Ado(1), adCmdText, "SELECT * FROM Inmueble ORDER BY Nombre", gcPath _
& "\sac.mdb")
'
End Sub

Private Sub Listar_Empleados()
'variables locales
Dim strArchivo As String

strArchivo = gcPath & "\" & dtc(0) & "\inm.mdb"

If Dir(strArchivo) <> "" Then
    Call Configurar_ADO(Ado(2), adCmdText, "SELECT *,Apellidos & ', ' & Nombres as Empleado " _
    & "FROM Emp WHERE CodInm='" & dtc(0) & "' AND CodEstado = 0 ORDER BY CodEmp", gcPath & "\sac.mdb")
    Call Configurar_ADO(Ado(3), adCmdText, "SELECT Emp.*,Emp.Apellidos & ', ' & Emp.Nombres as " _
    & "Empleado, Emp_Cargos.NombreCargo FROM Emp INNER JOIN Emp_Cargos ON Emp.CodCargo = Emp_Car" _
    & "gos.CodCargo WHERE CodInm='" & dtc(0) & "' AND CodEstado = 0 ORDER BY Apellidos", gcPath _
    & "\sac.mdb")
    Set dtc(2).RowSource = Ado(2)
    Set dtc(3).RowSource = Ado(3)
    dtc(2).SetFocus
Else
    Set dtc(2).RowSource = Nothing
    Set dtc(3).RowSource = Nothing
    
End If
cmd(0).Enabled = False
cmd(2).Enabled = False
'
End Sub

Private Sub Busca_Inmueble()
'variables locales
dtc(2) = ""
dtc(3) = ""
'
With Ado(1).Recordset

    If Not .EOF And Not .BOF Then
        .MoveFirst
        .Find "Nombre LIKE '%" & dtc(1) & "%'"
        If Not .EOF Then
            dtc(0) = !CodInm
            dtc(1) = !Nombre
            dtc(2).SetFocus
        Else
            .MoveFirst
            dtc(0) = ""
        End If
    End If
End With
Call Listar_Empleados
'
End Sub

Private Sub Busca_emp()
'variables locales
For I = 0 To 11: Txt(I) = ""
Next
If dtc(3) = "" Then Exit Sub
With Ado(3).Recordset
    If Not (.BOF And .EOF) Then
        
        .MoveFirst
            
            .Find "Empleado LIKE '%" & dtc(3) & "%'"
            If Not .EOF Then
                dtc(2) = !CodEmp
                dtc(3) = !Empleado
                Txt(0) = Format(!Cedula, "#,##0 ")
                Txt(1) = !NombreCargo
                Txt(2) = Format(!Sueldo + (!Sueldo * !BonoNoc / 100), "#,##0.00 ")
                Txt(3) = Format((!Sueldo + (!Sueldo * !BonoNoc / 100)) / 30, "#,##0.00 ")
                Txt(6) = !Fingreso
                Txt(7) = Fix(DateDiff("m", !Fingreso, Date) / 12)
                cmd(0).Enabled = True
                cmd(2).Enabled = True
            Else
                cmd(0).Enabled = False
                cmd(2).Enabled = False
            End If
    End If
    
End With
'
End Sub

Private Sub Borrar_Entradas()
'variables locales
dtc(2) = ""
dtc(3) = ""
Txt(0) = ""
Txt(1) = ""
Txt(2) = ""
Txt(3) = ""
Txt(4) = ""
Txt(5) = ""
Txt(6) = ""
Txt(7) = ""
Txt(8) = ""
Txt(11) = ""
End Sub

Private Sub Calcula_Vacacion(Optional DiasFeriados As Long)
'vairables locales
Dim Dia_Vaca As Long, Dia_Bono As Long, lngAno As Long, lngAdi As Long
Dim pSSO@, pSPF@, pLPH@
Dim rstlocal As New ADODB.Recordset
Dim strSQL As String
Dim UltCobro As Date, Fecha1 As Date
Dim IDNom As Long, Lunes As Long
Dim dbSueldo1 As Double
'
If Txt(7) = 0 Then
    MsgBox "Este empleado no tiene el año cumplido", vbInformation, App.ProductName
    Exit Sub
End If
If Not IsDate(Txt(9)) Then
    MsgBox "Revise el valor del campo Inicio de vacaciones", vbCritical, App.ProductName
    Exit Sub
'Else
'    If CDate(txt(9)) < Date Then
'        MsgBox "La fecha de inicio de vacaciones no puede ser menor a la fecha actual", vbCritical, App.ProductName
'        Exit Sub
'    End If
End If
'If Not IsDate(txt(10)) Then
'    MsgBox "Revise el valor del campo fin del período de vacaciones", vbCritical, App.ProductName
'    Exit Sub
'Else
'If CDate(txt(10)) < CDate(txt(9)) Then
'    MsgBox "La fecha de incorporación no puede ser anterior a la fecha de inicio", vbCritical, App.ProductName
'    Exit Sub
'End If
    
'End If

rstlocal.Open "Nom_Calc", cnnConexion, adOpenKeyset, adLockOptimistic, adCmdTable
If Not rstlocal.EOF And Not rstlocal.BOF Then

    If IsNull(rstlocal!Dia_Vacacion) Then
        MsgBox "Revise en los parámtros de la nómina el valor de: 'Dias Vacación'", _
        vbInformation, App.ProductName
        Exit Sub
    End If
    '
    If IsNull(rstlocal!Dia_Bono) Then
        MsgBox "Revise en los parámtros de la nómina el valor de:'Dias Bono'", _
        vbInformation, App.ProductName
        Exit Sub
    End If
    '
    If IsNull(rstlocal!Adi_Ano) Or IsNull(rstlocal!Dia_Adi) Then
        MsgBox "Revise los parámtros de la nómina", vbInformation, App.ProductName
        Exit Sub
    End If
    
    Dia_Vaca = rstlocal!Dia_Vacacion
    Dia_Bono = rstlocal!Dia_Bono
    lngAno = rstlocal!Adi_Ano
    lngAdi = rstlocal!Dia_Adi
    pSSO = rstlocal!SSO / 100
    pSPF = rstlocal!SPF / 100
    pLPH = rstlocal!LPH / 100
End If
rstlocal.Close
'strSql = "SELECT Nom_Detalle.CodEmp, Nom_Inf.Fecha, Nom_Inf.IDNomina, Nom_Inf.EFECTIVO " _
'& "FROM Nom_Inf INNER JOIN Nom_Detalle ON Nom_Inf.IDNomina = Nom_Detalle.IDNom " _
'& "Where (((Nom_Detalle.CodEmp) =" & dtc(2) & ") And ((Left([Nom_Inf].[IDNomina], 1)) <> 3)) " _
'& "ORDER BY Nom_Inf.EFECTIVO DESC;"
'rstlocal.Open strSql, cnnConexion, adOpenKeyset, adLockOptimistic, adCmdText
'
'If Not (rstlocal.EOF And rstlocal.BOF) Then IDNom = rstlocal!IDNomina
'If IDNom > 0 Then
'
'    If Left(IDNom, 1) = 1 Then
'        UltCobro = "15/" & Mid(IDNom, 2, 2) & "/" & Right(IDNom, 4)
'    Else
'
'        UltCobro = "01/" & Mid(IDNom, 2, 2) & "/" & Right(IDNom, 4)
'        UltCobro = DateAdd("M", 1, UltCobro)
'        UltCobro = DateAdd("D", -1, UltCobro)
'
'    End If
'End If
'rstlocal.Close
If Day(Txt(9)) <= 15 Then
    UltCobro = DateAdd("d", -Day(Txt(9)), Txt(9))
Else
    UltCobro = DateAdd("d", 15 - Day(Txt(9)), Txt(9))
End If
'
'perdiodo correspondientes
Txt(4) = DateAdd("yyyy", Txt(7) + IIf(CDate(Format(Txt(6), "dd/mm") & "/" _
& Year(Date)) >= Date, -1, -1), Txt(6)) 'desde
If CDate(Txt(4)) < CDate(Txt(6)) Then Txt(4) = Txt(6)
Txt(5) = DateAdd("yyyy", 1, Txt(4)) 'hasta
Txt(11) = Dia_Vaca + Txt(7) - 1 'dias a pagar
Txt(8) = Dia_Bono + Txt(7) - 1  'dias de bono
'
'If CDate(txt(10)) < DateAdd("d", txt(11), txt(9)) Then
'    MsgBox "El período de vacaciones es menor al los días calculados", vbCritical, App.ProductName
'End If
'calcula la fecha de reintegro del empleado
Txt(10) = DateAdd("d", Txt(11), Txt(9))

For Fecha1 = Txt(9) To Txt(10)
    If Weekday(Fecha1, vbSunday) = vbSunday Then
       Txt(10) = DateAdd("d", 1, Txt(10))
       If Weekday(Txt(10), vbSunday) = vbSunday Then Txt(10) = DateAdd("d", 1, Txt(10))
    End If
    'If Weekday(txt(10), vbSunday) = vbSunday Then txt(10) = DateAdd("d", 1, txt(10))
Next

If Weekday(Txt(10), vbSunday) = vbSunday Then Txt(10) = DateAdd("d", 1, Txt(10))
strSQL = "SELECT * FROM Nom_Feriados WHERE Fecha>=#" & Format(Txt(9), "mm/dd/yy") & "# AND Fecha<=#" & Format(Txt(10), "mm/dd/yy") & "#"
'rstLocal.Open "nom_feriados", cnnConexion, adOpenStatic, adLockReadOnly, adCmdTable
rstlocal.Open strSQL, cnnConexion, adOpenKeyset, adLockOptimistic, adCmdText
'rstLocal.Filter = "Fecha>=#" & Format(txt(9), "mm/dd/yy") & "# AND Fecha<=#" & Format(txt(10), "mm/dd/yy") & "#"

If Not (rstlocal.EOF And rstlocal.BOF) Then
    rstlocal.MoveFirst
    Do
        If Weekday(rstlocal("Fecha"), vbSunday) <> vbSunday Then Txt(10) = DateAdd("d", 1, Txt(10))
        rstlocal.MoveNext
    Loop Until rstlocal.EOF
    'txt(10) = DateAdd("d", rstlocal.RecordCount, txt(10))
End If
If Weekday(Txt(10), vbSunday) = vbSunday Then Txt(10) = DateAdd("d", 1, Txt(10))
'txt(10) = DateAdd("d", 1, txt(10))
rstlocal.Close
'SUELDOS DIARIOS
'sueldo actual
'lbl(38) = txt(3)
'sueldo al momento de vencerse las vacaciones
Set rstlocal = Nothing
Set rstlocal = New ADODB.Recordset
strSQL = "SELECT * FROM Nom_Detalle WHERE IDNom=" & "02" & _
Format(DateAdd("M", -1, Txt(5)), "MMYYYY") & " AND CodEmp=" & dtc(2)
With rstlocal
    .Open strSQL, cnnConexion, adOpenKeyset, adLockOptimistic, adCmdText
    'Format((!Sueldo + (!Sueldo * !BonoNoc / 100)) / 30, "#,##0.00 ")
    If Not .EOF And Not .BOF Then
        dbSueldo1 = (!Sueldo + (!Sueldo * !Porc_BonoNoc / 100)) / 30
    End If
    Lbl(35) = Format(dbSueldo1, "#,##0.00 ")
    Rem lbl(35) = txt(3)
    Lbl(36) = Lbl(35)
    Lbl(37) = Lbl(35)
    .Close
End With
Set rstlocal = Nothing
'DIAS CALCULO
Lbl(29) = Txt(11)   'días hábiles
'dias feriados
Lbl(30) = IIf(DiasFeriados = 0, DateDiff("D", Txt(9), Txt(10)) - Lbl(29), DiasFeriados)
Lbl(31) = Txt(8)
'lbl(32) = DateDiff("D", UltCobro, DateAdd("d", -1, txt(9)))
'lbl(26) = "Días Trabajados de " & Format(DateAdd("d", 1, UltCobro), "dd/mm/yy") & " al " & _
Format(txt(9), "dd/mm/yy")
'TOTALES ASIGNACIONES
Lbl(39) = Format(CCur(Lbl(29)) * CCur(Lbl(35)), "#,##0.00")
Lbl(40) = Format(CCur(Lbl(30)) * CCur(Lbl(36)), "#,##0.00")
Lbl(41) = Format(CCur(Lbl(31)) * CCur(Lbl(37)), "#,##0.00")
'lbl(42) = Format(CCur(lbl(32)) * CCur(lbl(38)), "#,##0.00")
Lbl(44) = Format(CCur(Lbl(39)) + CCur(Lbl(40)) + CCur(Lbl(41)), "#,##0.00")
'DEDUCCIONES
Dim j As Date, K As Date, q As Date
Lbl(49) = "0,00"
Lbl(50) = "0,00"
Lbl(51) = "0,00"
Lbl(53) = "0,00"
If Day(UltCobro) = 15 Then
    If Month(UltCobro) <> Month(Txt(10)) Then
        K = "01/" & Format(UltCobro, "mm/yy")
        q = "01/" & Format(DateAdd("m", 1, UltCobro), "mm/yy")
        q = DateAdd("d", -1, q)
    
        For j = K To q: If (Weekday(j, vbSunday)) = 2 Then Lunes = Lunes + 1
        Next

        Lbl(49) = Format((CCur(Txt(2)) * 12 / 52) * pSSO * Lunes, "#,#00.00")
        Lbl(50) = Format((CCur(Txt(2)) * 12 / 52) * pSPF * Lunes, "#,#00.00")
        Lbl(51) = Format(CCur(Txt(2)) * pLPH, "#,#00.00")
        Lbl(53) = Format(CCur(Lbl(49)) + CCur(Lbl(50)) + CCur(Lbl(51)), "#,##0.00")
    End If
    
Else

    If Month(Txt(9)) <> Month(Txt(10)) Then
        K = "01/" & Format(Txt(9), "mm/yy")
        q = "01/" & Format(DateAdd("m", 1, Txt(9)), "mm/yy")
        q = DateAdd("d", -1, q)
        
        For j = K To q: If (Weekday(j, vbSunday)) = 2 Then Lunes = Lunes + 1
        Next

        Lbl(49) = Format((CCur(Txt(2)) * 12 / 52) * pSSO * Lunes, "#,#00.00")
        Lbl(50) = Format((CCur(Txt(2)) * 12 / 52) * pSPF * Lunes, "#,#00.00")
        Lbl(51) = Format(CCur(Txt(2)) * pLPH, "#,#00.00")
        Lbl(53) = Format(CCur(Lbl(49)) + CCur(Lbl(50)) + CCur(Lbl(51)), "#,##0.00")
        
    End If
    
End If
'neto a pagar
Lbl(54) = Format(CCur(Lbl(44)) - CCur(Lbl(53)), "#,##0.00")

End Sub

Private Sub Form_Resize()
'variables Locales
Dim strTag As String
'
For Each ctl In Controls
'
    strTag = IIf(ScaleWidth > 15000, Left(ctl.Tag, 20), Right(ctl.Tag, 20))
    '
    If Len(strTag) = 20 Then
        '
        If TypeName(ctl) = "Line" Then
            '
            ctl.X1 = Left(strTag, 5)
            ctl.X2 = Mid(strTag, 6, 5)
            ctl.Y1 = Mid(strTag, 11, 5)
            ctl.Y2 = Right(strTag, 5)
            '
        Else
            '
            ctl.Top = Left(strTag, 5)
            ctl.Left = Mid(strTag, 6, 5)
            ctl.Height = Mid(strTag, 11, 5)
            ctl.Width = Right(strTag, 5)
            '
        End If
    '
    End If
    
Next
'
End Sub

Private Sub txt_KeyPress(Index As Integer, KeyAscii As Integer)
'VARIABLES LOCALES
KeyAscii = Asc(UCase(Chr(KeyAscii)))

Select Case Index
    
    Case 9, 10
        If KeyAscii = Asc("-") Then KeyAscii = Asc("/")
        Call Validacion(KeyAscii, "0123456789/")
    
    Case 12
        Call Validacion(KeyAscii, "0123456789")
        If KeyAscii = 13 Then
            Call Calcula_Vacacion(Txt(12))
            
        End If
End Select
If KeyAscii = 13 Then SendKeys vbTab
'
End Sub

Private Sub Procesar_Vacaciones()
'variables locales
Dim rstlocal As ADODB.Recordset
Dim Mensaje As String
Dim strDet$, NFactura$
Dim rpReporte As ctlReport
'
Mensaje = "Se procesará las vacaciones de:" & vbCrLf & vbCrLf & dtc(2) & " - " & dtc(3) & vbCrLf _
& vbCrLf & "¿Seguro de llevar a cabo esta operación?"
'
If Respuesta(Mensaje) Then
    strDet = "CANC. VACACIONES PERIODO: " & Year(Txt(4)) & "-" & Year(Txt(5)) & " " & dtc(1) & _
    " (" & dtc(0) & ")."
    NFactura = Registro_Proveedor(dtc(3), True, "NO S/N", strDet, Lbl(54), dtc(0))
'    Set rstlocal = New ADODB.Recordset
'
'    rstlocal.Open "Proveedores", cnnConexion, adOpenKeyset, adLockOptimistic, adCmdTable
'
'    If Not (rstlocal.EOF And rstlocal.BOF) Then
'        '
'        rstlocal.MoveFirst
'        rstlocal.Find "NombProv ='" & Dtc(1) & "'"
'        '
'        If rstlocal.EOF Then
'            '
'            rstlocal.Close
'            rstlocal.Open "SELECT Max(Codigo) FROM Proveedores ", cnnConexion, adOpenKeyset, _
'            adLockOptimistic, adCmdText
'            If IsNull(rstlocal.Fields(0)) Then
'                Nproveedor = 1
'            Else
'                Nproveedor = rstlocal.Fields(0) + 1
'            End If
'            '
'            cnnConexion.Execute "INSERT INTO Proveedores (Codigo,Rif,NombProv,Ramo,Actividad," _
'            & "FecReg,Nacional,Beneficiario,Usuario,Freg) VALUES ('" & Nproveedor & "','" & Txt(0) _
'            & "','" & Dtc(3) & "','NOMINA','" & Txt(1) & "',Date(),-1,'" & Dtc(3) & "','" & gcUsuario & "',Date())"
'            '
'        End If
'
'    End If
'    '
'    rstlocal.Close
'    'agrega la factura por pagar
'    NFactura = FrmFactura.FntStrDoc
'    cnnConexion.Execute "INSERT INTO CPP(Tipo,NDoc,Fact,CodProv,Benef,Detalle," _
'    & "Monto,Ivm,Total,FRecep,Fecr,Fven,CodInm,Moneda,Estatus,Usuario,Freg) VALUES ('NO'" _
'    & ",'" & NFactura & "','" & NFactura & "','" & Nproveedor & "','" & Dtc(3) & "','CANC. VACACIONES" _
'    & " PERIODO: " & Year(Txt(4)) & "-" & Year(Txt(5)) & " " & Dtc(1) & " (" & Dtc(0) & ")." & "','" _
'    & lbl(54) & "',0,'" & lbl(54) & "',Date(),Date(),Date(),'" & Dtc(0) & "','Bs','PENDIENTE','" & gcUsuario _
'    & "',Date())"
    'actualiza la fecha de reintegro del employes
    cnnConexion.Execute "UPDATE Emp SET FechaReintegro ='" & Txt(10) & _
    "',Usuario = '" & gcUsuario & "', FecAct=Date() WHERE CodEmp=" & dtc(2)
'    'actualiza el estado del empleado si sale el primer dia de la nómina
'    If Day(txt(9)) = 1 Or Day(txt(9)) = 16 Then
'        cnnConexion.Execute "UPDATE Emp SET CodEstado=3 WHERE CodEmp=" & dtc(2)
'    End If
    'agrega el registro en la tabla emp_vaca
    cnnConexion.Execute "INSERT INTO Nom_Vaca (CodEmp,Desde,Hasta,Dias_Vaca,Dias_Bono," _
    & "Dias_Feri,Dias_Traba,Inicia,Incorpora,SueldoMensual,SSO,SPF,LPH,Usuario,Freg,Fact," _
    & "IDNomina,NFact,Sueldo1) VALUES ('" & dtc(2) & "','" & Txt(4) & "','" & Txt(5) & "'," & Txt(11) & "," _
    & Txt(8) & "," & Lbl(30) & ",0,'" & Txt(9) & "','" & Txt(10) & "','" & Txt(2) & "','" & _
    Lbl(49) & "','" & Lbl(50) & "','" & Lbl(51) & "','" & gcUsuario & "',Date(),Date()," & _
    IIf(Day(Txt(10)) <= 15, 1, 2) & Format(CDate(Txt(10)), "mmyyyy") & ",'" & NFactura & "','" & Lbl(35) & "')"
    'imprime el comprobante de las vacaciones
    cmd(0).Enabled = False
    cmd(2).Enabled = False
    
    'Call clear_Crystal(FrmAdmin.rptReporte)
    Set rpReporte = New ctlReport
    With rpReporte
        '
        '.ProgressDialog = False
        .Reporte = gcReport & "nom_vaca.rpt"
        .OrigenDatos(0) = gcPath & "\sac.mdb"
        Desde = Format(Txt(4), "yyyy,mm,dd") & ")"
        .FormuladeSeleccion = "{Nom_Vaca.CodEmp}=" & dtc(2) & " and {Nom_Vaca.Desde}=cDate (" & Desde
        '.Formulas(0) = "Dias_tra='" & lbl(26) & "'"
        .Salida = crImpresora
        '.CopiesToPrinter = 3
        'errlocal = .PrintReport
        .Imprimir 3
        
    End With
    Set rpReporte = Nothing
    '
End If
'
End Sub
