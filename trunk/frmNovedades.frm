VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmNovedades 
   Caption         =   "Novedades"
   ClientHeight    =   570
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   2895
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   570
   ScaleWidth      =   2895
   Tag             =   "1"
   WindowState     =   2  'Maximized
   Begin ComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   480
      Left            =   0
      TabIndex        =   30
      Top             =   0
      Width           =   2895
      _ExtentX        =   5106
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
            Enabled         =   0   'False
            Object.Visible         =   0   'False
            Key             =   "First"
            Object.ToolTipText     =   "Primer Registro"
            Object.Tag             =   ""
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Enabled         =   0   'False
            Object.Visible         =   0   'False
            Key             =   "Previous"
            Object.ToolTipText     =   "Registro Anterior"
            Object.Tag             =   ""
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Enabled         =   0   'False
            Object.Visible         =   0   'False
            Key             =   "Next"
            Object.ToolTipText     =   "Siguiente Registro"
            Object.Tag             =   ""
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Enabled         =   0   'False
            Object.Visible         =   0   'False
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
            Object.Visible         =   0   'False
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
      MouseIcon       =   "frmNovedades.frx":0000
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   6810
      Left            =   210
      TabIndex        =   13
      Top             =   690
      Width           =   11310
      _ExtentX        =   19950
      _ExtentY        =   12012
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      ShowFocusRect   =   0   'False
      TabCaption(0)   =   "Datos Generales"
      TabPicture(0)   =   "frmNovedades.frx":031A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "fraAgasto(1)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "fraAgasto(2)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "fraAgasto(4)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "Lista"
      TabPicture(1)   =   "frmNovedades.frx":0336
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "DataGrid1"
      Tab(1).Control(1)=   "fraNov(0)"
      Tab(1).Control(2)=   "fraNov(1)"
      Tab(1).Control(3)=   "fraNov(2)"
      Tab(1).ControlCount=   4
      Begin VB.Frame fraAgasto 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   3405
         Index           =   4
         Left            =   6210
         TabIndex        =   51
         Top             =   1365
         Visible         =   0   'False
         Width           =   3300
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid GridGastos 
            Height          =   2760
            Index           =   2
            Left            =   60
            TabIndex        =   57
            Top             =   525
            Width           =   3165
            _ExtentX        =   5583
            _ExtentY        =   4868
            _Version        =   393216
            Cols            =   3
            BackColorFixed  =   -2147483646
            ForeColorFixed  =   -2147483639
            BackColorBkg    =   -2147483636
            WordWrap        =   -1  'True
            _NumberOfBands  =   1
            _Band(0).Cols   =   3
         End
         Begin VB.CommandButton cmdHelp 
            BackColor       =   &H00FFFFFF&
            Height          =   300
            Index           =   3
            Left            =   1530
            Picture         =   "frmNovedades.frx":0352
            Style           =   1  'Graphical
            TabIndex        =   59
            Top             =   2250
            Width           =   300
         End
         Begin VB.CommandButton cmdHelp 
            BackColor       =   &H00FFFFFF&
            Height          =   300
            Index           =   1
            Left            =   1530
            Picture         =   "frmNovedades.frx":04A0
            Style           =   1  'Graphical
            TabIndex        =   58
            Top             =   1785
            Width           =   300
         End
         Begin VB.ListBox LisAGasto 
            Height          =   2010
            Index           =   1
            Left            =   1860
            Sorted          =   -1  'True
            TabIndex        =   56
            Top             =   1215
            Width           =   1350
         End
         Begin VB.ListBox LisAGasto 
            BackColor       =   &H00FFFFFF&
            Height          =   2010
            Index           =   0
            ItemData        =   "frmNovedades.frx":05EE
            Left            =   150
            List            =   "frmNovedades.frx":05F0
            MultiSelect     =   2  'Extended
            Sorted          =   -1  'True
            TabIndex        =   55
            Top             =   1215
            Width           =   1350
         End
         Begin VB.CommandButton cmdHelp 
            BackColor       =   &H00FFFFFF&
            Caption         =   "&Aceptar"
            Height          =   315
            Index           =   0
            Left            =   2040
            Style           =   1  'Graphical
            TabIndex        =   54
            Top             =   120
            Width           =   1170
         End
         Begin VB.OptionButton Option1 
            BackColor       =   &H00FFFFFF&
            Caption         =   "&Propietarios Determinados"
            Height          =   270
            Index           =   1
            Left            =   100
            TabIndex        =   53
            Top             =   540
            Width           =   2190
         End
         Begin VB.OptionButton Option1 
            BackColor       =   &H00FFFFFF&
            Caption         =   "&Todos los Propietarios"
            Height          =   270
            Index           =   0
            Left            =   100
            TabIndex        =   52
            Top             =   165
            Value           =   -1  'True
            Width           =   2190
         End
      End
      Begin VB.Frame fraNov 
         Caption         =   "Registrar Factura"
         Height          =   1395
         Index           =   2
         Left            =   -67050
         TabIndex        =   47
         Top             =   5205
         Width           =   3090
         Begin VB.CommandButton cmdNov 
            Caption         =   "Aceptar"
            Height          =   390
            Left            =   150
            TabIndex        =   49
            Top             =   840
            Width           =   2805
         End
         Begin VB.TextBox TxtAGasto 
            Height          =   280
            Index           =   14
            Left            =   1395
            MaxLength       =   7
            TabIndex        =   48
            Top             =   345
            Width           =   1500
         End
         Begin VB.Label label1 
            Caption         =   "Nº Documento:"
            Height          =   285
            Index           =   14
            Left            =   210
            TabIndex        =   50
            Top             =   390
            Width           =   1110
         End
      End
      Begin VB.Frame fraNov 
         Caption         =   "Aaplicar filtro:"
         Height          =   1380
         Index           =   1
         Left            =   -71265
         TabIndex        =   36
         Top             =   5205
         Width           =   4125
         Begin VB.TextBox TxtAGasto 
            Height          =   280
            Index           =   13
            Left            =   2985
            MaxLength       =   15
            TabIndex        =   46
            Top             =   570
            Width           =   975
         End
         Begin VB.TextBox TxtAGasto 
            Height          =   280
            Index           =   12
            Left            =   2985
            MaxLength       =   10
            TabIndex        =   45
            Top             =   225
            Width           =   975
         End
         Begin VB.TextBox TxtAGasto 
            Height          =   280
            Index           =   11
            Left            =   1020
            MaxLength       =   10
            TabIndex        =   44
            Top             =   915
            Width           =   975
         End
         Begin VB.TextBox TxtAGasto 
            Height          =   280
            Index           =   10
            Left            =   1020
            MaxLength       =   7
            TabIndex        =   43
            Top             =   570
            Width           =   975
         End
         Begin VB.TextBox TxtAGasto 
            Height          =   280
            Index           =   9
            Left            =   1020
            MaxLength       =   6
            TabIndex        =   42
            Top             =   225
            Width           =   975
         End
         Begin VB.Label label1 
            Caption         =   "Usuario:"
            Height          =   285
            Index           =   13
            Left            =   2250
            TabIndex        =   41
            Top             =   630
            Width           =   1005
         End
         Begin VB.Label label1 
            Caption         =   "Fecha:"
            Height          =   285
            Index           =   12
            Left            =   2250
            TabIndex        =   40
            Top             =   285
            Width           =   1005
         End
         Begin VB.Label label1 
            Caption         =   "Monto:"
            Height          =   285
            Index           =   11
            Left            =   195
            TabIndex        =   39
            Top             =   975
            Width           =   1215
         End
         Begin VB.Label label1 
            Caption         =   "Cargado:"
            Height          =   285
            Index           =   10
            Left            =   195
            TabIndex        =   38
            Top             =   630
            Width           =   1215
         End
         Begin VB.Label label1 
            Caption         =   "Cod.Gasto:"
            Height          =   285
            Index           =   7
            Left            =   195
            TabIndex        =   37
            Top             =   285
            Width           =   900
         End
      End
      Begin VB.Frame fraNov 
         Caption         =   "Ordenar por:"
         Height          =   1380
         Index           =   0
         Left            =   -74745
         TabIndex        =   32
         Top             =   5205
         Width           =   3375
         Begin VB.OptionButton OptNov 
            Caption         =   "Monto"
            Height          =   285
            Index           =   2
            Left            =   735
            TabIndex        =   35
            Tag             =   "Monto,Fecha,CodGasto"
            Top             =   990
            Width           =   1425
         End
         Begin VB.OptionButton OptNov 
            Caption         =   "Código del Gasto"
            Height          =   285
            Index           =   1
            Left            =   420
            TabIndex        =   34
            Tag             =   "CodGasto,Fecha,Monto"
            Top             =   645
            Width           =   1845
         End
         Begin VB.OptionButton OptNov 
            Caption         =   "Fecha"
            Height          =   285
            Index           =   0
            Left            =   210
            TabIndex        =   33
            Tag             =   "Fecha,CodGasto,Monto"
            Top             =   300
            Value           =   -1  'True
            Width           =   1425
         End
      End
      Begin VB.Frame fraAgasto 
         Caption         =   "Distribución del Gasto:"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   2265
         Index           =   2
         Left            =   105
         TabIndex        =   16
         Top             =   465
         Width           =   11025
         Begin VB.CheckBox Check1 
            BackColor       =   &H80000002&
            Caption         =   "Couta X / N"
            ForeColor       =   &H80000009&
            Height          =   390
            Index           =   6
            Left            =   8715
            TabIndex        =   29
            Top             =   1365
            Width           =   1590
         End
         Begin VB.TextBox TxtAGasto 
            Alignment       =   2  'Center
            Enabled         =   0   'False
            Height          =   280
            Index           =   8
            Left            =   6915
            MaxLength       =   6
            TabIndex        =   17
            Top             =   1380
            Width           =   495
         End
         Begin VB.CheckBox Check1 
            BackColor       =   &H80000002&
            Caption         =   "Dividir en               períodos"
            ForeColor       =   &H80000009&
            Height          =   390
            Index           =   5
            Left            =   5880
            TabIndex        =   28
            Top             =   1365
            Width           =   2295
         End
         Begin VB.CheckBox Check1 
            BackColor       =   &H80000002&
            Caption         =   "Reintegrar factura"
            ForeColor       =   &H80000009&
            Height          =   390
            Index           =   4
            Left            =   465
            TabIndex        =   27
            Top             =   1365
            Width           =   1230
         End
         Begin VB.CheckBox Check1 
            BackColor       =   &H80000005&
            Height          =   195
            Index           =   3
            Left            =   7200
            TabIndex        =   7
            Top             =   640
            Width           =   195
         End
         Begin VB.CheckBox Check1 
            BackColor       =   &H80000005&
            Height          =   195
            Index           =   2
            Left            =   6420
            TabIndex        =   5
            Top             =   640
            Width           =   180
         End
         Begin VB.TextBox TxtAGasto 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            Height          =   280
            Index           =   7
            Left            =   3885
            Locked          =   -1  'True
            TabIndex        =   23
            Top             =   1410
            Width           =   1290
         End
         Begin VB.TextBox TxtAGasto 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            Height          =   280
            Index           =   6
            Left            =   2595
            Locked          =   -1  'True
            TabIndex        =   22
            Top             =   1410
            Width           =   1290
         End
         Begin VB.TextBox TxtAGasto 
            Enabled         =   0   'False
            Height          =   280
            Index           =   5
            Left            =   1725
            MaxLength       =   7
            TabIndex        =   21
            Top             =   1410
            Width           =   870
         End
         Begin VB.TextBox TxtAGasto 
            Height          =   315
            Index           =   0
            Left            =   195
            MaxLength       =   6
            TabIndex        =   1
            Top             =   570
            Width           =   1470
         End
         Begin VB.TextBox TxtAGasto 
            Enabled         =   0   'False
            Height          =   315
            Index           =   2
            Left            =   6105
            Locked          =   -1  'True
            TabIndex        =   20
            Top             =   570
            Width           =   795
         End
         Begin VB.TextBox TxtAGasto 
            Enabled         =   0   'False
            Height          =   315
            Index           =   3
            Left            =   6900
            Locked          =   -1  'True
            TabIndex        =   19
            Top             =   570
            Width           =   750
         End
         Begin VB.TextBox TxtAGasto 
            Alignment       =   1  'Right Justify
            Height          =   315
            Index           =   4
            Left            =   7650
            TabIndex        =   9
            Top             =   570
            Width           =   1740
         End
         Begin VB.ComboBox Combo1 
            Height          =   315
            Index           =   0
            ItemData        =   "frmNovedades.frx":05F2
            Left            =   9390
            List            =   "frmNovedades.frx":061C
            Style           =   2  'Dropdown List
            TabIndex        =   11
            Top             =   570
            Width           =   870
         End
         Begin VB.ComboBox Combo1 
            Height          =   315
            Index           =   1
            ItemData        =   "frmNovedades.frx":065C
            Left            =   10245
            List            =   "frmNovedades.frx":065E
            Style           =   2  'Dropdown List
            TabIndex        =   12
            Top             =   570
            Width           =   585
         End
         Begin VB.CommandButton cmdHelp 
            Enabled         =   0   'False
            Height          =   300
            Index           =   2
            Left            =   5265
            Picture         =   "frmNovedades.frx":0660
            Style           =   1  'Graphical
            TabIndex        =   18
            Top             =   1410
            Width           =   225
         End
         Begin VB.TextBox TxtAGasto 
            Height          =   315
            Index           =   1
            Left            =   1665
            TabIndex        =   3
            Top             =   570
            Width           =   4440
         End
         Begin VB.Label label1 
            Alignment       =   2  'Center
            BackColor       =   &H80000002&
            BackStyle       =   0  'Transparent
            Caption         =   "DIFERENCIA"
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000009&
            Height          =   210
            Index           =   9
            Left            =   3885
            TabIndex        =   24
            Top             =   1215
            Width           =   1290
         End
         Begin VB.Label label1 
            Alignment       =   2  'Center
            BackColor       =   &H80000002&
            BackStyle       =   0  'Transparent
            Caption         =   "MONTO"
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000009&
            Height          =   210
            Index           =   8
            Left            =   2595
            TabIndex        =   25
            Top             =   1200
            Width           =   1290
         End
         Begin VB.Label label1 
            Alignment       =   2  'Center
            BackColor       =   &H80000002&
            BackStyle       =   0  'Transparent
            Caption         =   "# DOC."
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000009&
            Height          =   210
            Index           =   6
            Left            =   1845
            TabIndex        =   26
            Top             =   1200
            Width           =   750
         End
         Begin VB.Label label1 
            Alignment       =   2  'Center
            BackColor       =   &H80000002&
            Caption         =   "&ALICUOTA"
            ForeColor       =   &H80000009&
            Height          =   210
            Index           =   3
            Left            =   6900
            TabIndex        =   6
            Top             =   360
            Width           =   810
         End
         Begin VB.Shape Shape1 
            BorderColor     =   &H00FFFFFF&
            FillColor       =   &H80000002&
            FillStyle       =   0  'Solid
            Height          =   855
            Index           =   1
            Left            =   195
            Shape           =   4  'Rounded Rectangle
            Top             =   1140
            Width           =   5400
         End
         Begin VB.Label label1 
            Alignment       =   2  'Center
            BackColor       =   &H80000002&
            Caption         =   "COM&UN"
            ForeColor       =   &H80000009&
            Height          =   210
            Index           =   2
            Left            =   6105
            TabIndex        =   4
            Top             =   360
            Width           =   810
         End
         Begin VB.Label label1 
            Alignment       =   2  'Center
            BackColor       =   &H80000002&
            Caption         =   "&CUENTA"
            ForeColor       =   &H80000009&
            Height          =   210
            Index           =   0
            Left            =   210
            TabIndex        =   0
            Top             =   360
            Width           =   1455
         End
         Begin VB.Label label1 
            Alignment       =   2  'Center
            BackColor       =   &H80000002&
            Caption         =   "DE&SCRIPCION"
            ForeColor       =   &H80000009&
            Height          =   210
            Index           =   1
            Left            =   1665
            TabIndex        =   2
            Top             =   360
            Width           =   4440
         End
         Begin VB.Label label1 
            Alignment       =   2  'Center
            BackColor       =   &H80000002&
            Caption         =   "&MONTO"
            ForeColor       =   &H80000009&
            Height          =   210
            Index           =   4
            Left            =   7665
            TabIndex        =   8
            Top             =   360
            Width           =   1740
         End
         Begin VB.Label label1 
            Alignment       =   2  'Center
            BackColor       =   &H80000002&
            Caption         =   "PERIODO"
            ForeColor       =   &H80000009&
            Height          =   210
            Index           =   5
            Left            =   9405
            TabIndex        =   10
            Top             =   360
            Width           =   1410
         End
         Begin VB.Shape Shape1 
            BorderColor     =   &H00FFFFFF&
            FillColor       =   &H80000002&
            FillStyle       =   0  'Solid
            Height          =   855
            Index           =   2
            Left            =   5670
            Shape           =   4  'Rounded Rectangle
            Top             =   1155
            Width           =   5175
         End
         Begin VB.Shape Shape1 
            BorderColor     =   &H00808080&
            BorderWidth     =   2
            Height          =   870
            Index           =   0
            Left            =   195
            Shape           =   4  'Rounded Rectangle
            Top             =   1140
            Width           =   5415
         End
         Begin VB.Shape Shape1 
            BorderColor     =   &H00808080&
            BorderWidth     =   2
            Height          =   870
            Index           =   3
            Left            =   5670
            Shape           =   4  'Rounded Rectangle
            Top             =   1155
            Width           =   5190
         End
      End
      Begin VB.Frame fraAgasto 
         Caption         =   "Detalle Asignacion de Gasto:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   3795
         Index           =   1
         Left            =   120
         TabIndex        =   14
         Top             =   2820
         Width           =   11070
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid GridGastos 
            Height          =   3090
            Index           =   0
            Left            =   180
            TabIndex        =   15
            Tag             =   "1200|4700|900|900|1500|1300"
            Top             =   480
            Width           =   10725
            _ExtentX        =   18918
            _ExtentY        =   5450
            _Version        =   393216
            Cols            =   6
            FixedCols       =   0
            BackColorFixed  =   -2147483646
            ForeColorFixed  =   -2147483639
            BackColorSel    =   65280
            AllowBigSelection=   0   'False
            BorderStyle     =   0
            FormatString    =   "^Cuenta |Descripción |Común|Alícuota |>Monto |Período"
            BandDisplay     =   1
            _NumberOfBands  =   1
            _Band(0).Cols   =   6
         End
         Begin VB.Image ImgOK 
            Enabled         =   0   'False
            Height          =   240
            Index           =   1
            Left            =   0
            Top             =   720
            Visible         =   0   'False
            Width           =   240
         End
         Begin VB.Image ImgOK 
            Enabled         =   0   'False
            Height          =   240
            Index           =   0
            Left            =   30
            Top             =   240
            Visible         =   0   'False
            Width           =   240
         End
      End
      Begin MSDataGridLib.DataGrid DataGrid1 
         Height          =   4365
         Left            =   -74730
         TabIndex        =   31
         Top             =   675
         Width           =   10725
         _ExtentX        =   18918
         _ExtentY        =   7699
         _Version        =   393216
         AllowUpdate     =   0   'False
         HeadLines       =   1
         RowHeight       =   15
         RowDividerStyle =   3
         FormatLocked    =   -1  'True
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
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
         Caption         =   "LISTADO DE NOVEDADES"
         ColumnCount     =   8
         BeginProperty Column00 
            DataField       =   "Ndoc"
            Caption         =   "Doc"
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
         BeginProperty Column01 
            DataField       =   "CodGasto"
            Caption         =   "Gasto"
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
            DataField       =   "Detalle"
            Caption         =   "Descripción"
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
            DataField       =   "Periodo"
            Caption         =   "Periodo"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   "MMM-YYYY"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   8202
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column04 
            DataField       =   "Monto"
            Caption         =   "Monto"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   "#,##0.00 "
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   8202
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column05 
            DataField       =   "Fecha"
            Caption         =   "Fecha"
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
         BeginProperty Column06 
            DataField       =   "Hora"
            Caption         =   "Hora"
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
         BeginProperty Column07 
            DataField       =   "Usuario"
            Caption         =   "Usuario"
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
               DividerStyle    =   3
               ColumnWidth     =   14,74
            EndProperty
            BeginProperty Column01 
               Alignment       =   2
               DividerStyle    =   5
               ColumnWidth     =   1049,953
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   4559,811
            EndProperty
            BeginProperty Column03 
               ColumnWidth     =   1049,953
            EndProperty
            BeginProperty Column04 
               Alignment       =   1
               ColumnWidth     =   1454,74
            EndProperty
            BeginProperty Column05 
               ColumnWidth     =   1154,835
            EndProperty
            BeginProperty Column06 
               ColumnWidth     =   1260,284
            EndProperty
            BeginProperty Column07 
            EndProperty
         EndProperty
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
            Picture         =   "frmNovedades.frx":07AA
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmNovedades.frx":092C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmNovedades.frx":0AAE
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmNovedades.frx":0C30
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmNovedades.frx":0DB2
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmNovedades.frx":0F34
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmNovedades.frx":10B6
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmNovedades.frx":1238
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmNovedades.frx":13BA
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmNovedades.frx":153C
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmNovedades.frx":16BE
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmNovedades.frx":1840
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmNovedades"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
    'Módulo Facturacion------SINAI TECH,C.A
    '*********************************************************************************************
    '   Asginar Gastos ocasionales, que no provienen de facturas, no son gastos fijos
    '   del edificio, no son gastos no comunes
    '*********************************************************************************************
    Dim rstNov(4) As New ADODB.Recordset   'Conjunto de ADODB.Recordsets
    Private Enum Novedades
        Tgastos
        Fact
        mPeriodo
        Nov
    End Enum
    Dim CnnInm As New ADODB.Connection  'Conexion a los datos del inmueble seleccionado
    Dim i As Integer    'Variable entera de iteracion
    Dim datPeriodo As Date  'Ultimo período facturado
    Dim vDistribucion()
    Dim Fila As Integer
        
    Private Sub Check1_Click(index As Integer)
    On Error Resume Next
    Select Case index
        Case 2, 3
            ReDim vDistribucion(0, 0)
            TxtAGasto(4).SetFocus
            fraAgasto(4).Visible = False
        Case 4: Call Reintegro
        Case 5: TxtAGasto(8).Enabled = Check1(5).Value
    End Select
    End Sub


Private Sub Check1_MouseDown(index As Integer, Button As Integer, Shift As Integer, _
X As Single, Y As Single)
'
If Button = 2 Then      'Si presiona el segundo boton del mouse
        Select Case index   'muestra un menu emergente
    '
            Case 2  'Gasto Comun / No_Comun
            '------------------------------
                
                If Check1(2).Value = False Then
                    cmdHelp(0).Tag = 1
                    Option1(0).Visible = True
                    Option1(1).Visible = True
                    fraAgasto(4).Visible = True
                    LisAGasto(0).Visible = True
                    LisAGasto(1).Visible = True
                    cmdHelp(1).Visible = True
                    cmdHelp(3).Visible = True
                    GridGastos(2).Visible = False
                    If Option1(0) Then
                        fraAgasto(4).Height = 960
                    Else
                        fraAgasto(4).Height = 3405
                    End If
                    '
                End If
    '
        End Select
    '
    End If

End Sub

Private Sub CmdHelp_Click(index As Integer)
'VER EL CARGADO DE LA FACTURA
If index = 0 Then
    If cmdHelp(0).Tag = 1 Then ReDim vDistribucion(0, 0)
    
    If Option1(0) Then
    
    Else
        If LisAGasto(1).ListCount < 1 Then
            MsgBox "Debe seleccionar por lo menos un propietario", vbInformation, App.ProductName
            Exit Sub
        End If
    End If
    
    fraAgasto(4).Visible = False
    With TxtAGasto(4)
        .SelStart = 0
        .SelLength = Len(.Text)
        .SetFocus
    End With
ElseIf index = 1 Then   'pasar todo
    For i = 0 To LisAGasto(0).ListCount - 1
        LisAGasto(1).AddItem LisAGasto(0).List(i)
    Next
    LisAGasto(0).Clear
ElseIf index = 3 Then   'regresar todo
    For i = 0 To LisAGasto(1).ListCount - 1
        LisAGasto(0).AddItem LisAGasto(1).List(i)
        
    Next
    LisAGasto(1).Clear
End If
End Sub

Private Sub cmdNov_Click()
'Asigna una novedad a una factura recibida
Dim strMsg$
If TxtAGasto(14) = "" Then
    MsgBox "Debe Ingresar el Nº de Documento...", vbInformation, App.ProductName
    TxtAGasto(14).SetFocus
Else
    If Validar Then Exit Sub
    If Not rstNov(Nov).EOF Or Not rstNov(Nov).BOF Then
    '
        strMsg = "Seguro desea relacionar la siguiente novedad:" & vbCrLf & "Gasto: " & _
        rstNov(Nov)!codGasto & "" & vbCrLf & "Por Bs." & Format(rstNov(Nov)!Monto, "#,##0.00") & _
        vbCrLf & "Cargado a: " & UCase(Format(rstNov(Nov)!Periodo, "MMM-YYYY")) & vbCrLf & "Con" _
        & " el documento Nº" & TxtAGasto(14)
        'si el usuario presiona 'SI'
        If Respuesta(strMsg) Then
            'Actualiza AsigaGasto/GastonoComun
            CnnInm.Execute "UPDATE AsignaGasto SET Ndoc='" & TxtAGasto(14) & "' WHERE Ndoc='NOV" _
            & "' AND CodGasto='" & rstNov(Nov)!codGasto & "' AND Monto=" & rstNov(Nov)!Monto
            'Actualiza el estatus del documento
            cnnConexion.Execute "UPDATE Cpp SET Estatus='ASIGNADO' WHERE Ndoc='" & _
            TxtAGasto(14) & "'"
            rstNov(Nov).Update "Ndoc", TxtAGasto(14)
            rstNov(Nov).Requery '
            
        End If
        '
    Else
        MsgBox "No Existen Novedades Registradas para este Condominio...", vbInformation, _
        App.ProductName
        '
    End If
    '
End If
'
End Sub

    Private Sub Combo1_KeyPress(index As Integer, KeyAscii As Integer)
    '
    Dim Bolos@   'variables locales
    Dim Concepto$
    Dim fecha As Date
    
    If KeyAscii = 13 Then   'Presionó {enter}
    '
        Select Case index   'Ejecuta el procedimiento de acuero al
    '                        control que hace el llamdo
            Case 0  'Lista de los Meses
    '
                If Combo1(0) = "" Then
                    MsgBox "Seleccione un Mes..", vbExclamation, App.ProductName
                    Exit Sub
                End If
                Combo1(1) = Format(Date, "YY")
                Combo1(1).SetFocus
                
            Case 1  'Combo del año {Guarda el registro actual}
    '
                If Validar_Guardar Then Exit Sub
                Concepto = TxtAGasto(1)
                If Check1(4) Then TxtAGasto(7) = Format(CCur(TxtAGasto(7)) + CCur(TxtAGasto(4)), "#,##0.00")
                fecha = "01/" & Combo1(0) & "/" & Combo1(1)
                If Check1(5).Value = 1 Then 'Distribuye en períodos
                    Bolos = CSng(TxtAGasto(4) / TxtAGasto(8)) 'divide el monto entre los periodos
                    If Check1(6).Value = 1 Then 'agrega el final n/x cuotas
                        For Z = 1 To CInt(TxtAGasto(8))
                            Call RtnGuardar(Bolos, fecha, Concepto & " " & Z & "/" & TxtAGasto(8))
                            fecha = DateAdd("m", 1, fecha)
                            Combo1(0) = Format(fecha, "mmm")
                            Combo1(1) = Format(fecha, "yy")
                        Next
                        Call Inicializar
                    End If
                Else    'es un solo registro
                    Bolos = TxtAGasto(4)
                    Call RtnGuardar(Bolos, fecha, Concepto)
                End If
                '
                CnnInm.Execute "UPDATE Cpp IN '" & gcPath & "\sac.mdb' SET ESTATUS='REINTEGRADO" _
                & "', Usuario='" & gcUsuario & "',Freg=Date() WHERE Ndoc='" & TxtAGasto(5) & "'"
                '
        End Select
    '
    End If

    End Sub

    Private Sub Form_Load()
    
    ReDim vDistribucion(0, 0)
    ImgOK(0).Picture = LoadResPicture("Unchecked", vbResBitmap)
    ImgOK(1).Picture = LoadResPicture("Checked", vbResBitmap)
    '
    For i = 0 To 2
        Combo1(1).AddItem (Format(DateAdd("yyyy", -1 + i, Date), "yy"))
        rstNov(i).CursorLocation = adUseClient
        rstNov(i + 1).CursorLocation = adUseClient
    Next i
    CnnInm.Open cnnOLEDB + mcDatos  'abre la conexion a los datos del inmueble
    'selecciona todos los registros del catálogo de gastos
    
    rstNov(Tgastos).Open "TGASTOS", CnnInm, adOpenStatic, adLockReadOnly, adCmdTable
    'rstNov(Tgastos).Sort = "CodGasto"    'los ordena por código
    '
    'Selecciona las facturas del inmueble que tengan el estatus 'ASIGNADO'
    rstNov(Fact).Open "SELECT * FROM Cpp WHERE Codinm='" & gcCodInm & "' AND Estatus='ASIGNADO' " _
    & "ORDER BY Freg;", cnnConexion, adOpenKeyset, adLockOptimistic, adCmdText
    '
    'Selecciona el último período facturado
    rstNov(mPeriodo).Open "SELECT MAX(Periodo) FROM Factura where fACT nOT Like 'CHD%';", _
    CnnInm, adOpenStatic, adLockReadOnly, adCmdText
    '
    If Not IsNull(rstNov(mPeriodo).Fields(0)) Then
        datPeriodo = rstNov(mPeriodo).Fields(0)
    Else
        datPeriodo = DateAdd("YYYY", -10, Date)
    End If
    '
    With Toolbar1
        .Buttons("Save").Enabled = False
        .Buttons("Undo").Enabled = False
    End With
    rstNov(Nov).Open "SELECT * FROM Cargado WHERE Ndoc='NOV'", CnnInm, adOpenKeyset, _
    adLockOptimistic, adCmdText
    If rstNov(Nov).EOF And rstNov(Nov).BOF Then Toolbar1.Buttons("Delete").Enabled = False
    '---
    Set DataGrid1.DataSource = rstNov(Nov)
    Set GridGastos(0).FontFixed = LetraTitulo(LoadResString(527), 7.5)
    Set GridGastos(0).Font = LetraTitulo(LoadResString(528), 8)
    Call centra_titulo(GridGastos(0), True)
    '
    End Sub

    Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    For i = 0 To 4
        rstNov(i).Close
        Set rstNov(i) = Nothing
    Next
    CnnInm.Close
    Set CnnInm = Nothing
    End Sub


Private Sub GridGastos_EnterCell(index As Integer)

    If index = 2 And GridGastos(2).Col = 2 And GridGastos(2).RowSel > 0 Then
        With GridGastos(2)
            
            .Text = CCur(GridGastos(2).Text)
            Dim curMonto As Currency
            For i = 1 To .Rows - 1
                If Not IsNumeric(.TextMatrix(i, 2)) Then .TextMatrix(i, 2) = 0
                
                curMonto = curMonto + .TextMatrix(i, 2)
            Next
            vDistribucion(0, 0) = curMonto
            TxtAGasto(4) = Format(curMonto, "#,##0.00")
            'Debug.Print TxtAGasto(4)
        End With
    End If

End Sub

Private Sub GridGastos_KeyPress(index As Integer, KeyAscii As Integer)

If index = 2 And GridGastos(2).Col = 2 Then
    '
        With GridGastos(2)
    '
            If .RowSel > 0 Then
    '
                If KeyAscii = 46 Then KeyAscii = 44 'CONVIERTE PUNTO(.) EN COMA(,)
                
                Call Validacion(KeyAscii, "0123456789,0-")
                If KeyAscii = 8 Then KeyAscii = 0
                
    '
                If KeyAscii = 13 Then
    '
                    GridGastos_LeaveCell 2
                    If .RowSel = (.Rows - 1) Then
    '                   Si esta en la última cuadricula vuelve a la priemra
                        .Row = 1
                    Else
    '                   Baja a la cuadricula inferior siguiente
                        .Row = (.RowSel + 1)
                    End If
                    GridGastos_EnterCell 2
                    GridGastos(2).SetFocus
                    GridGastos(2).TopRow = IIf(GridGastos(2).Row - 17 < 1, 1, GridGastos(2).Row - 17)
    '
                Else
                    If Fila <> .Row Then
                    .TextMatrix(.RowSel, 2) = Chr(KeyAscii)
                Else
                    .TextMatrix(.RowSel, 2) = _
                    .TextMatrix(.RowSel, 2) & Chr(KeyAscii)
                End If
                Fila = .Row
                End If
    '
            End If
    '
        End With
    '
    End If
End Sub

Private Sub GridGastos_KeyUp(index As Integer, KeyCode As Integer, Shift As Integer)
    'Permite Borrar el contenido de la celda (total/parcial)
    If index = 2 And GridGastos(2).Col = 2 And GridGastos(2).RowSel > 0 Then
    '
        With GridGastos(2)
    '
            Select Case KeyCode 'Selecciona la tecla que presionó
    '       -----------------------------------------------------
                Case 46 '{SUPRIMIR} 'Borra todo
    '           --------------------------------
                    '.TextMatrix(.RowSel, 2) = ""
                    .Text = ""
    '
                Case 8  '{BACKSPACE} 'Borra el último caracter
    '           ----------------------------------------------
                    If Len(.Text) > 0 Then
                        .Text = Left(.Text, Len(.Text) - 1)
                    End If
    '
            End Select
    '
        End With
    '
    End If

End Sub

Private Sub GridGastos_LeaveCell(index As Integer)
    If index = 2 And GridGastos(2).Col = 2 Then
        Fila = GridGastos(2).Row
        If Not IsNumeric(GridGastos(2).Text) Then GridGastos(2).Text = "0,00"
        vDistribucion(Fila, 1) = GridGastos(2).Text
        GridGastos(2).Text = Format(GridGastos(2).Text, "#,##0.00")
    End If
End Sub

Private Sub LisAGasto_DblClick(index As Integer)
Select Case index
        Case 0  'Lista de Todos los propietarios
    '   -----------------------------------------------
            LisAGasto(1).AddItem (LisAGasto(0).Text)
            LisAGasto(0).RemoveItem (LisAGasto(0).ListIndex)
        
        Case 1  'Lista de propietarios pre-seleccionados
    '   -------------------------------------------------
            LisAGasto(0).AddItem (LisAGasto(1).Text)
            LisAGasto(1).RemoveItem (LisAGasto(1).ListIndex)
            
    '
    End Select
End Sub

    Private Sub Option1_Click(index As Integer)
    Select Case index
    '----------------
        ReDim vDistribucion(0, 0)
        'Todos los propietarios
        Case 0: fraAgasto(4).Height = 960
        'Determinados propietarios
        Case 1
            If LisAGasto(0).ListCount = 0 And LisAGasto(1).ListCount = 0 Then
            Set AdogastosTran = New ADODB.Recordset
            
            With AdogastosTran
    '           Selecciona todos los codigos de propietarios del inmueble
    '           y llena la lista para que el usuario realice su selección
                .Open "SELECT Codigo,Nombre  FROM Propietarios WHERE Codigo <> 'U" & gcCodInm _
                & "' ORDER BY Codigo", cnnOLEDB + gcPath + gcUbica + "inm.mdb", adOpenKeyset, adLockReadOnly
                If .EOF Then Exit Sub
                .MoveFirst
                LisAGasto(0).Clear
                Do Until .EOF
                    LisAGasto(0).AddItem (!Codigo)
                    .MoveNext
                Loop
                AdogastosTran.Close
                Set AdogastosTran = Nothing
    '
            End With
            End If
            If fraAgasto(4).Height < 3405 Then
                fraAgasto(4).Height = 3405
            End If
    '
    End Select
    '---------

    End Sub

    Private Sub OptNov_Click(index%): rstNov(Nov).Sort = OptNov(index).Tag
    End Sub


    Private Sub SSTab1_Click(PreviousTab As Integer)
    If SSTab1.Tab = 0 Then
            Toolbar1.Buttons("Delete").Enabled = False
        Else
            If Not rstNov(Nov).EOF And Not rstNov(Nov).BOF Then
                Toolbar1.Buttons("Delete").Enabled = True
            End If
        End If
    End Sub

    Private Sub SSTab1_DblClick()
    If SSTab1.Tab = 0 Then
        Toolbar1.Buttons("Delete").Enabled = False
    Else
        Toolbar1.Buttons("Delete").Enabled = True
    End If
    End Sub

    Private Sub Toolbar1_ButtonClick(ByVal Button As ComctlLib.Button)
    '
    Dim msg As String
    Dim Hora As Date
    '
    Select Case Button.Key
        Case "New"
            fraAgasto(2).Enabled = True
            TxtAGasto(0).SetFocus
            SSTab1.Tab = 0
            CnnInm.BeginTrans
            
        Case "Save"
            'valida que la diferencia final sea igual a cero
            If Check1(4) Then
                If CCur(TxtAGasto(7)) <> 0 Then
                    MsgBox "Imposible guardar esta operación." & vbCrLf & "Restan por asignar " _
                    & IDMoneda & " " & TxtAGasto(7), vbExclamation, App.ProductName
                    Exit Sub
                End If
            End If
            fraAgasto(2).Enabled = False
            If Respuesta("¿Desea guardar toda la operación?") Then
                CnnInm.CommitTrans
                For i = 0 To 10000  'retardo
                Next
                rstNov(Nov).Requery
            Else
                CnnInm.RollbackTrans
            End If
            Call Inicializar
            Call grid_clear
            
            
        Case "Undo"
            Call Inicializar
            Call grid_clear
            fraAgasto(4).Visible = False
            fraAgasto(2).Enabled = False
            CnnInm.RollbackTrans
            
        Case "Find"
            SSTab1.Tab = 1
            
        Case "Print": Call Printer_Report
        
        Case "Close"
            Unload Me
        
        Case "Delete"
            'si el período ya fué facturado sale de la rutina
            If rstNov(Nov)!Periodo <= datPeriodo Then
                MsgBox "Imposible eliminar esta novedad. Período ya facturado!!!", _
                vbExclamation, App.ProductName
                Exit Sub
            End If
            msg = "¿Desea eliminar la siguiente novedad?" & vbCrLf & "Gasto: " & _
            rstNov(Nov)!codGasto & vbCrLf & "Descripción: " & rstNov(Nov)!Detalle & vbCrLf _
            & "Cargado a: " & Format(rstNov(Nov)!Periodo, "mm-yyyy") & vbCrLf _
            & "Monto: " & Format(rstNov(Nov)!Monto, "#,##0.00")
            '
            If Respuesta(msg) Then
            
                On Error Resume Next
                
                CnnInm.BeginTrans
                'ACTUALIZA EL ESTATUS DE LA LA TABLA CPP
                CnnInm.Execute "UPDATE Cpp IN '" & gcPath & "\SAC.MDB' SET Estatus='ASIGNADO', " _
                & "Usuario='" & gcUsuario & "', Freg = DATE() WHERE Estatus = 'REINTEGRADO' and" _
                & " Total = " & Replace(Abs(rstNov(Nov)!Monto), ",", ".") & "  AND CodInm='" & gcCodInm & "'", K
                
                'elimina la inf. de la tabla asignagasto
                CnnInm.Execute "DELETE * FROM AsignaGasto WHERE Ndoc='NOV' AND CodGasto='" & _
                rstNov(Nov)!codGasto & "' AND Descripcion='" & rstNov(Nov)!Detalle & "' AND Car" _
                & "gado=#" & Format(rstNov(Nov)!Periodo, "mm/dd/yy") & "# AND Monto * 100 =" & _
                rstNov(Nov)!Monto * 100, K
                'elimina la inf. de la tabla cargado
                CnnInm.Execute "DELETE * FROM Cargado WHERE Ndoc='NOV' AND CodGasto='" & _
                rstNov(Nov)!codGasto & "' AND Detalle='" & rstNov(Nov)!Detalle & "' AND Periodo" _
                & "=#" & Format(rstNov(Nov)!Periodo, "mm/dd/yy") & "# AND Monto * 100 =" & _
                rstNov(Nov)!Monto * 100, K
                '
                'aqui elimina los gastos no comunes (si es el caso)
                CnnInm.Execute "DELETE * FROM GastoNoComun WHERE PF=0 AND CodGasto='" & _
                rstNov(Nov)!codGasto & "' AND Concepto='" & rstNov(Nov)!Detalle & _
                "' and Periodo=#" & Format(rstNov(Nov)!Periodo, "mm/dd/yy") & _
                "# AND Usuario='" & rstNov(Nov)!Usuario & "' AND Left(Hora,5)='" & _
                Left(rstNov(Nov)!Hora, 5) & "'", K
                
                If Err.Number = 0 Then  'no ocurrion ningún error
                    CnnInm.CommitTrans
                    MsgBox "Información eliminada"
                    For i = 0 To 10000  'retardo
                    Next
                Else
                    CnnInm.RollbackTrans
                    MsgBox Err.Description, vbCritical, "Error " & Err.Number
                End If
                '
            End If
            '
    End Select
    Call RtnEstado(Button.index, Toolbar1)
    '
    End Sub

Private Sub TxtAGasto_DblClick(index As Integer)
If TxtAGasto(4) = "" Then Exit Sub

If Check1(2).Value = 0 And index = 4 And TxtAGasto(4) <> "" Then
    cmdHelp(0).Tag = 0
    Option1(0).Visible = False
    GridGastos(2).Visible = True
    fraAgasto(4).Visible = True
    LisAGasto(0).Visible = False
    LisAGasto(1).Visible = False
    cmdHelp(1).Visible = False
    cmdHelp(3).Visible = False
    cmdHelp(2).BackColor = &H80000004
    TxtAGasto(4) = Format(TxtAGasto(4), "#,##0.00#")
    With GridGastos(2)  'Configura el diseño del Grid
    
        .TextArray(0) = "Codigo": .ColWidth(0) = 700
        .TextArray(1) = "Propietario": .ColWidth(1) = 1350
        .TextArray(2) = "Monto"
        Call centra_titulo(GridGastos(2))
       .ColAlignment(0) = flexAlignCenterCenter
       .ColAlignment(1) = flexAlignCenterCenter
        Call RtnGrid
        .Refresh
        
    End With
    
End If
End Sub

    Private Sub TxtAGasto_GotFocus(index As Integer)
    If IsNumeric(TxtAGasto(4)) Then TxtAGasto(4) = CCur(TxtAGasto(4))
    End Sub

    Private Sub TxtAGasto_KeyPress(index As Integer, KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 39 Then KeyAscii = 0
    
    Select Case index
    '
        Case 0
            Call Validacion(KeyAscii, "0123456789")
            If KeyAscii = 13 Then Call Busca_gasto("CodGasto='" & TxtAGasto(0) & "'")
            
        Case 1
            If KeyAscii = 13 Then Call Busca_gasto("Titulo='" & TxtAGasto(1) & "'")
            
        Case 5: If KeyAscii = 13 Then Call Busca_Fact
        
        Case 4: If KeyAscii = 13 Then SendKeys vbTab
        
        Case 9
            Call Validacion(KeyAscii, "1234567890")
            If KeyAscii = 13 Then Call rtnFiltrar
            
        Case 10, 11, 12
            Call Validacion(KeyAscii, "0123456789/")
            If KeyAscii = 13 Then Call rtnFiltrar
        
        Case 14: Call Validacion(KeyAscii, "0123456789")
            
    End Select
    '
    End Sub

    '---------------------------------------------------------------------------------------------
    '   Rutina: Busca_Fact
    '
    '   Busca una factura determinada dentro del depósito de facturas asignadas
    '---------------------------------------------------------------------------------------------
    Private Sub Busca_Fact()
    '
    With rstNov(Fact)
        '
        .Requery    'recupera el conjunta de registros
        .Find "Ndoc='" & TxtAGasto(5) & "'" 'busca la factura determinada
        If Not .EOF Then    'si hay coincidencia
            TxtAGasto(6) = Format(!Total, "#,##0.00")   'asigna valor a los controles
            TxtAGasto(7) = Format(!Total, "#,##0.00")
        Else    'no existe coincidencia
            MsgBox "Verifique el Número de Documento que introdujo. Para el Inmueble '" & _
            gcCodInm & "' NO EXISTE esa factura registrada, o su estatus no es 'ASIGNADO'", _
            vbInformation, App.ProductName
        End If
    End With
    '
    End Sub

    Private Sub Reintegro()
    Dim Estado As Boolean
    Estado = Check1(4).Value
    For i = 5 To 7
        TxtAGasto(i).Enabled = Estado
        TxtAGasto(i) = ""
    Next i
    cmdHelp(2).Enabled = Estado
    End Sub

    Private Sub TxtAGasto_KeyUp(index%, KeyCode%, Shift%)
    'variables locales
    Dim intC As Integer
    Dim StrP As String
    '
    If index = 1 Then
        If KeyCode = 8 Then KeyCode = 46
        If KeyCode = 46 Or TxtAGasto(1) = "" Then Exit Sub
        intC = TxtAGasto(1).SelStart
        If intC <= 0 Then Exit Sub
        StrP = Left(TxtAGasto(1), intC)
        If StrP = "" Then Exit Sub
        With rstNov(Tgastos)
                .Find "Titulo Like '" & StrP & "*'"
                If .EOF = True Then
                    If .BOF = True Then Exit Sub
                    'Buca desde el principio
                    .MoveFirst
                    .Find "Titulo Like '" & StrP & "*'"
                    If Not .EOF Then GoTo Marcar
                Else
Marcar:        With TxtAGasto(1)
                        .Text = rstNov(Tgastos)!Titulo
                        .SelStart = intC
                        .SelLength = Len(.Text)
                        .SetFocus
                    End With
                End If
        End With
        '
    End If
    End Sub

    Private Sub Busca_gasto(Criterio As String)
    With rstNov(Tgastos)
    '
        .Requery
        .Find Criterio
        If .EOF Then
            MsgBox "Gasto NO ENCONTRADO", vbInformation, App.ProductName
        Else
            TxtAGasto(0) = !codGasto
            TxtAGasto(1) = IIf(IsNull(!Titulo), "--VACIO--", !Titulo)
            Check1(2).Value = IIf(!Comun, 1, 0)
            Check1(3).Value = IIf(!Alicuota, 1, 0)
            TxtAGasto(4).SetFocus
        End If
    '
    End With
    End Sub

'---------------------------------------------------------------------------------------------
    '   Function:   Validar_Guardar
    '
    '   Función que verifica los datos necesarios para procesar el registro
    '   periodo facturado, campos nulos, etc. Si todo es correcto devuelve
    '   un valor false, si encuenta un error devuelve True
    '---------------------------------------------------------------------------------------------
    Private Function Validar_Guardar() As Boolean
    'variables locales
    Dim strMensaje As String
    Dim datCargado As Date
    '
    datCargado = "01-" + Combo1(0) + "-" + Combo1(1)
    If Combo1(0) = "" Or Combo1(1) = "" Then 'Sale de la rutina si esta vacio
        strMensaje = "Verifique los datos del Período"
    ElseIf TxtAGasto(0) = "" Or TxtAGasto(1) = "" Then
        strMensaje = "Verifique el Código y/o la Descripción de la Cuenta"
    ElseIf TxtAGasto(4) = "" Then
        strMensaje = strMensaje + "Debe Introducir una cantidad" & vbCrLf
    ElseIf datCargado <= datPeriodo Then
        strMensaje = "Imposible Asignar el Gasto a un Período Facturado...."
    ElseIf Check1(6).Value = 1 And Check1(5).Value = 0 Then
        strMensaje = "La opción 'Cuota X/N' no puede estar seleccionada si no selecciona 'divid" _
        & "ir en X períodos'"
'    ElseIf Check1(4) And CCur(TxtAGasto(4)) > CCur(TxtAGasto(7)) Then
'        strMensaje = "El monto de este gasto sobrepasa el resto a distribuir de la factura"
    End If
    '
    If strMensaje <> "" Then Validar_Guardar = MsgBox(strMensaje, vbCritical, App.ProductName)
    '
    End Function

    '---------------------------------------------------------------------------------------------
    '   Rutina:     RtnGuardar
    '
    '   Rutina que guarda cada registro según la distribución del gassto: comun /
    '   No común. Por Alícuota o Partes Iguales. Para Propietarios Determinados.
    '   Montos Específicos. De acuerdo a la selección del usuario
    '---------------------------------------------------------------------------------------------
    Private Sub RtnGuardar(Monto@, Periodo As Date, Descripcion$) '
    '
    On Error GoTo salir
    
    If booFondo(TxtAGasto(0)) = True Then 'Si el gasto afecta el Fondo de Reserva
    '
        CnnInm.Execute "INSERT INTO MovFondo(CodGasto,Fecha,Tipo,Periodo,Concepto,Debe,Haber) VALU" _
        & "ES('" & TxtAGasto(0) & "',Date() ,'ND','01/" & Format(Date, "mm/yyyy") & "','" & _
        Descripcion & "','" & Abs(Monto) & "',0)"
        'Actuliza el Saldo Actual
        CnnInm.Execute "UPDATE Tgastos SET SaldoActual = SaldoActual - '" & Monto & "', Freg=DA" _
        & "TE(),Usuario='" & gcUsuario & "' WHERE CodGasto='" & TxtAGasto(0) & "';"
        'Agrega el cargado
        CnnInm.Execute "INSERT INTO Cargado (Ndoc,CodGasto,Detalle,Monto,Fecha,Hora,Usuario,Period" _
        & "o) VALUES('NOV','" & TxtAGasto(0) & "','" & Descripcion & "','" & _
        Monto & "',Date(),Time() ,'" & gcUsuario & "','" & Periodo & "');"
        
    Else
        CnnInm.Execute "INSERT INTO Cargado (Ndoc,CodGasto,Detalle,Periodo,Monto,Fecha,Hora,Usuari" _
        & "o) VALUES('NOV','" & TxtAGasto(0) & "','" & Descripcion & "','" & _
        Periodo & "','" & Monto & "',Date(),Time() ,'" & gcUsuario & "')"
        
        If TxtAGasto(0) = "100025" Or TxtAGasto(0) = "101025" Or TxtAGasto(0) = "102025" Then GoTo 20
    '
        If gcCodInm = sysCodInm Then Check1(2).Value = 1
        
        If Check1(2).Value = 0 Then  'Gasto {NoComun}
            
'            Call RtnAnexarReg(Check1(3).Value, Monto, Periodo, Descripcion)
            If Option1(0).Value = True Then 'Todos los Propietarios
                If Check1(3).Value = 1 Then  'Aplicado por alícuota
                    Call RtnAnexarReg("TRUE")
                ElseIf Check1(3).Value = 0 Then 'Aplicado en partes iguales
                    Call RtnAnexarReg("FALSE")
                End If
            ElseIf Option1(1).Value = True Then 'Propietarios Seleccionados
                With GridGastos(2)
        '       ----------------------|
                For i = 1 To .Rows - 1
                    CnnInm.Execute "INSERT INTO GastoNoComun (CodApto,CodGasto,Concepto,Monto,Periodo," _
                    & "Fecha,Hora,Usuario) VALUES ('" & LisAGasto(1).List(i - 1) & "','" _
                    & TxtAGasto(0) & "','" & TxtAGasto(1) & "','" & .TextMatrix(i, 2) & "','" _
                    & Periodo & "',Date(),Time() ,'" & gcUsuario & "')"
                Next
        '       ---------------------|
                End With
        '
            End If
            
        Else    'Gasto Comun
            'Guarda el registro en la Tabla AsignaGasto
            CnnInm.Execute "INSERT INTO AsignaGasto (Ndoc,Cargado,CodGasto,Descripcion,Comun, " _
            & "Alicuota,Monto,Usuario,Fecha,Hora) VALUES ('NOV','01/" _
            & Combo1(0) & "/" & Combo1(1) & "','" & TxtAGasto(0) & "','" & Descripcion & _
            "'," & Check1(2) & "," & Check1(3) & ",'" & Monto & "','" & gcUsuario _
            & "',Date() ,Time())"
            '
        End If
        '
    End If
salir:
    If Err.Number = 0 Then
20    Call Agregar_Gasto(Monto)
    Else
        MsgBox Err.Description, vbCritical, "Error " & Err.Number
    End If
    '
    End Sub


    '---------------------------------------------------------------------------------------------
    '   Rutina:     RtnAnexarReg
    '
    '   Entrada:    Variable boolLlamada. Por alicuota=1;Partes iguales=0
    '
    '   Agrega Nuevos registros a la tabla GastosNoComunes, según parametros
    '   enviados por el usuario,{alícuota / partes iguales}
    '---------------------------------------------------------------------------------------------
    Private Sub RtnAnexarReg(BooLlamada As Boolean) '-
    '
    Dim strMonto As String  'variables locales
    Dim datPeri As Date
    '
    datPeri = "01/" & Combo1(0) & "/" & Combo1(1)
    '------------------|
    If UBound(vDistribucion, 1) > 1 Then
        For i = 1 To UBound(vDistribucion, 1)
            strMonto = Replace(CCur(vDistribucion(i, 1)), ",", ".")
            CnnInm.Execute "INSERT INTO GastoNoComun (CodApto,CodGasto,Concepto,Monto,Periodo," _
            & "Fecha,Hora,Usuario) VALUES('" & vDistribucion(i, 0) & "','" & TxtAGasto(0) & "','" _
            & TxtAGasto(1) & "'" & ",'" & vDistribucion(i, 1) & "','" & datPeri & _
            "',Date(),Time(),'" & gcUsuario & "')"
        Next
    Else
        MsgBox "No se han cargado los gastos." & vbCrLf & _
        "Comuníquese con el administrador del sistema", vbInformation, App.ProductName
        Call rtnBitacora("No se distribuyeron los gastos")
    End If
    '
    End Sub


    '---------------------------------------------------------------------------------------------
    Private Sub Agregar_Gasto(Monto As Currency) '
    '---------------------------------------------------------------------------------------------
    
    'If Check1(4).Value = 1 Then
    '    TxtAGasto(7) = Format(CCur(TxtAGasto(7)) - Monto, "#,##0.00")
    'End If
    '           ------------------'envia el registro a la cuadrícula
    With GridGastos(0)
        .Row = .Rows - 1
        For i = 0 To 5
            Select Case Val(i)
                Case 0, 1
                    .TextMatrix(.Row, i) = TxtAGasto(i)
                    'TxtAGasto(i) = ""
                Case 2, 3
                    .Col = i
                    Set .CellPicture = IIf(Check1(i).Value = 1, ImgOK(1), ImgOK(0))
                    .CellPictureAlignment = flexAlignCenterCenter
                    
                Case 4
                    .TextMatrix(.Row, i) = Format(Monto, "#,##0.00 ")
                Case 5
                    .TextMatrix(.Row, i) = Combo1(0) + "-" + Combo1(1)
                    
            End Select
        Next
        GridGastos(0).AddItem ("")
        TxtAGasto(0).SetFocus
    '
    End With
    '
    End Sub


    Private Sub TxtAGasto_LostFocus(index As Integer)
    If IsNumeric(TxtAGasto(4)) Then TxtAGasto(4) = Format(TxtAGasto(4), "#,##0.00")
    If index = 4 Then Call RtnGrid
    End Sub

    Private Sub Inicializar()
    Dim D%
    For D = 0 To 8: TxtAGasto(D) = ""   'blanquea los cuadros de textos
    Next D
    For D = 2 To 6: Check1(D) = vbUnchecked 'inicializa los check
    Next D
    For D = 0 To 1: Combo1(D).ListIndex = -1    'blanquea los combobox
    Next D
    ReDim vDistribucion(0, 0)

    End Sub

    Private Sub grid_clear()
    With GridGastos(0)
        .Rows = 2
        .Row = 1
        .Col = 2
        Set .CellPicture = Nothing
        .Col = 3
        Set .CellPicture = Nothing
    End With
    Call rtnLimpiar_Grid(GridGastos(0))
    End Sub

    '---------------------------------------------------------------------------------------------
    '
    '   Función:    Filtrar
    '
    '   Selecciona un conjunto de los registro según parametros determinados
    '   por el ususario
    '---------------------------------------------------------------------------------------------
    Private Sub rtnFiltrar()
    'variables locales
    Dim strFiltro As String
    Dim fecha As String
    'Aplica filtro por
    If TxtAGasto(9) <> "" Then  'código de gasto
        strFiltro = "CodGasto='" & TxtAGasto(9) & "'"
    End If
    If TxtAGasto(10) <> "" Then 'Período
        If IsDate("01/" & TxtAGasto(10)) Then
            fecha = "01/" & TxtAGasto(10)
            strFiltro = IIf(strFiltro = "", "Periodo=#" & fecha & "#", strFiltro & " AND Periodo=#" & fecha & "#")
        Else
            MsgBox "Introdujo una valor inválido en el campo cargado", vbCritical, App.ProductName
        End If
    End If
    If TxtAGasto(11) <> "" Then 'Monto
        If IsNumeric(TxtAGasto(11)) Then
            strFiltro = IIf(strFiltro = "", "Monto=" & TxtAGasto(11), strFiltro & " AND Monto=" & TxtAGasto(11))
        Else
            MsgBox "Introdujo un valor invalido", vbInformation, App.ProductName
        End If
    End If
    '
    If TxtAGasto(12) <> "" Then 'Fecha
        If IsDate(TxtAGasto(12)) Then
            fecha = Format(TxtAGasto(12), "mm/dd/yy")
            fecha = TxtAGasto(12)
            strFiltro = IIf(strFiltro = "", "Fecha=#" & fecha & "#", strFiltro & " AND Fecha=#" & fecha & "#")
        Else
            MsgBox "Introdujo una valor inválido en el campo cargado", vbCritical, App.ProductName
        End If
    End If
    '
    If TxtAGasto(13) <> "" Then 'Usuario
        strFiltro = IIf(strFiltro = "", "Usuario='" & TxtAGasto(13) & "'", strFiltro & " AND Us" _
        & "uario='" & TxtAGasto(13) & "'")
    End If
    '
    rstNov(Nov).Filter = IIf(strFiltro = "", 0, strFiltro)
    If rstNov(Nov).Filter = 0 Then rstNov(Nov).Requery
    End Sub

    '---------------------------------------------------------------------------------------------
    '   función:    validar
    '
    '   comprueba el estatus de la factura y si esta està registrada, de no ser así
    '   devuelve un mensaje de errir
    '---------------------------------------------------------------------------------------------
    Private Function Validar() As Boolean
    '
    Dim rstCpp As New ADODB.Recordset   'variable local
    
    rstCpp.Open "SELECT * FROM Cpp WHERE Ndoc='" & TxtAGasto(14) & "'", cnnConexion, adOpenKeyset, adLockOptimistic, adCmdText
    If rstCpp.EOF Or rstCpp.BOF Then
        Validar = MsgBox("El Documento Nº " & TxtAGasto(14) & " no esta registrado", _
        vbInformation, App.ProductName)
    Else
        If rstCpp!CodInm <> gcCodInm Then
            Validar = MsgBox("Este documento no esta registrado en las Ctas. por Pagar del  " _
            & "inmueble " & gcCodInm, vbInformation + vbOKOnly, App.ProductName)
        ElseIf rstCpp!Estatus <> "PENDIENTE" Then
            Validar = MsgBox("Imposible efectuar esta operación. Este documento esta " & _
            rstCpp!Estatus, vbInformation, App.ProductName)
        ElseIf rstCpp!Total <> rstNov(Nov)!Monto Then
            Validar = MsgBox("Imposible efectuar esta operación. El monto en Bs. de la novedad no concue" _
            & "rda con la factura", vbInformation, App.ProductName)
        End If
    End If
    rstCpp.Close
    Set rstcopp = Nothing
    End Function

    '---------------------------------------------------------------------------------------------
    '   Rutina: Printer_report
    '
    '   Imprimir el reporte de novedades según las mostrada en la lista
    '---------------------------------------------------------------------------------------------
    Private Sub Printer_Report()
    'Variables locales
    Dim strSql As String
    Dim errLocal As Long, rpReporte As ctlReport
    '
    
    If rstNov(Nov).EOF Or rstNov(Nov).BOF Then
        MsgBox "No existen registros que imprimir...", vbInformation, App.ProductName
        Exit Sub
    End If
    MousePointer = vbHourglass
    'Crea una consulta con la propiedad source y filter del ADODB.Recordset
    strSql = rstNov(Nov).Source
    If Not rstNov(Nov).Filter = 0 Then strSql = strSql & " AND " & rstNov(Nov).Filter
    Call rtnGenerator(mcDatos, strSql, "qdfNovFact")
    'Imprime el reporte
    Set rpReporte = New ctlReport
    With rpReporte
        .Reporte = gcReport + "fact_nov.rpt"
        .OrigenDatos(0) = mcDatos
        .Formulas(0) = "Inmueble='" & gcCodInm & "-" & gcNomInm & "'"
        .Formulas(1) = "User='" & gcUsuario & "'"
        errLocal = .Imprimir
        'registra en la bitacora la impresión del reporte
        Call rtnBitacora("Printer Nov. Fact. Inm.:" & gcCodInm)
        If errLocal <> 0 Then   'si ocurre un error da un mensaje al usuario
                                            ' y lo registra en la bitacora del sistema
            MsgBox "Error al imprimir el reporte." & vbCrLf & Err.Description, vbExclamation, _
            "Error: " & Err
            Call rtnBitacora(Err.Description & "-" & Err)
        End If
    End With
    Set rpReporte = Nothing
    MousePointer = vbDefault
    '
    End Sub

    '-------------------------------------------------------------------------------------------------
    Private Sub RtnGrid() 'RUTINA QUE DISTRIBUYE EL GRID PROP.SELECCIONADOS
    '-------------------------------------------------------------------------------------------------
    Dim CurPorcion As Currency
    Dim ByFila As Byte
    Dim rstProp As ADODB.Recordset
    
    
    If Check1(2).Value = 0 And TxtAGasto(4) <> "" Then
        Call rtnLimpiar_Grid(GridGastos(2))
        
        
        If Option1(1).Value = True Then 'propietarios seleccionados
        
            With GridGastos(2)
                .Rows = LisAGasto(1).ListCount + 1
                CurPorcion = TxtAGasto(4) / LisAGasto(1).ListCount
                For i = 1 To LisAGasto(1).ListCount
                    .TextMatrix(i, 0) = LisAGasto(1).List(i - 1)
                    .TextMatrix(i, 1) = LisAGasto(1).List(i - 1)
                    If UBound(vDistribucion, 1) > 1 Then
                        If vDistribucion(0, 0) <> CCur(TxtAGasto(4)) Then
                            .TextMatrix(i, 2) = Format(CurPorcion, "#,##0.00")
                        Else
                            .TextMatrix(i, 2) = Format(vDistribucion(i, 1), "#,##0.00")
                        End If
                    Else
                        .TextMatrix(i, 2) = Format(CurPorcion, "#,##0.00")
                    End If
                Next
'                If UBound(vDistribucion, 1) > 1 Then
'                    GridGastos(2).Rows = UBound(vDistribucion, 1) + 1
'                    For I = 1 To UBound(vDistribucion, 1)
'                        GridGastos(2).TextMatrix(I, 0) = vDistribucion(I, 0)
'                        GridGastos(2).TextMatrix(I, 1) = vDistribucion(I, 0) 'Right(LisAGasto(1).list(I - 1), _
'                        Len(LisAGasto(1).list(I)))
'                        GridGastos(2).TextMatrix(I, 2) = vDistribucion(I, 1)
'                    Next
'                Else
'
'                    CurPorcion = CLng(TxtAGasto(4) / LisAGasto(1).ListCount)
'                    For I = 0 To (LisAGasto(1).ListCount - 1)
'                        'Llena el Grid con los datos de la lista
'                        ByFila = I + 1
'                        If ByFila >= .Rows Then .AddItem ("")
'                        .TextMatrix(ByFila, 0) = RTrim(Left(LisAGasto(1).list(I), 6))
'                        .TextMatrix(ByFila, 1) = Right(LisAGasto(1).list(I), _
'                        Len(LisAGasto(1).list(I)) - 0)
'                        .TextMatrix(I + 1, 2) = Format(CurPorcion, "#,##0.00")
'                    Next
'                End If
            End With
            
        Else
            'todos los propietarios
            Set rstProp = New ADODB.Recordset
            rstProp.CursorLocation = adUseClient
            rstProp.Open "Propietarios", CnnInm, adOpenKeyset, adLockReadOnly, admcdtable
            rstProp.Sort = "Codigo"
            rstProp.Filter = "Codigo <> 'U" & gcCodInm & "'"
            
            If Not (rstProp.EOF And rstProp.BOF) Then
                rstProp.MoveFirst
                'partes iguales
                If Check1(3).Value = vbUnchecked Then
                    CurPorcion = TxtAGasto(4) / (rstProp.RecordCount)
                Else
                    CurPorcion = TxtAGasto(4) * rstProp("Alicuota")
                End If
                Do
                    
                    With GridGastos(2)
                        ByFila = ByFila + 1
                        If Check1(3).Value = vbChecked Then CurPorcion = _
                        TxtAGasto(4) * rstProp!Alicuota / 100
                        If ByFila >= .Rows Then .AddItem ("")
                        .TextMatrix(ByFila, 0) = rstProp!Codigo
                        .TextMatrix(ByFila, 1) = rstProp!Nombre
                        If UBound(vDistribucion, 1) > 1 Then
                            If vDistribucion(0, 0) <> CCur(TxtAGasto(4)) Then
                                .TextMatrix(ByFila, 2) = Format(CurPorcion, "#,##0.00 ")
                            Else
                                .TextMatrix(ByFila, 2) = Format(vDistribucion(ByFila, 1), "#,##0.00")
                            End If
                        Else
                            .TextMatrix(ByFila, 2) = Format(CurPorcion, "#,##0.00 ")
                        End If
                        rstProp.MoveNext
                    End With
                Loop Until rstProp.EOF
                
            End If
            rstProp.Close
            Set rstProp = Nothing
                    
        End If
        '
    End If
    
    
    'configura la presentacion del grid en pantalla
    With GridGastos(2)
        'llena el vector vdistribucion
        If IsNumeric(.TextMatrix(1, 2)) Then
        ReDim vDistribucion(.Rows - 1, 1)
        vDistribucion(0, 0) = CCur(IIf(TxtAGasto(4) = "", 0, TxtAGasto(4)))
        For i = 1 To .Rows - 1
            vDistribucion(i, 0) = .TextMatrix(i, 0)
            vDistribucion(i, 1) = CCur(.TextMatrix(i, 2))
        Next
        End If
        'AltoFila = .RowHeight(1)
        .Height = (.Rows + 1) * .RowHeight(1)
        If Not (ActiveControl = LisAGasto(0) Or ActiveControl = LisAGasto(1)) Then
            fraAgasto(4).Height = .Height + 100 + .Top
        End If
        If fraAgasto(4).Height > SSTab1.Height - fraAgasto(4).Top - 100 Then
            fraAgasto(4).Height = SSTab1.Height - fraAgasto(4).Top - 100
            .Height = fraAgasto(4).Height - .Top - 100
            If .RowHeight(1) * (.Rows + 1) > .Height Then
                fraAgasto(4).Left = TxtAGasto(2).Left - 200
                fraAgasto(4).Width = TxtAGasto(4).Left + TxtAGasto(4).Width - TxtAGasto(2).Left + 300
                
                .Width = fraAgasto(4).Width - 200
'                .Left = 200
            End If
            
        End If
    End With
    End Sub



