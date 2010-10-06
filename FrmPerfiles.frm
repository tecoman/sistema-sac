VERSION 5.00
Begin VB.Form FrmPerfiles 
   Caption         =   "Perfíl y Accesos"
   ClientHeight    =   3195
   ClientLeft      =   1260
   ClientTop       =   1560
   ClientWidth     =   4740
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "FrmPerfiles.frx":0000
   LinkTopic       =   "form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   3195
   ScaleWidth      =   4740
   WindowState     =   2  'Maximized
   Begin VB.PictureBox pctPerfil 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2280
      Index           =   69
      Left            =   5745
      ScaleHeight     =   2280
      ScaleWidth      =   3225
      TabIndex        =   178
      Tag             =   "2"
      Top             =   3030
      Visible         =   0   'False
      Width           =   3225
      Begin VB.PictureBox pctPerfil 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   2250
         Index           =   70
         Left            =   15
         ScaleHeight     =   2250
         ScaleWidth      =   3195
         TabIndex        =   179
         Top             =   15
         Width           =   3195
         Begin VB.CheckBox Check1 
            Caption         =   "Reporte de Morosos"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   315
            Index           =   134
            Left            =   15
            TabIndex        =   186
            Tag             =   "00AC40421(6)"
            Top             =   1905
            Width           =   3150
         End
         Begin VB.CheckBox Check1 
            Caption         =   "Consulta de Honorarios"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   315
            Index           =   133
            Left            =   15
            TabIndex        =   185
            Tag             =   "00AC40421(5)"
            Top             =   1590
            Width           =   3150
         End
         Begin VB.CheckBox Check1 
            Caption         =   "Reporte de Honorarios (Legal)"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   315
            Index           =   132
            Left            =   15
            TabIndex        =   184
            Tag             =   "00AC40421(4)"
            Top             =   1275
            Width           =   3150
         End
         Begin VB.CheckBox Check1 
            Caption         =   "Reporte de Atrasos"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   315
            Index           =   131
            Left            =   15
            TabIndex        =   183
            Tag             =   "00AC40421(3)"
            Top             =   960
            Width           =   3150
         End
         Begin VB.CheckBox Check1 
            Caption         =   "Estado de Cuenta por Inmueble"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   315
            Index           =   130
            Left            =   15
            TabIndex        =   182
            Tag             =   "00AC40421(2)"
            Top             =   645
            Width           =   3150
         End
         Begin VB.CheckBox Check1 
            Caption         =   "Relación Recibos por Enviar"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   315
            Index           =   129
            Left            =   30
            TabIndex        =   181
            Tag             =   "00AC40421(1)"
            Top             =   330
            Width           =   3150
         End
         Begin VB.CheckBox Check1 
            Caption         =   "Relación de CxC al Cobrador"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   315
            Index           =   128
            Left            =   30
            TabIndex        =   180
            Tag             =   "00AC40421(0)"
            Top             =   30
            Width           =   3150
         End
         Begin VB.Line lnePerfil 
            BorderColor     =   &H00808080&
            Index           =   41
            X1              =   3180
            X2              =   3180
            Y1              =   0
            Y2              =   2550
         End
         Begin VB.Line lnePerfil 
            BorderColor     =   &H00808080&
            Index           =   40
            X1              =   15
            X2              =   3195
            Y1              =   2235
            Y2              =   2235
         End
      End
      Begin VB.PictureBox pctPerfil 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   2265
         Index           =   71
         Left            =   15
         ScaleHeight     =   2265
         ScaleWidth      =   3210
         TabIndex        =   187
         Top             =   15
         Width           =   3210
      End
   End
   Begin VB.PictureBox pctPerfil 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   690
      Index           =   60
      Left            =   8955
      ScaleHeight     =   690
      ScaleWidth      =   2070
      TabIndex        =   157
      Tag             =   "1"
      Top             =   3030
      Visible         =   0   'False
      Width           =   2070
      Begin VB.PictureBox pctPerfil 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   660
         Index           =   61
         Left            =   15
         ScaleHeight     =   660
         ScaleWidth      =   2040
         TabIndex        =   158
         Top             =   15
         Width           =   2040
         Begin VB.CheckBox Check1 
            Caption         =   "Estadísticos       »"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   315
            Index           =   113
            Left            =   15
            TabIndex        =   160
            Tag             =   "72AC4042(1)"
            Top             =   330
            Width           =   2010
         End
         Begin VB.CheckBox Check1 
            Caption         =   "Administrativos   »"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   315
            Index           =   112
            Left            =   15
            TabIndex        =   159
            Tag             =   "69AC4042(0)"
            Top             =   30
            Width           =   2010
         End
         Begin VB.Line lnePerfil 
            BorderColor     =   &H00808080&
            Index           =   39
            X1              =   2025
            X2              =   2025
            Y1              =   15
            Y2              =   660
         End
         Begin VB.Line lnePerfil 
            BorderColor     =   &H00808080&
            Index           =   38
            X1              =   15
            X2              =   2040
            Y1              =   645
            Y2              =   645
         End
      End
      Begin VB.PictureBox pctPerfil 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   700
         Index           =   62
         Left            =   15
         ScaleHeight     =   705
         ScaleWidth      =   2055
         TabIndex        =   161
         Top             =   15
         Width           =   2055
      End
   End
   Begin VB.PictureBox pctPerfil 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1320
      Index           =   63
      Left            =   8955
      ScaleHeight     =   1320
      ScaleWidth      =   2190
      TabIndex        =   162
      Tag             =   "1"
      Top             =   2715
      Visible         =   0   'False
      Width           =   2190
      Begin VB.PictureBox pctPerfil 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1290
         Index           =   64
         Left            =   15
         ScaleHeight     =   1290
         ScaleWidth      =   2160
         TabIndex        =   163
         Top             =   15
         Width           =   2160
         Begin VB.CheckBox Check1 
            Caption         =   "Gestión de Cobro"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   315
            Index           =   120
            Left            =   15
            TabIndex        =   251
            Tag             =   "00AC4041(0)"
            Top             =   0
            Width           =   2115
         End
         Begin VB.CheckBox Check1 
            Caption         =   "Emisión de Giros"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   315
            Index           =   116
            Left            =   15
            TabIndex        =   166
            Tag             =   "00AC4041(3)"
            Top             =   945
            Width           =   2115
         End
         Begin VB.CheckBox Check1 
            Caption         =   "Convenio de Pago"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   315
            Index           =   115
            Left            =   15
            TabIndex        =   165
            Tag             =   "00AC4041(2)"
            Top             =   630
            Width           =   2115
         End
         Begin VB.CheckBox Check1 
            Caption         =   "Registro de Pagos"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   315
            Index           =   114
            Left            =   15
            TabIndex        =   164
            Tag             =   "00AC4041(1)"
            Top             =   315
            Width           =   2115
         End
         Begin VB.Line lnePerfil 
            BorderColor     =   &H00808080&
            Index           =   37
            X1              =   2145
            X2              =   2145
            Y1              =   0
            Y2              =   1320
         End
         Begin VB.Line lnePerfil 
            BorderColor     =   &H00808080&
            Index           =   42
            X1              =   0
            X2              =   2160
            Y1              =   1275
            Y2              =   1275
         End
      End
      Begin VB.PictureBox pctPerfil 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1305
         Index           =   65
         Left            =   15
         ScaleHeight     =   1305
         ScaleWidth      =   2175
         TabIndex        =   167
         Top             =   15
         Width           =   2175
      End
   End
   Begin VB.PictureBox pctPerfil 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2295
      Index           =   66
      Left            =   5880
      ScaleHeight     =   2295
      ScaleWidth      =   3225
      TabIndex        =   168
      Tag             =   "0"
      Top             =   1095
      Visible         =   0   'False
      Width           =   3225
      Begin VB.PictureBox pctPerfil 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   2265
         Index           =   67
         Left            =   15
         ScaleHeight     =   2265
         ScaleWidth      =   3195
         TabIndex        =   169
         Top             =   15
         Width           =   3195
         Begin VB.CheckBox Check1 
            Caption         =   "Estado de Cuenta"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   315
            Index           =   126
            Left            =   15
            TabIndex        =   176
            Tag             =   "00AC404(0)"
            Top             =   30
            Width           =   3150
         End
         Begin VB.CheckBox Check1 
            Caption         =   "Avisos de Cobro"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   315
            Index           =   125
            Left            =   15
            TabIndex        =   175
            Tag             =   "00AC404(1)"
            Top             =   345
            Width           =   3150
         End
         Begin VB.CheckBox Check1 
            Caption         =   "Devolución de Cheque"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   315
            Index           =   123
            Left            =   15
            TabIndex        =   174
            Tag             =   "00AC404(2)"
            Top             =   660
            Width           =   3150
         End
         Begin VB.CheckBox Check1 
            Caption         =   "Consulta Administrativa"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   315
            Index           =   122
            Left            =   15
            TabIndex        =   173
            Tag             =   "00AC404(3)"
            Top             =   975
            Width           =   3150
         End
         Begin VB.CheckBox Check1 
            Caption         =   "Asignar Cobrador al Cliente"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   315
            Index           =   121
            Left            =   15
            TabIndex        =   172
            Tag             =   "00AC404(4)"
            Top             =   1290
            Width           =   3150
         End
         Begin VB.CheckBox Check1 
            Caption         =   "Departamento Jurídico           »"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   315
            Index           =   119
            Left            =   15
            TabIndex        =   171
            Tag             =   "63AC404(5)"
            Top             =   1605
            Width           =   3150
         End
         Begin VB.CheckBox Check1 
            Caption         =   "[CxC] Consultas y Reportes     »"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   315
            Index           =   117
            Left            =   15
            TabIndex        =   170
            Tag             =   "60AC404(6)"
            Top             =   1920
            Width           =   3150
         End
         Begin VB.Line lnePerfil 
            BorderColor     =   &H00808080&
            Index           =   36
            X1              =   3180
            X2              =   3180
            Y1              =   0
            Y2              =   3180
         End
         Begin VB.Line lnePerfil 
            BorderColor     =   &H00808080&
            Index           =   35
            X1              =   0
            X2              =   3195
            Y1              =   2250
            Y2              =   2250
         End
      End
      Begin VB.PictureBox pctPerfil 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   2280
         Index           =   68
         Left            =   15
         ScaleHeight     =   2280
         ScaleWidth      =   3210
         TabIndex        =   177
         Top             =   15
         Width           =   3210
      End
   End
   Begin VB.PictureBox pctPerfil 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      DrawStyle       =   2  'Dot
      DrawWidth       =   16847
      ForeColor       =   &H80000008&
      Height          =   690
      Index           =   87
      Left            =   7050
      ScaleHeight     =   690
      ScaleWidth      =   1845
      TabIndex        =   234
      Tag             =   "1"
      Top             =   3060
      Visible         =   0   'False
      Width           =   1845
      Begin VB.PictureBox pctPerfil 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         BorderStyle     =   0  'None
         DrawStyle       =   2  'Dot
         DrawWidth       =   16847
         ForeColor       =   &H80000008&
         Height          =   660
         Index           =   88
         Left            =   15
         ScaleHeight     =   660
         ScaleWidth      =   1815
         TabIndex        =   235
         Top             =   15
         Width           =   1815
         Begin VB.CheckBox Check1 
            BackColor       =   &H80000004&
            Caption         =   "Operativos"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   315
            Index           =   165
            Left            =   15
            TabIndex        =   237
            Tag             =   "00AC70502"
            Top             =   330
            Width           =   1770
         End
         Begin VB.CheckBox Check1 
            BackColor       =   &H80000004&
            Caption         =   "Administrativos"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   315
            Index           =   164
            Left            =   15
            TabIndex        =   236
            Tag             =   "00AC70501"
            Top             =   15
            Width           =   1770
         End
      End
      Begin VB.PictureBox pctPerfil 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         DrawStyle       =   2  'Dot
         DrawWidth       =   16847
         ForeColor       =   &H80000008&
         Height          =   675
         Index           =   89
         Left            =   15
         ScaleHeight     =   675
         ScaleWidth      =   1830
         TabIndex        =   238
         Top             =   15
         Width           =   1830
      End
   End
   Begin VB.PictureBox pctPerfil 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2790
      Index           =   84
      Left            =   8880
      ScaleHeight     =   2790
      ScaleWidth      =   2985
      TabIndex        =   225
      Tag             =   "0"
      Top             =   1095
      Visible         =   0   'False
      Width           =   2985
      Begin VB.PictureBox pctPerfil 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   2760
         Index           =   85
         Left            =   15
         ScaleHeight     =   2760
         ScaleWidth      =   2955
         TabIndex        =   226
         Top             =   15
         Width           =   2955
         Begin VB.CheckBox Check1 
            Caption         =   "Quórum"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   315
            Index           =   110
            Left            =   15
            TabIndex        =   268
            Tag             =   "00AC707(1)"
            Top             =   2350
            Width           =   2925
         End
         Begin VB.CheckBox Check1 
            Caption         =   "Control de Asistencia"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   315
            Index           =   111
            Left            =   15
            TabIndex        =   267
            Tag             =   "00AC708"
            Top             =   1600
            Width           =   2925
         End
         Begin VB.CheckBox Check1 
            Caption         =   "Parámetros del Sistema      »"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   315
            Index           =   163
            Left            =   15
            TabIndex        =   232
            Tag             =   "87AC705"
            Top             =   1935
            Width           =   2925
         End
         Begin VB.CheckBox Check1 
            Caption         =   "Bitácora del Sistema"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   315
            Index           =   162
            Left            =   15
            TabIndex        =   231
            Tag             =   "00AC706"
            Top             =   1275
            Width           =   2925
         End
         Begin VB.CheckBox Check1 
            Caption         =   "Mantenimiento de B.D."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   315
            Index           =   161
            Left            =   15
            TabIndex        =   230
            Tag             =   "00AC704"
            Top             =   960
            Width           =   2925
         End
         Begin VB.CheckBox Check1 
            Caption         =   "Datos de la Empresa"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   315
            Index           =   160
            Left            =   15
            TabIndex        =   229
            Tag             =   "00AC703"
            Top             =   645
            Width           =   2925
         End
         Begin VB.CheckBox Check1 
            Caption         =   "Usuarios y Perfiles"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   315
            Index           =   159
            Left            =   15
            TabIndex        =   228
            Tag             =   "00AC702"
            Top             =   330
            Width           =   2925
         End
         Begin VB.CheckBox Check1 
            Caption         =   "Perfiles de Acceso"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   315
            Index           =   158
            Left            =   15
            TabIndex        =   227
            Tag             =   "00AC701"
            Top             =   15
            Width           =   2925
         End
         Begin VB.Line lnePerfil 
            BorderColor     =   &H00FFFFFF&
            Index           =   61
            X1              =   0
            X2              =   2985
            Y1              =   2280
            Y2              =   2280
         End
         Begin VB.Line lnePerfil 
            BorderColor     =   &H00808080&
            BorderWidth     =   2
            Index           =   31
            X1              =   0
            X2              =   2985
            Y1              =   2280
            Y2              =   2280
         End
         Begin VB.Line Line1 
            BorderColor     =   &H00808080&
            X1              =   2940
            X2              =   2940
            Y1              =   1890
            Y2              =   15
         End
         Begin VB.Line lnePerfil 
            BorderColor     =   &H00808080&
            Index           =   51
            X1              =   15
            X2              =   3000
            Y1              =   2745
            Y2              =   2745
         End
      End
      Begin VB.PictureBox pctPerfil 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   2775
         Index           =   86
         Left            =   15
         ScaleHeight     =   2775
         ScaleWidth      =   2970
         TabIndex        =   233
         Top             =   15
         Width           =   2970
      End
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H80000000&
      Caption         =   "Relación de Deuda por Propietario"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      Index           =   89
      Left            =   0
      TabIndex        =   266
      Tag             =   "00"
      Top             =   0
      Visible         =   0   'False
      Width           =   3390
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H80000001&
      Caption         =   "Relación de Deuda por Propietario"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      Index           =   109
      Left            =   0
      TabIndex        =   265
      Tag             =   "00"
      Top             =   0
      Visible         =   0   'False
      Width           =   3390
   End
   Begin VB.PictureBox pctPerfil 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   690
      Index           =   51
      Left            =   7770
      ScaleHeight     =   690
      ScaleWidth      =   2145
      TabIndex        =   146
      Tag             =   "1"
      Top             =   2850
      Visible         =   0   'False
      Width           =   2145
      Begin VB.PictureBox pctPerfil 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   660
         Index           =   52
         Left            =   15
         ScaleHeight     =   660
         ScaleWidth      =   2115
         TabIndex        =   147
         Top             =   15
         Width           =   2115
         Begin VB.CheckBox Check1 
            Caption         =   "Cierre de Caja"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   315
            Index           =   105
            Left            =   15
            TabIndex        =   149
            Tag             =   "00AC4001(1)"
            Top             =   330
            Width           =   2085
         End
         Begin VB.CheckBox Check1 
            Caption         =   "Deducciones"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   315
            Index           =   104
            Left            =   15
            TabIndex        =   148
            Tag             =   "00AC4001(0)"
            Top             =   30
            Width           =   2085
         End
         Begin VB.Line lnePerfil 
            BorderColor     =   &H00808080&
            Index           =   34
            X1              =   2100
            X2              =   2100
            Y1              =   0
            Y2              =   660
         End
         Begin VB.Line lnePerfil 
            BorderColor     =   &H00808080&
            Index           =   33
            X1              =   15
            X2              =   2115
            Y1              =   645
            Y2              =   645
         End
      End
      Begin VB.PictureBox pctPerfil 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   675
         Index           =   53
         Left            =   15
         ScaleHeight     =   675
         ScaleWidth      =   2130
         TabIndex        =   150
         Top             =   15
         Width           =   2130
      End
   End
   Begin VB.PictureBox pctPerfil 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2745
      Index           =   48
      Left            =   5205
      ScaleHeight     =   2745
      ScaleWidth      =   2715
      TabIndex        =   135
      Tag             =   "0"
      Top             =   1095
      Visible         =   0   'False
      Width           =   2715
      Begin VB.PictureBox pctPerfil 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   2715
         Index           =   49
         Left            =   15
         ScaleHeight     =   2715
         ScaleWidth      =   2685
         TabIndex        =   136
         Top             =   15
         Width           =   2685
         Begin VB.CheckBox Check1 
            Caption         =   "Abrir Caja"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   315
            Index           =   103
            Left            =   15
            TabIndex        =   144
            Tag             =   "00AC400(0)"
            Top             =   15
            Width           =   2655
         End
         Begin VB.CheckBox Check1 
            Caption         =   "Cobranza por Caja"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   315
            Index           =   102
            Left            =   15
            TabIndex        =   143
            Tag             =   "00AC400(1)"
            Top             =   330
            Width           =   2655
         End
         Begin VB.CheckBox Check1 
            Caption         =   "Cuadre de Caja"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   315
            Index           =   101
            Left            =   15
            TabIndex        =   142
            Tag             =   "00AC400(3)"
            Top             =   735
            Width           =   2655
         End
         Begin VB.CheckBox Check1 
            Caption         =   "Portadas de Caja"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   315
            Index           =   100
            Left            =   15
            TabIndex        =   141
            Tag             =   "00AC400(4)"
            Top             =   1050
            Width           =   2655
         End
         Begin VB.CheckBox Check1 
            Caption         =   "[Caja] Consultas y Rep. »"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   315
            Index           =   99
            Left            =   15
            TabIndex        =   140
            Tag             =   "54AC400(5)"
            Top             =   1365
            Width           =   2655
         End
         Begin VB.CheckBox Check1 
            Caption         =   "Autorizar...                   »"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   315
            Index           =   98
            Left            =   15
            TabIndex        =   139
            Tag             =   "51AC400(7)"
            Top             =   1770
            Width           =   2655
         End
         Begin VB.CheckBox Check1 
            Caption         =   "Cerrar Caja"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   315
            Index           =   97
            Left            =   15
            TabIndex        =   138
            Tag             =   "00AC400(8)"
            Top             =   2085
            Width           =   2655
         End
         Begin VB.CheckBox Check1 
            Caption         =   "Aplicar Abonos..."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   315
            Index           =   96
            Left            =   15
            TabIndex        =   137
            Tag             =   "00AC400(9)"
            Top             =   2385
            Width           =   2655
         End
         Begin VB.Line lnePerfil 
            BorderColor     =   &H00808080&
            Index           =   4
            X1              =   30
            X2              =   2650
            Y1              =   1725
            Y2              =   1725
         End
         Begin VB.Line lnePerfil 
            BorderColor     =   &H00FFFFFF&
            BorderWidth     =   2
            Index           =   7
            X1              =   30
            X2              =   2650
            Y1              =   1740
            Y2              =   1740
         End
         Begin VB.Line lnePerfil 
            BorderColor     =   &H00808080&
            Index           =   30
            X1              =   0
            X2              =   2685
            Y1              =   2700
            Y2              =   2700
         End
         Begin VB.Line lnePerfil 
            BorderColor     =   &H00808080&
            Index           =   29
            X1              =   2670
            X2              =   2670
            Y1              =   0
            Y2              =   2700
         End
         Begin VB.Line lnePerfil 
            BorderColor     =   &H00808080&
            Index           =   5
            X1              =   30
            X2              =   2650
            Y1              =   675
            Y2              =   675
         End
         Begin VB.Line lnePerfil 
            BorderColor     =   &H00FFFFFF&
            BorderWidth     =   2
            Index           =   6
            X1              =   30
            X2              =   2650
            Y1              =   690
            Y2              =   690
         End
      End
      Begin VB.PictureBox pctPerfil 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   2730
         Index           =   50
         Left            =   15
         ScaleHeight     =   2730
         ScaleWidth      =   2700
         TabIndex        =   145
         Top             =   15
         Width           =   2700
      End
   End
   Begin VB.PictureBox pctPerfil 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2250
      Index           =   6
      Left            =   960
      ScaleHeight     =   2250
      ScaleWidth      =   2865
      TabIndex        =   29
      Tag             =   "0"
      Top             =   1095
      Visible         =   0   'False
      Width           =   2865
      Begin VB.PictureBox pctPerfil 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   2220
         Index           =   7
         Left            =   15
         ScaleHeight     =   2220
         ScaleWidth      =   2835
         TabIndex        =   30
         Top             =   15
         Width           =   2835
         Begin VB.CheckBox Check1 
            Caption         =   "Aviso de Cobro [Cpo. Msg]"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   118
            Left            =   15
            TabIndex        =   264
            Tag             =   "00AC107"
            Top             =   1245
            Width           =   2790
         End
         Begin VB.CheckBox Check1 
            Caption         =   "Ficha de Inmueble"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   23
            Left            =   15
            TabIndex        =   36
            Tag             =   "00AC101"
            Top             =   75
            Width           =   2790
         End
         Begin VB.CheckBox Check1 
            Caption         =   "Ficha del Propietario"
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
            Index           =   22
            Left            =   15
            TabIndex        =   35
            Tag             =   "00AC102"
            Top             =   330
            Width           =   2790
         End
         Begin VB.CheckBox Check1 
            Caption         =   "Catálogo de Gastos"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   21
            Left            =   15
            TabIndex        =   34
            Tag             =   "00AC103"
            Top             =   615
            Width           =   2790
         End
         Begin VB.CheckBox Check1 
            Caption         =   "Catálogo de Fondos"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   20
            Left            =   15
            TabIndex        =   33
            Tag             =   "00AC104"
            Top             =   930
            Width           =   2790
         End
         Begin VB.CheckBox Check1 
            Caption         =   "Editar Cartas de Morosidad »"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   19
            Left            =   15
            TabIndex        =   32
            Tag             =   "09AC105"
            Top             =   1560
            Width           =   2790
         End
         Begin VB.CheckBox Check1 
            Caption         =   "Consultas y Reportes         »"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   18
            Left            =   15
            TabIndex        =   31
            Tag             =   "12AC106"
            Top             =   1875
            Width           =   2790
         End
         Begin VB.Line lnePerfil 
            BorderColor     =   &H00808080&
            Index           =   11
            X1              =   2820
            X2              =   2820
            Y1              =   45
            Y2              =   2430
         End
         Begin VB.Line lnePerfil 
            BorderColor     =   &H00808080&
            Index           =   12
            X1              =   0
            X2              =   2805
            Y1              =   2205
            Y2              =   2205
         End
      End
      Begin VB.PictureBox pctPerfil 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   2235
         Index           =   8
         Left            =   15
         ScaleHeight     =   2235
         ScaleWidth      =   2850
         TabIndex        =   37
         Top             =   15
         Width           =   2850
      End
   End
   Begin VB.PictureBox pctPerfil 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2910
      Index           =   36
      Left            =   6660
      ScaleHeight     =   2910
      ScaleWidth      =   3405
      TabIndex        =   97
      Tag             =   "2"
      Top             =   3765
      Visible         =   0   'False
      Width           =   3405
      Begin VB.PictureBox pctPerfil 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   2880
         Index           =   37
         Left            =   15
         ScaleHeight     =   2880
         ScaleWidth      =   3360
         TabIndex        =   98
         Top             =   15
         Width           =   3360
         Begin VB.CheckBox Check1 
            Caption         =   "Reporte de Transacciones"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   315
            Index           =   73
            Left            =   15
            TabIndex        =   107
            Tag             =   "00AC2100202"
            Top             =   15
            Width           =   3330
         End
         Begin VB.CheckBox Check1 
            Caption         =   "Estado de Cuenta por Proveedor"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   315
            Index           =   72
            Left            =   15
            TabIndex        =   106
            Tag             =   "00AC2100203"
            Top             =   330
            Width           =   3330
         End
         Begin VB.CheckBox Check1 
            Caption         =   "Estado de Cuenta por Concepto"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   315
            Index           =   71
            Left            =   15
            TabIndex        =   105
            Tag             =   "00AC2100204"
            Top             =   645
            Width           =   3330
         End
         Begin VB.CheckBox Check1 
            Caption         =   "Cuenas por Pagar"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   315
            Index           =   70
            Left            =   15
            TabIndex        =   104
            Tag             =   "00AC2100205"
            Top             =   960
            Width           =   3330
         End
         Begin VB.CheckBox Check1 
            Caption         =   "Libro de Compras"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   315
            Index           =   69
            Left            =   15
            TabIndex        =   103
            Tag             =   "00AC2100206"
            Top             =   1275
            Width           =   3330
         End
         Begin VB.CheckBox Check1 
            Caption         =   "Resumen Anual de Gastos"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   315
            Index           =   68
            Left            =   15
            TabIndex        =   102
            Tag             =   "00AC2100207"
            Top             =   1590
            Width           =   3330
         End
         Begin VB.CheckBox Check1 
            Caption         =   "Análisis de Vencimiento"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   315
            Index           =   67
            Left            =   15
            TabIndex        =   101
            Tag             =   "00AC2100208"
            Top             =   1905
            Width           =   3330
         End
         Begin VB.CheckBox Check1 
            Caption         =   "Relaciones de I.S.L.R."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   315
            Index           =   66
            Left            =   15
            TabIndex        =   100
            Tag             =   "00AC2100209"
            Top             =   2220
            Width           =   3330
         End
         Begin VB.CheckBox Check1 
            Caption         =   "Estadísticas de Compras"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   315
            Index           =   65
            Left            =   15
            TabIndex        =   99
            Tag             =   "00AC2100210"
            Top             =   2535
            Width           =   3330
         End
         Begin VB.Line lnePerfil 
            BorderColor     =   &H00808080&
            Index           =   53
            X1              =   3345
            X2              =   3345
            Y1              =   3150
            Y2              =   0
         End
         Begin VB.Line lnePerfil 
            BorderColor     =   &H00808080&
            Index           =   52
            X1              =   15
            X2              =   3360
            Y1              =   2865
            Y2              =   2865
         End
      End
      Begin VB.PictureBox pctPerfil 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   2895
         Index           =   38
         Left            =   15
         ScaleHeight     =   2895
         ScaleWidth      =   3390
         TabIndex        =   108
         Top             =   15
         Width           =   3390
      End
   End
   Begin VB.PictureBox pctPerfil 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1005
      Index           =   33
      Left            =   4320
      ScaleHeight     =   1005
      ScaleWidth      =   2385
      TabIndex        =   91
      Tag             =   "3"
      Top             =   4080
      Visible         =   0   'False
      Width           =   2385
      Begin VB.PictureBox pctPerfil 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   975
         Index           =   34
         Left            =   15
         ScaleHeight     =   975
         ScaleWidth      =   2355
         TabIndex        =   92
         Top             =   15
         Width           =   2355
         Begin VB.CheckBox Check1 
            Caption         =   "Proveedores"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   315
            Index           =   64
            Left            =   15
            TabIndex        =   95
            Tag             =   "00AC210010401"
            Top             =   0
            Width           =   2325
         End
         Begin VB.CheckBox Check1 
            Caption         =   "Propietarios"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   315
            Index           =   63
            Left            =   15
            TabIndex        =   94
            Tag             =   "00AC210010402"
            Top             =   315
            Width           =   2325
         End
         Begin VB.CheckBox Check1 
            Caption         =   "Junta de Condominio"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   315
            Index           =   62
            Left            =   15
            TabIndex        =   93
            Tag             =   "00AC210010403"
            Top             =   630
            Width           =   2325
         End
         Begin VB.Line lnePerfil 
            BorderColor     =   &H00808080&
            Index           =   55
            X1              =   2340
            X2              =   2340
            Y1              =   1000
            Y2              =   0
         End
         Begin VB.Line lnePerfil 
            BorderColor     =   &H00808080&
            Index           =   54
            X1              =   15
            X2              =   2355
            Y1              =   960
            Y2              =   960
         End
      End
      Begin VB.PictureBox pctPerfil 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   990
         Index           =   35
         Left            =   15
         ScaleHeight     =   990
         ScaleWidth      =   2370
         TabIndex        =   96
         Top             =   15
         Width           =   2370
      End
   End
   Begin VB.PictureBox pctPerfil 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1965
      Index           =   30
      Left            =   6660
      ScaleHeight     =   1965
      ScaleWidth      =   3435
      TabIndex        =   82
      Tag             =   "2"
      Top             =   3375
      Visible         =   0   'False
      Width           =   3435
      Begin VB.PictureBox pctPerfil 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1935
         Index           =   31
         Left            =   15
         ScaleHeight     =   1935
         ScaleWidth      =   3390
         TabIndex        =   83
         Top             =   15
         Width           =   3390
         Begin VB.CheckBox Check1 
            Caption         =   "Lista de Proveedores"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   315
            Index           =   61
            Left            =   15
            TabIndex        =   89
            Tag             =   "00AC2100101"
            Top             =   15
            Width           =   3360
         End
         Begin VB.CheckBox Check1 
            Caption         =   "Remesas Registradas"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   315
            Index           =   60
            Left            =   15
            TabIndex        =   88
            Tag             =   "00AC2100102"
            Top             =   330
            Width           =   3360
         End
         Begin VB.CheckBox Check1 
            Caption         =   "Agenda Telefónica                     »"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   315
            Index           =   58
            Left            =   15
            TabIndex        =   87
            Tag             =   "33AC2100104"
            Top             =   645
            Width           =   3360
         End
         Begin VB.CheckBox Check1 
            Caption         =   "Relación de Facturas Recibidas"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   315
            Index           =   57
            Left            =   15
            TabIndex        =   86
            Tag             =   "00AC2100105"
            Top             =   960
            Width           =   3360
         End
         Begin VB.CheckBox Check1 
            Caption         =   "Relación de Cheques"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   315
            Index           =   56
            Left            =   15
            TabIndex        =   85
            Tag             =   "00AC2100106"
            Top             =   1275
            Width           =   3360
         End
         Begin VB.CheckBox Check1 
            Caption         =   "Cxp Facturación"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   315
            Index           =   55
            Left            =   15
            TabIndex        =   84
            Tag             =   "00AC2100107"
            Top             =   1590
            Width           =   3360
         End
         Begin VB.Line lnePerfil 
            BorderColor     =   &H00808080&
            Index           =   58
            X1              =   0
            X2              =   3390
            Y1              =   1920
            Y2              =   1920
         End
      End
      Begin VB.PictureBox pctPerfil 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1950
         Index           =   32
         Left            =   15
         ScaleHeight     =   1950
         ScaleWidth      =   3420
         TabIndex        =   90
         Top             =   15
         Width           =   3420
      End
   End
   Begin VB.PictureBox pctPerfil 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   690
      Index           =   27
      Left            =   4710
      ScaleHeight     =   690
      ScaleWidth      =   2220
      TabIndex        =   77
      Tag             =   "1"
      Top             =   3375
      Visible         =   0   'False
      Width           =   2220
      Begin VB.PictureBox pctPerfil 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   660
         Index           =   28
         Left            =   15
         ScaleHeight     =   660
         ScaleWidth      =   2190
         TabIndex        =   78
         Top             =   15
         Width           =   2190
         Begin VB.CheckBox Check1 
            Caption         =   "Operativos         »"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   315
            Index           =   54
            Left            =   15
            TabIndex        =   80
            Tag             =   "36AC21002"
            Top             =   330
            Width           =   2160
         End
         Begin VB.CheckBox Check1 
            Caption         =   "Administrativos   »"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   315
            Index           =   53
            Left            =   15
            TabIndex        =   79
            Tag             =   "30AC21001"
            Top             =   15
            Width           =   2160
         End
      End
      Begin VB.PictureBox pctPerfil 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   675
         Index           =   29
         Left            =   15
         ScaleHeight     =   675
         ScaleWidth      =   2205
         TabIndex        =   81
         Top             =   15
         Width           =   2205
      End
   End
   Begin VB.PictureBox pctPerfil 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2610
      Index           =   18
      Left            =   2205
      ScaleHeight     =   2610
      ScaleWidth      =   2715
      TabIndex        =   63
      Tag             =   "0"
      Top             =   1095
      Visible         =   0   'False
      Width           =   2715
      Begin VB.PictureBox pctPerfil 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   2580
         Index           =   19
         Left            =   15
         ScaleHeight     =   2580
         ScaleWidth      =   2685
         TabIndex        =   64
         Top             =   15
         Width           =   2685
         Begin VB.CheckBox Check1 
            Caption         =   "[CxP] Consultas y Rep. »"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   315
            Index           =   38
            Left            =   15
            TabIndex        =   250
            Tag             =   "27AC210"
            Top             =   2235
            Width           =   2655
         End
         Begin VB.CheckBox Check1 
            Caption         =   "Emisión Orden de Compra"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   315
            Index           =   39
            Left            =   15
            TabIndex        =   249
            Tag             =   "00AC209"
            Top             =   1920
            Width           =   2655
         End
         Begin VB.CheckBox Check1 
            Caption         =   "Agenda Telefónica       »"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   315
            Index           =   41
            Left            =   15
            TabIndex        =   248
            Tag             =   "24AC207"
            Top             =   1605
            Width           =   2655
         End
         Begin VB.CheckBox Check1 
            Caption         =   "Cronograma de Pagos   »"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   315
            Index           =   42
            Left            =   15
            TabIndex        =   247
            Tag             =   "21AC206"
            Top             =   1290
            Width           =   2655
         End
         Begin VB.CheckBox Check1 
            Caption         =   "Asignación de Pagos"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   315
            Index           =   44
            Left            =   15
            TabIndex        =   246
            Tag             =   "00AC204"
            Top             =   975
            Width           =   2655
         End
         Begin VB.CheckBox Check1 
            Caption         =   "Registrar Remesa"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   315
            Index           =   45
            Left            =   15
            TabIndex        =   245
            Tag             =   "00AC203"
            Top             =   660
            Width           =   2655
         End
         Begin VB.CheckBox Check1 
            Caption         =   "Recepción de Facturas"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   315
            Index           =   46
            Left            =   15
            TabIndex        =   244
            Tag             =   "00AC202"
            Top             =   345
            Width           =   2655
         End
         Begin VB.CheckBox Check1 
            Caption         =   "Ficha del Proveedor"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   315
            Index           =   47
            Left            =   15
            TabIndex        =   243
            Tag             =   "00AC201"
            Top             =   30
            Width           =   2655
         End
         Begin VB.Line lnePerfil 
            BorderColor     =   &H00808080&
            Index           =   18
            X1              =   0
            X2              =   2685
            Y1              =   2565
            Y2              =   2565
         End
         Begin VB.Line lnePerfil 
            BorderColor     =   &H00808080&
            Index           =   17
            X1              =   2670
            X2              =   2670
            Y1              =   0
            Y2              =   3180
         End
      End
      Begin VB.PictureBox pctPerfil 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   2595
         Index           =   20
         Left            =   15
         ScaleHeight     =   2595
         ScaleWidth      =   2700
         TabIndex        =   65
         Top             =   15
         Width           =   2700
      End
   End
   Begin VB.PictureBox pctPerfil 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2400
      Index           =   45
      Left            =   4305
      ScaleHeight     =   2400
      ScaleWidth      =   2700
      TabIndex        =   128
      Tag             =   "2"
      Top             =   3315
      Visible         =   0   'False
      Width           =   2700
      Begin VB.PictureBox pctPerfil 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   2370
         Index           =   46
         Left            =   15
         ScaleHeight     =   2370
         ScaleWidth      =   2670
         TabIndex        =   129
         Top             =   15
         Width           =   2670
         Begin VB.CheckBox Check1 
            BackColor       =   &H80000004&
            Caption         =   "Control de Facturación"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   315
            Index           =   26
            Left            =   15
            TabIndex        =   259
            Tag             =   "00AC3100107"
            Top             =   1680
            Width           =   2625
         End
         Begin VB.CheckBox Check1 
            BackColor       =   &H80000004&
            Caption         =   "Paquete Completo"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   315
            Index           =   167
            Left            =   15
            TabIndex        =   258
            Tag             =   "00AC3100106"
            Top             =   2025
            Width           =   2625
         End
         Begin VB.CheckBox Check1 
            BackColor       =   &H80000004&
            Caption         =   "Gastos No Comunes"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   315
            Index           =   86
            Left            =   15
            TabIndex        =   257
            Tag             =   "00AC3100105"
            Top             =   1335
            Width           =   2625
         End
         Begin VB.CheckBox Check1 
            BackColor       =   &H80000004&
            Caption         =   "Aviso de Cobro"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   315
            Index           =   92
            Left            =   15
            TabIndex        =   130
            Tag             =   "00AC3100101"
            Top             =   15
            Width           =   2625
         End
         Begin VB.CheckBox Check1 
            BackColor       =   &H80000004&
            Caption         =   "Análisis de Facturación"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   315
            Index           =   95
            Left            =   15
            TabIndex        =   133
            Tag             =   "00AC3100104"
            Top             =   1005
            Width           =   2625
         End
         Begin VB.CheckBox Check1 
            BackColor       =   &H80000004&
            Caption         =   "Reporte de Facturación"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   315
            Index           =   94
            Left            =   15
            TabIndex        =   132
            Tag             =   "00AC3100103"
            Top             =   675
            Width           =   2625
         End
         Begin VB.CheckBox Check1 
            BackColor       =   &H80000004&
            Caption         =   "Pre-Recibo"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   315
            Index           =   93
            Left            =   15
            TabIndex        =   131
            Tag             =   "00AC3100102"
            Top             =   345
            Width           =   2625
         End
         Begin VB.Line lnePerfil 
            BorderColor     =   &H00808080&
            Index           =   28
            X1              =   2655
            X2              =   2655
            Y1              =   0
            Y2              =   1275
         End
         Begin VB.Line lnePerfil 
            BorderColor     =   &H00808080&
            Index           =   27
            X1              =   0
            X2              =   2655
            Y1              =   2355
            Y2              =   2355
         End
      End
      Begin VB.PictureBox pctPerfil 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   2385
         Index           =   47
         Left            =   15
         ScaleHeight     =   2385
         ScaleWidth      =   2685
         TabIndex        =   134
         Top             =   15
         Width           =   2685
      End
   End
   Begin VB.PictureBox pctPerfil 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1740
      Index           =   42
      Left            =   6960
      ScaleHeight     =   1740
      ScaleWidth      =   3735
      TabIndex        =   120
      Tag             =   "1"
      Top             =   3315
      Visible         =   0   'False
      Width           =   3735
      Begin VB.PictureBox pctPerfil 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1710
         Index           =   43
         Left            =   15
         ScaleHeight     =   1710
         ScaleWidth      =   3705
         TabIndex        =   121
         Top             =   15
         Width           =   3705
         Begin VB.CheckBox Check1 
            BackColor       =   &H80000004&
            Caption         =   "Estados Financieros"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   315
            Index           =   91
            Left            =   15
            TabIndex        =   126
            Tag             =   "00AC31007"
            Top             =   1365
            Width           =   3660
         End
         Begin VB.CheckBox Check1 
            BackColor       =   &H80000004&
            Caption         =   "Libro de Ventas (I.V.A.)"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   315
            Index           =   90
            Left            =   15
            TabIndex        =   125
            Tag             =   "00AC31006"
            Top             =   1035
            Width           =   3660
         End
         Begin VB.CheckBox Check1 
            BackColor       =   &H80000004&
            Caption         =   "Reporte por Concepto de Gastos"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   315
            Index           =   88
            Left            =   15
            TabIndex        =   124
            Tag             =   "00AC31004"
            Top             =   705
            Width           =   3660
         End
         Begin VB.CheckBox Check1 
            BackColor       =   &H80000004&
            Caption         =   "Resumen de Gastos por Condominio"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   315
            Index           =   87
            Left            =   15
            TabIndex        =   123
            Tag             =   "00AC31003"
            Top             =   375
            Width           =   3660
         End
         Begin VB.CheckBox Check1 
            BackColor       =   &H80000004&
            Caption         =   "Facturación                                   »"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   315
            Index           =   85
            Left            =   60
            TabIndex        =   122
            Tag             =   "45AC31001"
            Top             =   45
            Width           =   3660
         End
         Begin VB.Line lnePerfil 
            BorderColor     =   &H00808080&
            Index           =   26
            X1              =   3690
            X2              =   3690
            Y1              =   0
            Y2              =   2220
         End
         Begin VB.Line lnePerfil 
            BorderColor     =   &H00808080&
            Index           =   25
            X1              =   0
            X2              =   3800
            Y1              =   1695
            Y2              =   1695
         End
      End
      Begin VB.PictureBox pctPerfil 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1725
         Index           =   44
         Left            =   15
         ScaleHeight     =   1725
         ScaleWidth      =   3720
         TabIndex        =   127
         Top             =   15
         Width           =   3720
      End
   End
   Begin VB.PictureBox pctPerfil 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      DrawStyle       =   2  'Dot
      DrawWidth       =   16847
      ForeColor       =   &H80000008&
      Height          =   3240
      Index           =   39
      Left            =   3840
      ScaleHeight     =   3240
      ScaleWidth      =   3285
      TabIndex        =   109
      Tag             =   "0"
      Top             =   1095
      Visible         =   0   'False
      Width           =   3285
      Begin VB.PictureBox pctPerfil 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         DrawStyle       =   2  'Dot
         DrawWidth       =   16847
         ForeColor       =   &H80000008&
         Height          =   3210
         Index           =   40
         Left            =   15
         ScaleHeight     =   3210
         ScaleWidth      =   3240
         TabIndex        =   110
         Top             =   15
         Width           =   3240
         Begin VB.CheckBox Check1 
            Caption         =   "Revisión..."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   315
            Index           =   127
            Left            =   15
            TabIndex        =   263
            Tag             =   "00AC314"
            Top             =   660
            Width           =   3210
         End
         Begin VB.CheckBox Check1 
            Caption         =   "Novedades....."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   315
            Index           =   124
            Left            =   15
            TabIndex        =   262
            Tag             =   "00AC313"
            Top             =   345
            Width           =   3210
         End
         Begin VB.CheckBox Check1 
            Caption         =   "Asignación de Gastos"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   315
            Index           =   84
            Left            =   15
            TabIndex        =   118
            Tag             =   "00AC301"
            Top             =   30
            Width           =   3210
         End
         Begin VB.CheckBox Check1 
            Caption         =   "PreFacturación"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   315
            Index           =   82
            Left            =   15
            TabIndex        =   117
            Tag             =   "00AC303"
            Top             =   975
            Width           =   3210
         End
         Begin VB.CheckBox Check1 
            Caption         =   "Parámetros de Facturación"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   315
            Index           =   81
            Left            =   15
            TabIndex        =   116
            Tag             =   "00AC304"
            Top             =   1290
            Width           =   3210
         End
         Begin VB.CheckBox Check1 
            Caption         =   "Registrar Gastos No Comunes"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   315
            Index           =   80
            Left            =   15
            TabIndex        =   115
            Tag             =   "00AC305"
            Top             =   1605
            Width           =   3210
         End
         Begin VB.CheckBox Check1 
            Caption         =   "Emisión de Facturas"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   315
            Index           =   79
            Left            =   15
            TabIndex        =   114
            Tag             =   "00AC306"
            Top             =   1920
            Width           =   3210
         End
         Begin VB.CheckBox Check1 
            Caption         =   "Revertir Facturación"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   315
            Index           =   78
            Left            =   15
            TabIndex        =   113
            Tag             =   "00AC307"
            Top             =   2235
            Width           =   3210
         End
         Begin VB.CheckBox Check1 
            Caption         =   "Cartas y Telegramas"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   315
            Index           =   76
            Left            =   15
            TabIndex        =   112
            Tag             =   "00AC309"
            Top             =   2550
            Width           =   3210
         End
         Begin VB.CheckBox Check1 
            Caption         =   "[Fact] Consultas y Reportes     »"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   315
            Index           =   75
            Left            =   15
            TabIndex        =   111
            Tag             =   "42AC310"
            Top             =   2880
            Width           =   3210
         End
         Begin VB.Line lnePerfil 
            BorderColor     =   &H00808080&
            Index           =   23
            X1              =   0
            X2              =   3255
            Y1              =   3195
            Y2              =   3195
         End
      End
      Begin VB.PictureBox pctPerfil 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         DrawStyle       =   2  'Dot
         DrawWidth       =   16847
         ForeColor       =   &H80000008&
         Height          =   3225
         Index           =   41
         Left            =   15
         ScaleHeight     =   3225
         ScaleWidth      =   3270
         TabIndex        =   119
         Top             =   15
         Width           =   3270
         Begin VB.Line lnePerfil 
            BorderColor     =   &H00808080&
            Index           =   24
            X1              =   3240
            X2              =   3240
            Y1              =   -15
            Y2              =   3180
         End
      End
   End
   Begin VB.PictureBox pctPerfil 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1845
      Index           =   54
      Left            =   7770
      ScaleHeight     =   1845
      ScaleWidth      =   2790
      TabIndex        =   151
      Tag             =   "1"
      Top             =   2490
      Visible         =   0   'False
      Width           =   2790
      Begin VB.PictureBox pctPerfil 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1815
         Index           =   55
         Left            =   15
         ScaleHeight     =   1815
         ScaleWidth      =   2745
         TabIndex        =   152
         Top             =   15
         Width           =   2745
         Begin VB.CheckBox Check1 
            Caption         =   "Re-Impresión Canc. Gastos"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   315
            Index           =   40
            Left            =   15
            TabIndex        =   261
            Tag             =   "00AC4002(7)"
            Top             =   1470
            Width           =   2715
         End
         Begin VB.CheckBox Check1 
            Caption         =   "Emisión Canc. Gastos"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   315
            Index           =   27
            Left            =   15
            TabIndex        =   260
            Tag             =   "00AC4002(6)"
            Top             =   1140
            Width           =   2715
         End
         Begin VB.CheckBox Check1 
            Caption         =   "Depóstios en Tránsito"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   315
            Index           =   108
            Left            =   15
            TabIndex        =   155
            Tag             =   "00AC4002(4)"
            Top             =   750
            Width           =   2715
         End
         Begin VB.CheckBox Check1 
            Caption         =   "&2.- Resumen de Caja"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   315
            Index           =   107
            Left            =   15
            TabIndex        =   154
            Tag             =   "00AC4002(2)"
            Top             =   345
            Width           =   2715
         End
         Begin VB.CheckBox Check1 
            Caption         =   "&1.- Reporte General     "
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   315
            Index           =   106
            Left            =   15
            TabIndex        =   153
            Tag             =   "00AC4002(1)"
            Top             =   15
            Width           =   2715
         End
         Begin VB.Line lnePerfil 
            BorderColor     =   &H00808080&
            Index           =   63
            X1              =   30
            X2              =   2700
            Y1              =   675
            Y2              =   675
         End
         Begin VB.Line lnePerfil 
            BorderColor     =   &H00FFFFFF&
            BorderWidth     =   2
            Index           =   64
            X1              =   30
            X2              =   2700
            Y1              =   690
            Y2              =   690
         End
         Begin VB.Line lnePerfil 
            BorderColor     =   &H00808080&
            Index           =   59
            X1              =   30
            X2              =   2700
            Y1              =   1095
            Y2              =   1095
         End
         Begin VB.Line lnePerfil 
            BorderColor     =   &H00808080&
            Index           =   32
            X1              =   0
            X2              =   2750
            Y1              =   1800
            Y2              =   1800
         End
         Begin VB.Line lnePerfil 
            BorderColor     =   &H00FFFFFF&
            BorderWidth     =   2
            Index           =   60
            X1              =   30
            X2              =   2700
            Y1              =   1110
            Y2              =   1110
         End
      End
      Begin VB.PictureBox pctPerfil 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1830
         Index           =   56
         Left            =   15
         ScaleHeight     =   1830
         ScaleWidth      =   2775
         TabIndex        =   156
         Top             =   15
         Width           =   2775
      End
   End
   Begin VB.PictureBox pctPerfil 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3225
      Index           =   3
      Left            =   2385
      ScaleHeight     =   3225
      ScaleWidth      =   2505
      TabIndex        =   10
      Tag             =   "l"
      Top             =   2175
      Visible         =   0   'False
      Width           =   2500
      Begin VB.PictureBox pctPerfil 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   3195
         Index           =   4
         Left            =   15
         ScaleHeight     =   3195
         ScaleWidth      =   2460
         TabIndex        =   11
         Top             =   15
         Width           =   2460
         Begin VB.CheckBox Check1 
            Caption         =   "Condiciones de Pago"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   270
            Index           =   13
            Left            =   15
            TabIndex        =   22
            Tag             =   "00AC0301(9)"
            Top             =   2865
            Width           =   2430
         End
         Begin VB.CheckBox Check1 
            Caption         =   "Tipos de Contratos"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   315
            Index           =   12
            Left            =   15
            TabIndex        =   21
            Tag             =   "00AC0301(8)"
            Top             =   2535
            Width           =   2430
         End
         Begin VB.CheckBox Check1 
            Caption         =   "Estado Civil"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   315
            Index           =   11
            Left            =   15
            TabIndex        =   20
            Tag             =   "00AC0301(7)"
            Top             =   2220
            Width           =   2430
         End
         Begin VB.CheckBox Check1 
            Caption         =   "Ramo"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   315
            Index           =   10
            Left            =   15
            TabIndex        =   19
            Tag             =   "00AC0301(6)"
            Top             =   1905
            Width           =   2430
         End
         Begin VB.CheckBox Check1 
            Caption         =   "Actividad"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   315
            Index           =   9
            Left            =   15
            TabIndex        =   18
            Tag             =   "00AC0301(5)"
            Top             =   1590
            Width           =   2430
         End
         Begin VB.CheckBox Check1 
            Caption         =   "Tipo de Inmueble"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   315
            Index           =   8
            Left            =   15
            TabIndex        =   17
            Tag             =   "00AC0301(4)"
            Top             =   1275
            Width           =   2430
         End
         Begin VB.CheckBox Check1 
            Caption         =   "Ocupación o Cargos"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   315
            Index           =   7
            Left            =   15
            TabIndex        =   15
            Tag             =   "00AC0301(3)"
            Top             =   960
            Width           =   2430
         End
         Begin VB.CheckBox Check1 
            Caption         =   "Bancos"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   315
            Index           =   6
            Left            =   15
            TabIndex        =   14
            Tag             =   "00AC0301(2)"
            Top             =   645
            Width           =   2430
         End
         Begin VB.CheckBox Check1 
            Caption         =   "Estados"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   315
            Index           =   5
            Left            =   15
            TabIndex        =   13
            Tag             =   "00AC0301(1)"
            Top             =   330
            Width           =   2430
         End
         Begin VB.CheckBox Check1 
            Caption         =   "Ciudades"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   315
            Index           =   4
            Left            =   15
            TabIndex        =   12
            Tag             =   "00AC0301(0)"
            Top             =   15
            Width           =   2430
         End
      End
      Begin VB.PictureBox pctPerfil 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   3210
         Index           =   5
         Left            =   15
         ScaleHeight     =   3210
         ScaleWidth      =   2505
         TabIndex        =   16
         Top             =   15
         Width           =   2500
         Begin VB.Line lnePerfil 
            BorderColor     =   &H00808080&
            BorderWidth     =   2
            Index           =   10
            X1              =   2460
            X2              =   2460
            Y1              =   -10
            Y2              =   3210
         End
      End
   End
   Begin VB.PictureBox pctPerfil 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1005
      Index           =   24
      Left            =   4710
      ScaleHeight     =   1005
      ScaleWidth      =   2385
      TabIndex        =   71
      Tag             =   "1"
      Top             =   2730
      Visible         =   0   'False
      Width           =   2385
      Begin VB.PictureBox pctPerfil 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   975
         Index           =   25
         Left            =   15
         ScaleHeight     =   975
         ScaleWidth      =   2355
         TabIndex        =   72
         Top             =   15
         Width           =   2355
         Begin VB.CheckBox Check1 
            Caption         =   "Junta de Condominio"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   315
            Index           =   52
            Left            =   15
            TabIndex        =   75
            Tag             =   "00AC20701(2)"
            Top             =   645
            Width           =   2325
         End
         Begin VB.CheckBox Check1 
            Caption         =   "Propietarios"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   315
            Index           =   51
            Left            =   15
            TabIndex        =   74
            Tag             =   "00AC20701(1)"
            Top             =   330
            Width           =   2325
         End
         Begin VB.CheckBox Check1 
            Caption         =   "Proveedores"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   315
            Index           =   50
            Left            =   15
            TabIndex        =   73
            Tag             =   "00AC20701(0)"
            Top             =   15
            Width           =   2325
         End
         Begin VB.Line lnePerfil 
            BorderColor     =   &H00808080&
            Index           =   22
            X1              =   2340
            X2              =   2340
            Y1              =   15
            Y2              =   975
         End
         Begin VB.Line lnePerfil 
            BorderColor     =   &H00808080&
            Index           =   21
            X1              =   0
            X2              =   2500
            Y1              =   960
            Y2              =   960
         End
      End
      Begin VB.PictureBox pctPerfil 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   990
         Index           =   26
         Left            =   15
         ScaleHeight     =   990
         ScaleWidth      =   2370
         TabIndex        =   76
         Top             =   15
         Width           =   2370
      End
   End
   Begin VB.PictureBox pctPerfil 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   690
      Index           =   21
      Left            =   4710
      ScaleHeight     =   690
      ScaleWidth      =   2850
      TabIndex        =   66
      Tag             =   "1"
      Top             =   2415
      Visible         =   0   'False
      Width           =   2850
      Begin VB.PictureBox pctPerfil 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   660
         Index           =   22
         Left            =   15
         ScaleHeight     =   660
         ScaleWidth      =   2820
         TabIndex        =   67
         Top             =   15
         Width           =   2820
         Begin VB.CheckBox Check1 
            Caption         =   "Actualizar Plan de Pagos"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   315
            Index           =   49
            Left            =   15
            TabIndex        =   69
            Tag             =   "00"
            Top             =   330
            Width           =   2775
         End
         Begin VB.CheckBox Check1 
            Caption         =   "Asignar Fecha de Pagos"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   315
            Index           =   48
            Left            =   15
            TabIndex        =   68
            Tag             =   "00"
            Top             =   15
            Width           =   2775
         End
         Begin VB.Line lnePerfil 
            BorderColor     =   &H00808080&
            Index           =   20
            X1              =   15
            X2              =   2820
            Y1              =   645
            Y2              =   645
         End
         Begin VB.Line lnePerfil 
            BorderColor     =   &H00808080&
            Index           =   19
            X1              =   2805
            X2              =   2805
            Y1              =   0
            Y2              =   660
         End
      End
      Begin VB.PictureBox pctPerfil 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   675
         Index           =   23
         Left            =   15
         ScaleHeight     =   675
         ScaleWidth      =   2835
         TabIndex        =   70
         Top             =   15
         Width           =   2835
      End
   End
   Begin VB.PictureBox pctPerfil 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1635
      Index           =   15
      Left            =   6960
      ScaleHeight     =   1635
      ScaleWidth      =   2370
      TabIndex        =   55
      Tag             =   "2"
      Top             =   4005
      Visible         =   0   'False
      Width           =   2370
      Begin VB.PictureBox pctPerfil 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1605
         Index           =   16
         Left            =   15
         ScaleHeight     =   1605
         ScaleWidth      =   2340
         TabIndex        =   56
         Top             =   15
         Width           =   2340
         Begin VB.CheckBox Check1 
            Caption         =   "Cuentas de Fondos"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   315
            Index           =   37
            Left            =   15
            TabIndex        =   61
            Tag             =   "00AC1060501(4)"
            Top             =   1260
            Width           =   2310
         End
         Begin VB.CheckBox Check1 
            Caption         =   "Gastos Fijos"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   315
            Index           =   36
            Left            =   15
            TabIndex        =   60
            Tag             =   "00AC1060501(3)"
            Top             =   960
            Width           =   2310
         End
         Begin VB.CheckBox Check1 
            Caption         =   "Gastos No Comunes"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   315
            Index           =   35
            Left            =   15
            TabIndex        =   59
            Tag             =   "00AC1060501(2)"
            Top             =   645
            Width           =   2310
         End
         Begin VB.CheckBox Check1 
            Caption         =   "Gastos Comunes"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   315
            Index           =   34
            Left            =   15
            TabIndex        =   58
            Tag             =   "00AC1060501(1)"
            Top             =   330
            Width           =   2310
         End
         Begin VB.CheckBox Check1 
            Caption         =   "Todos los Conceptos"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   315
            Index           =   33
            Left            =   15
            TabIndex        =   57
            Tag             =   "00AC1060501(0)"
            Top             =   15
            Width           =   2310
         End
         Begin VB.Line lnePerfil 
            BorderColor     =   &H00808080&
            Index           =   57
            X1              =   2325
            X2              =   2325
            Y1              =   1590
            Y2              =   0
         End
         Begin VB.Line lnePerfil 
            BorderColor     =   &H00808080&
            Index           =   56
            X1              =   0
            X2              =   2500
            Y1              =   1590
            Y2              =   1590
         End
      End
      Begin VB.PictureBox pctPerfil 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1620
         Index           =   17
         Left            =   15
         ScaleHeight     =   1620
         ScaleWidth      =   2355
         TabIndex        =   62
         Top             =   15
         Width           =   2355
      End
   End
   Begin VB.PictureBox pctPerfil 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2610
      Index           =   12
      Left            =   3720
      ScaleHeight     =   2610
      ScaleWidth      =   3465
      TabIndex        =   44
      Tag             =   "1"
      Top             =   2700
      Visible         =   0   'False
      Width           =   3465
      Begin VB.PictureBox pctPerfil 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   2580
         Index           =   13
         Left            =   15
         ScaleHeight     =   2580
         ScaleWidth      =   3435
         TabIndex        =   45
         Top             =   15
         Width           =   3435
         Begin VB.CheckBox Check1 
            Caption         =   "Ficha de Inmuebles"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   315
            Index           =   32
            Left            =   15
            TabIndex        =   53
            Tag             =   "00AC10601"
            Top             =   30
            Width           =   3390
         End
         Begin VB.CheckBox Check1 
            Caption         =   "Lista de Inmuebles"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   315
            Index           =   31
            Left            =   15
            TabIndex        =   52
            Tag             =   "00AC10602"
            Top             =   345
            Width           =   3390
         End
         Begin VB.CheckBox Check1 
            Caption         =   "Ficha de Propietarios"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   315
            Index           =   30
            Left            =   15
            TabIndex        =   51
            Tag             =   "00AC10603"
            Top             =   660
            Width           =   3390
         End
         Begin VB.CheckBox Check1 
            Caption         =   "Lista de Propietarios"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   315
            Index           =   29
            Left            =   15
            TabIndex        =   50
            Tag             =   "00AC10604"
            Top             =   975
            Width           =   3390
         End
         Begin VB.CheckBox Check1 
            Caption         =   "Catálogo Concepto de Gastos    »"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   315
            Index           =   28
            Left            =   15
            TabIndex        =   49
            Tag             =   "15AC10605"
            Top             =   1290
            Width           =   3390
         End
         Begin VB.CheckBox Check1 
            Caption         =   "Estado de Cuenta de Fondos"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   315
            Index           =   25
            Left            =   15
            TabIndex        =   48
            Tag             =   "00AC10608"
            Top             =   1605
            Width           =   3390
         End
         Begin VB.CheckBox Check1 
            Caption         =   "Estadísticas por Condominio"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   315
            Index           =   24
            Left            =   15
            TabIndex        =   47
            Tag             =   "00AC10609"
            Top             =   1920
            Width           =   3390
         End
         Begin VB.CheckBox Check1 
            Caption         =   "Reporte de Finiquito"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   315
            Index           =   17
            Left            =   15
            TabIndex        =   46
            Tag             =   "00AC10610"
            Top             =   2235
            Width           =   3390
         End
         Begin VB.Line lnePerfil 
            BorderColor     =   &H00808080&
            Index           =   16
            X1              =   -15
            X2              =   3450
            Y1              =   2565
            Y2              =   2565
         End
         Begin VB.Line lnePerfil 
            BorderColor     =   &H00808080&
            Index           =   15
            X1              =   3420
            X2              =   3420
            Y1              =   0
            Y2              =   3210
         End
      End
      Begin VB.PictureBox pctPerfil 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   2595
         Index           =   14
         Left            =   15
         ScaleHeight     =   2595
         ScaleWidth      =   3450
         TabIndex        =   54
         Top             =   15
         Width           =   3450
      End
   End
   Begin VB.PictureBox pctPerfil 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1005
      Index           =   9
      Left            =   3735
      ScaleHeight     =   1005
      ScaleWidth      =   2850
      TabIndex        =   38
      Tag             =   "1"
      Top             =   2385
      Visible         =   0   'False
      Width           =   2850
      Begin VB.PictureBox pctPerfil 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   975
         Index           =   10
         Left            =   15
         ScaleHeight     =   975
         ScaleWidth      =   2820
         TabIndex        =   39
         Top             =   15
         Width           =   2820
         Begin VB.CheckBox Check1 
            Caption         =   "3.- Telegramas"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   16
            Left            =   15
            TabIndex        =   42
            Tag             =   "00AC1050(2)"
            Top             =   675
            Width           =   2790
         End
         Begin VB.CheckBox Check1 
            Caption         =   "2.- Carta más 3 meses"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   315
            Index           =   15
            Left            =   15
            TabIndex        =   41
            Tag             =   "00AC1050(1)"
            Top             =   340
            Width           =   2790
         End
         Begin VB.CheckBox Check1 
            Caption         =   "1.- Carta 3 meses"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   315
            Index           =   14
            Left            =   15
            TabIndex        =   40
            Tag             =   "00AC1050(0)"
            Top             =   15
            Width           =   2790
         End
         Begin VB.Line lnePerfil 
            BorderColor     =   &H00808080&
            Index           =   14
            X1              =   0
            X2              =   2820
            Y1              =   960
            Y2              =   960
         End
         Begin VB.Line lnePerfil 
            BorderColor     =   &H00808080&
            Index           =   13
            X1              =   2805
            X2              =   2805
            Y1              =   0
            Y2              =   975
         End
      End
      Begin VB.PictureBox pctPerfil 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   990
         Index           =   11
         Left            =   15
         ScaleHeight     =   990
         ScaleWidth      =   2835
         TabIndex        =   43
         Top             =   15
         Width           =   2835
      End
   End
   Begin VB.PictureBox pctPerfil 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      DrawStyle       =   2  'Dot
      DrawWidth       =   16847
      ForeColor       =   &H80000008&
      Height          =   705
      Index           =   78
      Left            =   4980
      ScaleHeight     =   705
      ScaleWidth      =   2775
      TabIndex        =   207
      Tag             =   "1"
      Top             =   2085
      Visible         =   0   'False
      Width           =   2775
      Begin VB.PictureBox pctPerfil 
         Appearance      =   0  'Flat
         BackColor       =   &H80000001&
         BorderStyle     =   0  'None
         DrawStyle       =   2  'Dot
         DrawWidth       =   16847
         ForeColor       =   &H80000008&
         Height          =   675
         Index           =   79
         Left            =   15
         ScaleHeight     =   675
         ScaleWidth      =   2745
         TabIndex        =   208
         Top             =   15
         Width           =   2745
         Begin VB.CheckBox Check1 
            BackColor       =   &H80000001&
            Caption         =   "1"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   315
            Index           =   150
            Left            =   15
            TabIndex        =   211
            Tag             =   "00"
            Top             =   200
            Visible         =   0   'False
            Width           =   2700
         End
         Begin VB.CheckBox Check1 
            BackColor       =   &H80000002&
            Caption         =   "Asignación de Chequeras"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   315
            Index           =   149
            Left            =   15
            TabIndex        =   210
            Tag             =   "00AC50402"
            Top             =   330
            Width           =   2700
         End
         Begin VB.CheckBox Check1 
            BackColor       =   &H80000001&
            Caption         =   "Registro de Chequeras"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   315
            Index           =   148
            Left            =   15
            TabIndex        =   209
            Tag             =   "00AC50401"
            Top             =   0
            Width           =   2700
         End
         Begin VB.Line lnePerfil 
            BorderColor     =   &H00808080&
            Index           =   48
            X1              =   0
            X2              =   2730
            Y1              =   660
            Y2              =   660
         End
         Begin VB.Line lnePerfil 
            BorderColor     =   &H00808080&
            Index           =   47
            X1              =   2730
            X2              =   2730
            Y1              =   0
            Y2              =   975
         End
      End
      Begin VB.PictureBox pctPerfil 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         DrawStyle       =   2  'Dot
         DrawWidth       =   16847
         ForeColor       =   &H80000008&
         Height          =   690
         Index           =   80
         Left            =   15
         ScaleHeight     =   690
         ScaleWidth      =   2760
         TabIndex        =   212
         Top             =   15
         Width           =   2760
      End
   End
   Begin VB.PictureBox pctPerfil 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2280
      Index           =   81
      Left            =   4530
      ScaleHeight     =   2280
      ScaleWidth      =   3225
      TabIndex        =   213
      Tag             =   "1"
      Top             =   3030
      Visible         =   0   'False
      Width           =   3225
      Begin VB.PictureBox pctPerfil 
         Appearance      =   0  'Flat
         BackColor       =   &H80000001&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   2250
         Index           =   82
         Left            =   15
         ScaleHeight     =   2250
         ScaleWidth      =   3195
         TabIndex        =   214
         Top             =   15
         Width           =   3195
         Begin VB.CheckBox Check1 
            BackColor       =   &H80000001&
            Caption         =   "Lista de Bancos"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   315
            Index           =   157
            Left            =   15
            TabIndex        =   221
            Tag             =   "00AC50701"
            Top             =   15
            Width           =   3150
         End
         Begin VB.CheckBox Check1 
            BackColor       =   &H80000001&
            Caption         =   "Consulta de Saldos"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   315
            Index           =   156
            Left            =   15
            TabIndex        =   220
            Tag             =   "00AC50702"
            Top             =   330
            Width           =   3150
         End
         Begin VB.CheckBox Check1 
            BackColor       =   &H80000001&
            Caption         =   "Reporte de Transacciones"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   315
            Index           =   155
            Left            =   15
            TabIndex        =   219
            Tag             =   "00AC50703"
            Top             =   645
            Width           =   3150
         End
         Begin VB.CheckBox Check1 
            BackColor       =   &H80000001&
            Caption         =   "Estados de Cuenta"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   315
            Index           =   154
            Left            =   15
            TabIndex        =   218
            Tag             =   "00AC50704"
            Top             =   960
            Width           =   3150
         End
         Begin VB.CheckBox Check1 
            BackColor       =   &H80000001&
            Caption         =   "Disponibilidad Bancaria"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   315
            Index           =   153
            Left            =   15
            TabIndex        =   217
            Tag             =   "00AC50705"
            Top             =   1275
            Width           =   3150
         End
         Begin VB.CheckBox Check1 
            BackColor       =   &H80000001&
            Caption         =   "Relación de Cheques Devueltos"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   315
            Index           =   152
            Left            =   15
            TabIndex        =   216
            Tag             =   "00AC50706"
            Top             =   1590
            Width           =   3150
         End
         Begin VB.CheckBox Check1 
            BackColor       =   &H80000001&
            Caption         =   "Relación de Chequeras"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   315
            Index           =   151
            Left            =   15
            TabIndex        =   215
            Tag             =   "00AC50707"
            Top             =   1905
            Width           =   3150
         End
         Begin VB.Line lnePerfil 
            BorderColor     =   &H00808080&
            Index           =   50
            X1              =   30
            X2              =   3180
            Y1              =   2235
            Y2              =   2235
         End
         Begin VB.Line lnePerfil 
            BorderColor     =   &H00808080&
            Index           =   49
            X1              =   3180
            X2              =   3180
            Y1              =   2235
            Y2              =   0
         End
      End
      Begin VB.PictureBox pctPerfil 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         DrawStyle       =   2  'Dot
         DrawWidth       =   16847
         ForeColor       =   &H80000008&
         Height          =   2265
         Index           =   83
         Left            =   15
         ScaleHeight     =   2265
         ScaleWidth      =   3210
         TabIndex        =   222
         Top             =   15
         Width           =   3210
      End
   End
   Begin VB.PictureBox pctPerfil 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      DrawStyle       =   2  'Dot
      DrawWidth       =   16847
      ForeColor       =   &H80000008&
      Height          =   2280
      Index           =   75
      Left            =   7725
      ScaleHeight     =   2280
      ScaleWidth      =   3225
      TabIndex        =   197
      Tag             =   "0"
      Top             =   1095
      Visible         =   0   'False
      Width           =   3225
      Begin VB.PictureBox pctPerfil 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         DrawStyle       =   2  'Dot
         DrawWidth       =   16847
         ForeColor       =   &H80000008&
         Height          =   2250
         Index           =   76
         Left            =   15
         ScaleHeight     =   2250
         ScaleWidth      =   3195
         TabIndex        =   198
         Top             =   15
         Width           =   3195
         Begin VB.CheckBox Check1 
            Caption         =   "[Banco] Consultas y Reportes  »"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   315
            Index           =   147
            Left            =   15
            TabIndex        =   205
            Tag             =   "81AC507"
            Top             =   1905
            Width           =   3150
         End
         Begin VB.CheckBox Check1 
            Caption         =   "Conciliación Bancaria"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   315
            Index           =   146
            Left            =   15
            TabIndex        =   204
            Tag             =   "00AC506"
            Top             =   1590
            Width           =   3150
         End
         Begin VB.CheckBox Check1 
            Caption         =   "Registro Cheques Devueltos"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   315
            Index           =   145
            Left            =   15
            TabIndex        =   203
            Tag             =   "00AC505"
            Top             =   1275
            Width           =   3150
         End
         Begin VB.CheckBox Check1 
            Caption         =   "Administrador de Chequeras    »"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   315
            Index           =   144
            Left            =   15
            TabIndex        =   202
            Tag             =   "78AC504"
            Top             =   960
            Width           =   3150
         End
         Begin VB.CheckBox Check1 
            Caption         =   "Libro Banco"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   315
            Index           =   143
            Left            =   15
            TabIndex        =   201
            Tag             =   "00AC503"
            Top             =   645
            Width           =   3150
         End
         Begin VB.CheckBox Check1 
            Caption         =   "Cuentas Bancarias"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   315
            Index           =   142
            Left            =   15
            TabIndex        =   200
            Tag             =   "00AC502"
            Top             =   330
            Width           =   3150
         End
         Begin VB.CheckBox Check1 
            Caption         =   "Ficha Bancaria"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   315
            Index           =   141
            Left            =   15
            TabIndex        =   199
            Tag             =   "00AC501"
            Top             =   30
            Width           =   3150
         End
         Begin VB.Line lnePerfil 
            BorderColor     =   &H00808080&
            Index           =   46
            X1              =   3180
            X2              =   3180
            Y1              =   0
            Y2              =   2220
         End
         Begin VB.Line lnePerfil 
            BorderColor     =   &H00808080&
            Index           =   45
            X1              =   0
            X2              =   3225
            Y1              =   2235
            Y2              =   2235
         End
      End
      Begin VB.PictureBox pctPerfil 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         DrawStyle       =   2  'Dot
         DrawWidth       =   16847
         ForeColor       =   &H80000008&
         Height          =   2265
         Index           =   77
         Left            =   15
         ScaleHeight     =   2265
         ScaleWidth      =   3210
         TabIndex        =   206
         Top             =   15
         Width           =   3210
      End
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H80000001&
      Caption         =   "Relación de Deuda por Propietario"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      Index           =   83
      Left            =   0
      TabIndex        =   256
      Tag             =   "00"
      Top             =   0
      Visible         =   0   'False
      Width           =   3390
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Relación de Deuda por Propietario"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      Index           =   77
      Left            =   105
      TabIndex        =   255
      Tag             =   "00"
      Top             =   8070
      Visible         =   0   'False
      Width           =   3390
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Relación de Deuda por Propietario"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      Index           =   74
      Left            =   105
      TabIndex        =   254
      Tag             =   "00"
      Top             =   7650
      Visible         =   0   'False
      Width           =   3390
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Relación de Deuda por Propietario"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      Index           =   59
      Left            =   105
      TabIndex        =   253
      Tag             =   "00"
      Top             =   7275
      Visible         =   0   'False
      Width           =   3390
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Relación de Deuda por Propietario"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      Index           =   43
      Left            =   105
      TabIndex        =   252
      Tag             =   "00"
      Top             =   6870
      Visible         =   0   'False
      Width           =   3390
   End
   Begin VB.PictureBox pctPerfil 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1890
      Index           =   1
      Left            =   75
      ScaleHeight     =   1890
      ScaleWidth      =   2385
      TabIndex        =   3
      Tag             =   "0"
      Top             =   1095
      Visible         =   0   'False
      Width           =   2385
      Begin VB.PictureBox pctPerfil 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1860
         Index           =   0
         Left            =   15
         ScaleHeight     =   1860
         ScaleWidth      =   2355
         TabIndex        =   5
         Top             =   15
         Width           =   2355
         Begin VB.CheckBox Check1 
            Caption         =   "Cambiar Contraseña"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   315
            Index           =   166
            Left            =   30
            MaskColor       =   &H00800000&
            TabIndex        =   242
            Tag             =   "00AC021"
            Top             =   735
            Width           =   2295
         End
         Begin VB.CheckBox Check1 
            Caption         =   "Tablas del Sistema    »"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   315
            Index           =   2
            Left            =   30
            MaskColor       =   &H00800000&
            TabIndex        =   7
            Tag             =   "03AC03"
            Top             =   1050
            Width           =   2295
         End
         Begin VB.CheckBox Check1 
            Caption         =   "Seleccionar Inmueble"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   0
            Left            =   30
            MaskColor       =   &H00800000&
            TabIndex        =   9
            Tag             =   "00AC01"
            Top             =   15
            Width           =   2295
         End
         Begin VB.CheckBox Check1 
            Caption         =   "Seleccionar Usuario"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   315
            Index           =   1
            Left            =   30
            MaskColor       =   &H00800000&
            TabIndex        =   8
            Tag             =   "00AC02"
            Top             =   420
            Width           =   2295
         End
         Begin VB.CheckBox Check1 
            Caption         =   "Salir"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   315
            Index           =   3
            Left            =   30
            MaskColor       =   &H00800000&
            TabIndex        =   6
            Tag             =   "00AC06"
            Top             =   1500
            Width           =   2295
         End
         Begin VB.Line lnePerfil 
            BorderColor     =   &H00808080&
            Index           =   8
            X1              =   0
            X2              =   2460
            Y1              =   1845
            Y2              =   1845
         End
         Begin VB.Line lnePerfil 
            BorderColor     =   &H00808080&
            Index           =   9
            X1              =   2340
            X2              =   2340
            Y1              =   0
            Y2              =   1550
         End
         Begin VB.Line lnePerfil 
            BorderColor     =   &H00808080&
            Index           =   1
            X1              =   0
            X2              =   2460
            Y1              =   360
            Y2              =   360
         End
         Begin VB.Line lnePerfil 
            BorderColor     =   &H00808080&
            Index           =   3
            X1              =   0
            X2              =   2460
            Y1              =   1395
            Y2              =   1395
         End
         Begin VB.Line lnePerfil 
            BorderColor     =   &H00FFFFFF&
            BorderWidth     =   2
            Index           =   0
            X1              =   45
            X2              =   2505
            Y1              =   375
            Y2              =   375
         End
         Begin VB.Line lnePerfil 
            BorderColor     =   &H00FFFFFF&
            BorderWidth     =   2
            Index           =   2
            X1              =   0
            X2              =   2460
            Y1              =   1410
            Y2              =   1410
         End
      End
      Begin VB.PictureBox pctPerfil 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         ClipControls    =   0   'False
         ForeColor       =   &H80000008&
         Height          =   1875
         Index           =   2
         Left            =   15
         ScaleHeight     =   1875
         ScaleWidth      =   2370
         TabIndex        =   4
         Top             =   15
         Width           =   2370
      End
   End
   Begin VB.Frame fraPerfil 
      Height          =   1140
      Left            =   8985
      TabIndex        =   239
      Top             =   6540
      Width           =   2655
      Begin VB.CommandButton cmdPerfil 
         Caption         =   "&Salir"
         Height          =   765
         Index           =   1
         Left            =   1350
         Style           =   1  'Graphical
         TabIndex        =   241
         Top             =   240
         Width           =   1185
      End
      Begin VB.CommandButton cmdPerfil 
         Caption         =   "&Actualizar"
         Height          =   765
         Index           =   0
         Left            =   165
         Style           =   1  'Graphical
         TabIndex        =   240
         Top             =   240
         Width           =   1185
      End
   End
   Begin VB.ComboBox cmbPerfil 
      Height          =   315
      Left            =   1710
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   240
      Width           =   1770
   End
   Begin VB.PictureBox pctPerfil 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1965
      Index           =   72
      Left            =   5520
      ScaleHeight     =   1965
      ScaleWidth      =   3450
      TabIndex        =   188
      Tag             =   "2"
      Top             =   3375
      Visible         =   0   'False
      Width           =   3450
      Begin VB.PictureBox pctPerfil 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1935
         Index           =   73
         Left            =   15
         ScaleHeight     =   1935
         ScaleWidth      =   3420
         TabIndex        =   189
         Top             =   15
         Width           =   3420
         Begin VB.CheckBox Check1 
            Caption         =   "Estado Financiero Anual"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   315
            Index           =   140
            Left            =   15
            TabIndex        =   195
            Tag             =   "00AC40422(0)"
            Top             =   15
            Width           =   3375
         End
         Begin VB.CheckBox Check1 
            Caption         =   "Reporte de Cobranza Efectiva"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   315
            Index           =   139
            Left            =   15
            TabIndex        =   194
            Tag             =   "00AC40422(1)"
            Top             =   330
            Width           =   3375
         End
         Begin VB.CheckBox Check1 
            Caption         =   "Reporte de Deuda Mensual"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   315
            Index           =   138
            Left            =   15
            TabIndex        =   193
            Tag             =   "00AC40422(2)"
            Top             =   645
            Width           =   3375
         End
         Begin VB.CheckBox Check1 
            Caption         =   "Resumen Anual de Cobros"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   315
            Index           =   137
            Left            =   15
            TabIndex        =   192
            Tag             =   "00AC40422(3)"
            Top             =   960
            Width           =   3375
         End
         Begin VB.CheckBox Check1 
            Caption         =   "Relación Fondo - Deuda"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   315
            Index           =   136
            Left            =   15
            TabIndex        =   191
            Tag             =   "00AC40422(4)"
            Top             =   1275
            Width           =   3375
         End
         Begin VB.CheckBox Check1 
            Caption         =   "Estadísticas Mensuales de Cobro"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   315
            Index           =   135
            Left            =   15
            TabIndex        =   190
            Tag             =   "00AC40422(5)"
            Top             =   1590
            Width           =   3375
         End
         Begin VB.Line lnePerfil 
            BorderColor     =   &H00808080&
            Index           =   44
            X1              =   3405
            X2              =   3405
            Y1              =   0
            Y2              =   1935
         End
         Begin VB.Line lnePerfil 
            BorderColor     =   &H00808080&
            Index           =   43
            X1              =   0
            X2              =   3420
            Y1              =   1920
            Y2              =   1920
         End
      End
      Begin VB.PictureBox pctPerfil 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1950
         Index           =   74
         Left            =   15
         ScaleHeight     =   1950
         ScaleWidth      =   3435
         TabIndex        =   196
         Top             =   15
         Width           =   3435
      End
   End
   Begin VB.Label lblPerfil 
      Caption         =   "  Condominio"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      Index           =   2
      Left            =   1020
      TabIndex        =   224
      Tag             =   "6"
      Top             =   825
      Width           =   1185
   End
   Begin VB.Label lblPerfil 
      Caption         =   "  Cuentas x Pagar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      Index           =   3
      Left            =   2205
      TabIndex        =   223
      Tag             =   "18"
      Top             =   825
      Width           =   1680
   End
   Begin VB.Label lblPerfil 
      Caption         =   "  Utilidades"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      Index           =   9
      Left            =   9585
      TabIndex        =   28
      Tag             =   "84"
      Top             =   825
      Width           =   1170
   End
   Begin VB.Label lblPerfil 
      Caption         =   "  Nomina"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      Index           =   8
      Left            =   8640
      TabIndex        =   27
      Top             =   825
      Width           =   945
   End
   Begin VB.Label lblPerfil 
      Caption         =   "  Bancos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      Index           =   7
      Left            =   7695
      TabIndex        =   26
      Tag             =   "75"
      Top             =   825
      Width           =   945
   End
   Begin VB.Label lblPerfil 
      Caption         =   "  Cuentas x Cobrar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      Index           =   6
      Left            =   5925
      TabIndex        =   25
      Tag             =   "66"
      Top             =   825
      Width           =   1770
   End
   Begin VB.Label lblPerfil 
      Caption         =   "  Caja"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      Index           =   5
      Left            =   5205
      TabIndex        =   24
      Tag             =   "48"
      Top             =   825
      Width           =   720
   End
   Begin VB.Label lblPerfil 
      Caption         =   "  Facturación"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      Index           =   4
      Left            =   3885
      TabIndex        =   23
      Tag             =   "39"
      Top             =   825
      Width           =   1320
   End
   Begin VB.Label lblPerfil 
      Caption         =   "  Archivo"
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
      Index           =   1
      Left            =   75
      TabIndex        =   2
      Tag             =   "1"
      Top             =   825
      Width           =   885
   End
   Begin VB.Label lblPerfil 
      BackStyle       =   0  'Transparent
      Caption         =   "Nombre Usuario:"
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
      Left            =   150
      TabIndex        =   0
      Top             =   255
      Width           =   1620
   End
End
Attribute VB_Name = "FrmPerfiles"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
    'variables locales
    Dim intBoton As Boolean, booSw As Boolean
    Dim cnnTabla As New ADODB.Connection
    '
    Private Sub Check1_Click(Index As Integer)
    '
    If Left(Check1(Index).Tag, 2) <> "00" And booSw = False Then
    
        If pctPerfil(Left(Check1(Index).Tag, 2)).Visible Then
            pctPerfil(Left(Check1(Index).Tag, 2)).Visible = False
        Else
            Call muestra_submenu(pctPerfil(Left(Check1(Index).Tag, 2)))
        End If
        
    End If
    '
    End Sub

    Private Sub Check1_MouseMove(Index As Integer, Button As Integer, Shift As Integer, _
    X As Single, Y As Single)
    '
    On Error Resume Next
    For i = 0 To Check1.UBound
        
        If i = Index Then
            Check1(Index).BackColor = vbActiveTitleBar
            Check1(Index).ForeColor = vbActiveTitleBarText
                
        Else
            Check1(i).BackColor = vbButtonFace
            Check1(i).ForeColor = vbButtonText
        End If
        '
    Next
    
    End Sub

    Private Sub cmbPerfil_Click()
    Call Form_Click
    Call muestra_acceso
    End Sub

    Private Sub cmdPerfil_Click(Index As Integer)
    '
    Select Case Index
    
        Case 0  'Actualizar
        '--------------------
        If cmbPerfil.Text = "" Then Exit Sub
        Call Form_Click
        On Error Resume Next
        cnnTabla.BeginTrans
        Dim ctlPerfil As Control
        cnnTabla.Execute "DELETE FROM Perfiles WHERE Usuario='" & cmbPerfil & "'"
        Call RtnConfigUtility(True, "Guardando Información", "Espere por favor...", _
        "Activando Selección...")
        For i = Check1.LBound To Check1.UBound
            Call RtnProUtility(Check1(i).Caption, CInt(i * 6015 / Check1.UBound))
            If Check1(i).Value = 1 And Check1(i).Tag <> "" Then
                'Debug.Print CInt(I * 6015 / Check1.UBound)
                cnnTabla.Execute "INSERT INTO Perfiles(Usuario,Acceso) VALUES('" _
                & cmbPerfil & "','" & Mid(Check1(i).Tag, 3, Len(Check1(i).Tag)) & "');"
            End If
        Next
        Unload FrmUtility
        If Respuesta("Desea Completar esta operación") = True Then
            cnnTabla.CommitTrans
            Call rtnBitacora("Actualizar Perfil de Acceso Usuario: " & cmbPerfil)
        Else
            cnnTabla.RollbackTrans
            Call rtnBitacora("Cancelar Actualizar Perfil de Acceso Usuario: " & cmbPerfil)
        End If
        '
        'Salir
        Case 1: Unload Me: Set FrmPerfiles = Nothing
        
    End Select
    '
    End Sub

    Private Sub Form_Click()
    For i = 1 To lblPerfil.UBound
        lblPerfil(i).BorderStyle = 0
    Next
    Call ocultar_menu
    Call iniciar_checks
    End Sub

    Private Sub Form_Load()
    cmdPerfil(1).Picture = LoadResPicture("SALIR", vbResIcon)
    cmdPerfil(0).Picture = LoadResPicture("GUARDAR", vbResIcon)
    Dim rstUser As New ADODB.Recordset
    cnnTabla.Open cnnOLEDB & gcPath & "\tablas.mdb"
    rstUser.Open "SELECT * FROM USUARIOS WHERE Nivel >=" & gcNivel, _
    cnnTabla, adOpenStatic, adLockReadOnly
    With rstUser
        If Not .EOF Or .BOF Then .MoveFirst
        Do
            cmbPerfil.AddItem (!NombreUsuario)
            .MoveNext
        Loop Until .EOF
        .Close
    End With
    Set rstUser = Nothing
    End Sub

    Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    For i = 1 To lblPerfil.UBound
        lblPerfil(i).BackColor = &H8000000F
        lblPerfil(i).ForeColor = &H0&
        Call iniciar_checks
    Next
    End Sub

    Private Sub Form_Unload(Cancel As Integer)
    cnnTabla.Close: Set cnnTabla = Nothing
    End Sub

    Private Sub lblPerfil_Click(Index As Integer)
    '
    If Index <> 0 Then
        For i = 1 To lblPerfil.UBound
            If i = Index Then
                lblPerfil(i).BackColor = vbButtonFace
                lblPerfil(i).ForeColor = vbButtonText
                lblPerfil(i).BorderStyle = 1
                lblPerfil(i).Refresh
                intBoton = True
            Else
                lblPerfil(i).BorderStyle = 0
            End If
        Next
    End If
    If lblPerfil(Index).Tag <> "" Then Call muestra_Menu(pctPerfil(lblPerfil(Index).Tag))
    '
    End Sub

    Private Sub muestra_Menu(ctlPicture As PictureBox)
    'variables locales
    Dim intAncho As Integer
    Dim intLargo As Integer
    Dim intSalto As Integer
    '
    Call ocultar_menu
    With ctlPicture
        intAncho = .Width
        intLargo = .Height
        intSalto = intAncho / 20
        .Width = 0
        .Height = 0
        .Visible = True
        For i = 0 To intAncho Step intSalto
            .Width = i
            .Height = i * intLargo / intAncho
        Next
        If .Width < intAncho Then .Width = intAncho
        If .Height < intLargo Then .Height = intLargo
    End With
    '
    End Sub

    Private Sub lblPerfil_MouseMove(Index As Integer, Button As Integer, Shift As Integer, _
    X As Single, Y As Single)
    '
    If Index > 0 And Not intBoton And lblPerfil(Index).BorderStyle = 0 Then
        For i = 1 To lblPerfil.UBound
            If i = Index Then
                lblPerfil(i).BackColor = vbActiveTitleBar
                lblPerfil(i).ForeColor = vbActiveTitleBarText
            Else
                lblPerfil(i).BackColor = &H8000000F
                lblPerfil(i).ForeColor = &H0&
            End If
        Next
    End If
    intBoton = False
    End Sub

    Private Sub ocultar_menu()
    For J = lblPerfil.LBound To lblPerfil.UBound
        If lblPerfil(J).Tag <> "" Then pctPerfil(lblPerfil(J).Tag).Visible = False: pctPerfil(lblPerfil(J).Tag).Refresh
    Next
    'On Error Resume Next
    For i = 0 To Check1.UBound
        If Not Check1(i).Tag = "" Then
        If Left(Check1(i).Tag, 2) <> "00" Then pctPerfil(Left(Check1(i).Tag, 2)).Visible = False
        End If
    Next
    End Sub
    
    Private Sub iniciar_checks()
    'On Error Resume Next
    For i = 0 To Check1.UBound
        Check1(i).BackColor = vbButtonFace
        Check1(i).ForeColor = vbButtonText
    Next
    End Sub
    
    Private Sub muestra_submenu(ctlPicture As PictureBox)
    'variables locales
    Dim intAncho As Integer
    Dim intLargo As Integer
    Dim intSalto As Integer
    Dim intInd As Integer
    '
    'On Error Resume Next
    For i = 0 To Check1.UBound
        intInd = IIf(Left(Check1(i).Tag, 2) = "", 0, Left(Check1(i).Tag, 2))
        If intInd <> 0 Then
            If ctlPicture.Tag <= pctPerfil(intInd).Tag Then
                pctPerfil(intInd).Visible = False
            End If
        End If
    Next
    With ctlPicture
        intAncho = .Width
        intSalto = intAncho / 40
        .Width = 0
        .Visible = True
        For i = 0 To intAncho Step intSalto
            .Width = i
        Next
        If .Width < intAncho Then .Width = intAncho
    End With
    End Sub

    '---------------------------------------------------------------------------------------------
    '   Rutina:     muestra_acceso
    '
    '   Marca todos los menús activados para determinado usuario
    '   permite la edición del perfil del usuario seleccionado
    '---------------------------------------------------------------------------------------------
    Private Sub muestra_acceso()
    booSw = True
    'On Error Resume Next
    For i = 0 To Check1.UBound  'Iniciliza todos los checks
        Check1(i).Value = 0
    Next
    '
    Dim ctlPerfil As Control
    Dim rstPerfil As New ADODB.Recordset
    rstPerfil.Open "SELECT * FROM Perfiles WHERE Usuario='" & cmbPerfil & "'", _
    cnnTabla, adOpenKeyset, adLockOptimistic    'Selecciona el perfil del usuario
    With rstPerfil
        If Not .EOF Or Not .BOF Then .MoveFirst
        Do Until .EOF
        MousePointer = vbHourglass
        For J = Check1.LBound To Check1.UBound
            If Mid(Check1(J).Tag, 3, Len(Check1(J).Tag)) = !Acceso Then
                Check1(J).Value = 1
                Exit For
            End If
        Next
        .MoveNext
        Loop
        MousePointer = vbDefault
        .Close
    End With
    Set rstPerfil = Nothing
    booSw = False
    Call rtnBitacora("Consulta Perfíl Acceso de " & cmbPerfil)
    End Sub

