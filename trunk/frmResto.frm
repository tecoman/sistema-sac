VERSION 5.00
Begin VB.Form frmResto 
   Caption         =   "Resto"
   ClientHeight    =   2445
   ClientLeft      =   6405
   ClientTop       =   3945
   ClientWidth     =   3315
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   2445
   ScaleWidth      =   3315
   Begin VB.CommandButton Command1 
      Cancel          =   -1  'True
      Caption         =   "&Aceptar"
      Height          =   495
      Left            =   637
      TabIndex        =   6
      Top             =   1785
      Width           =   2040
   End
   Begin VB.TextBox TXT 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   2
      Left            =   1080
      Locked          =   -1  'True
      TabIndex        =   5
      Text            =   "0,00"
      Top             =   1170
      Width           =   2040
   End
   Begin VB.TextBox TXT 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   1
      Left            =   1080
      Locked          =   -1  'True
      TabIndex        =   4
      Text            =   "0,00"
      Top             =   690
      Width           =   2040
   End
   Begin VB.TextBox TXT 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   0
      Left            =   1080
      TabIndex        =   3
      Text            =   "0,00"
      Top             =   210
      Width           =   2040
   End
   Begin VB.Label LBL 
      Caption         =   "Resto:"
      Height          =   345
      Index           =   2
      Left            =   195
      TabIndex        =   2
      Top             =   1230
      Width           =   1215
   End
   Begin VB.Label LBL 
      Caption         =   "A Cobrar:"
      Height          =   345
      Index           =   1
      Left            =   195
      TabIndex        =   1
      Top             =   750
      Width           =   1215
   End
   Begin VB.Label LBL 
      Caption         =   "&Recibido:"
      Height          =   360
      Index           =   0
      Left            =   195
      TabIndex        =   0
      Top             =   255
      Width           =   1215
   End
End
Attribute VB_Name = "frmResto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public curBs As Currency   'monto a cobrar en caja

Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Form_Load()
TXT(1) = Format(curBs, "#,##0.00")
'Me.Show
TXT(0).SelStart = 0
TXT(0).SelLength = Len(TXT(0))
'TXT(0).SetFocus
End Sub

Private Sub TXT_KeyPress(Index As Integer, KeyAscii As Integer)
If Index = 0 Then
    If KeyAscii = Asc(".") Then KeyAscii = Asc(",")
    Call Validacion(KeyAscii, "0123456789,")
    If KeyAscii = 13 Then
        TXT(2) = Format(CCur(TXT(0) - curBs), "#,##0.00")
        TXT(0) = Format(TXT(0), "#,##0.00")
        TXT(0).SelStart = 0
        TXT(0).SelLength = Len(TXT(0))
    End If
End If
End Sub
