VERSION 5.00
Begin VB.Form frmToolTip 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "::Gestión de Cobranza::"
   ClientHeight    =   2655
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   3060
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2655
   ScaleWidth      =   3060
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Label lbl 
      Caption         =   "Nombre"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   4
      Left            =   1035
      TabIndex        =   4
      Top             =   570
      Width           =   1800
   End
   Begin VB.Label lbl 
      Caption         =   "Por:"
      Height          =   255
      Index           =   3
      Left            =   135
      TabIndex        =   3
      Top             =   615
      Width           =   765
   End
   Begin VB.Label lbl 
      Caption         =   "Gestión"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1515
      Index           =   2
      Left            =   120
      TabIndex        =   2
      Top             =   945
      Width           =   2835
   End
   Begin VB.Label lbl 
      Caption         =   "Nombre"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   1
      Left            =   1035
      TabIndex        =   1
      Top             =   240
      Width           =   1800
   End
   Begin VB.Label lbl 
      Caption         =   "Contacto:"
      Height          =   255
      Index           =   0
      Left            =   135
      TabIndex        =   0
      Top             =   240
      Width           =   765
   End
End
Attribute VB_Name = "frmToolTip"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
