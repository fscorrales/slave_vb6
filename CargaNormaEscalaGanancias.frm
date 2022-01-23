VERSION 5.00
Begin VB.Form CargaNormaEscalaGanancias 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Normas determinativas de Escalas 4ta Categoría Ganancias"
   ClientHeight    =   1770
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   6855
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1770
   ScaleWidth      =   6855
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Caption         =   "Norma Legal"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   975
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   6615
      Begin VB.TextBox txtFecha 
         Height          =   285
         Left            =   4920
         TabIndex        =   1
         Top             =   480
         Width           =   1575
      End
      Begin VB.TextBox txtNormaLegal 
         Height          =   285
         Left            =   1440
         TabIndex        =   0
         Top             =   480
         Width           =   1575
      End
      Begin VB.Label Label8 
         Caption         =   "Fecha"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Left            =   3600
         TabIndex        =   5
         Top             =   480
         Width           =   1455
      End
      Begin VB.Label Label7 
         Caption         =   "Nro. Norma"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   480
         Width           =   1455
      End
   End
   Begin VB.CommandButton cmdAgregar 
      BackColor       =   &H008080FF&
      Caption         =   "Guardar Datos"
      Height          =   495
      Left            =   2520
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1200
      Width           =   1815
   End
End
Attribute VB_Name = "CargaNormaEscalaGanancias"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()

    Call CenterMe(CargaNormaEscalaGanancias, 6950, 2200)

End Sub

Private Sub cmdAgregar_Click()

    GenerarNormaEscalaGanancias

End Sub

Private Sub Form_Unload(Cancel As Integer)

    strEditandoNormaEscalaGanancias = ""

End Sub

