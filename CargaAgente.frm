VERSION 5.00
Begin VB.Form CargaAgente 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Agentes"
   ClientHeight    =   2910
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   6870
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2910
   ScaleWidth      =   6870
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame3 
      Caption         =   "Datos del Agente"
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
      Height          =   2175
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6615
      Begin VB.CheckBox chkActivado 
         Alignment       =   1  'Right Justify
         Caption         =   "Agente  Activado"
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
         Height          =   315
         Left            =   3600
         TabIndex        =   6
         Top             =   1680
         Width           =   2350
      End
      Begin VB.TextBox txtLegajo 
         Height          =   285
         Left            =   1440
         TabIndex        =   5
         Top             =   1680
         Width           =   1575
      End
      Begin VB.TextBox txtCUIL 
         Height          =   285
         Left            =   4920
         TabIndex        =   2
         Top             =   480
         Width           =   1575
      End
      Begin VB.TextBox txtPuestoLaboral 
         Height          =   285
         Left            =   1440
         TabIndex        =   1
         Top             =   480
         Width           =   1575
      End
      Begin VB.TextBox txtDescripcion 
         Height          =   285
         Left            =   1440
         TabIndex        =   4
         Top             =   1080
         Width           =   5055
      End
      Begin VB.Label Label4 
         Caption         =   "Nro. Legajo"
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
         TabIndex        =   10
         Top             =   1680
         Width           =   1455
      End
      Begin VB.Label Label2 
         Caption         =   "CUIL / DNI"
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
         TabIndex        =   9
         Top             =   480
         Width           =   1455
      End
      Begin VB.Label Label1 
         Caption         =   "Puesto Laboral"
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
         TabIndex        =   7
         Top             =   480
         Width           =   1455
      End
      Begin VB.Label Label3 
         Caption         =   "Descripcion"
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
         TabIndex        =   3
         Top             =   1080
         Width           =   1455
      End
   End
   Begin VB.CommandButton cmdAgregar 
      BackColor       =   &H008080FF&
      Caption         =   "Guardar Datos"
      Height          =   495
      Left            =   2520
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   2400
      Width           =   1815
   End
End
Attribute VB_Name = "CargaAgente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdPrueba_Click()

    'Call CopiarTablaOrigenExterno
    Call SetNullToZero

End Sub

Private Sub Form_Load()

    Call CenterMe(CargaAgente, 6950, 3250)

End Sub

Private Sub cmdAgregar_Click()

    GenerarAgente

End Sub

Private Sub Form_Unload(Cancel As Integer)

    strEditandoAgente = ""

End Sub
