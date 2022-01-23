VERSION 5.00
Begin VB.Form CargaDeduccionesGenerales 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Deducciones Generales"
   ClientHeight    =   4635
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   6855
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4635
   ScaleWidth      =   6855
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdAgregar 
      BackColor       =   &H008080FF&
      Caption         =   "Guardar Datos"
      Height          =   495
      Left            =   2520
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   4080
      Width           =   1815
   End
   Begin VB.Frame Frame4 
      Caption         =   "Importe Anual de Deducciones Generales"
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
      TabIndex        =   9
      Top             =   1800
      Width           =   6615
      Begin VB.TextBox txtAlquileres 
         Height          =   285
         Left            =   1440
         TabIndex        =   4
         Top             =   1080
         Width           =   1575
      End
      Begin VB.TextBox txtSeguroDeVida 
         Height          =   285
         Left            =   4920
         TabIndex        =   3
         Top             =   480
         Width           =   1575
      End
      Begin VB.TextBox txtServicioDomestico 
         Height          =   285
         Left            =   1440
         TabIndex        =   2
         Top             =   480
         Width           =   1575
      End
      Begin VB.TextBox txtHonorariosMedicos 
         Height          =   285
         Left            =   4920
         TabIndex        =   7
         Top             =   1680
         Width           =   1575
      End
      Begin VB.TextBox txtCuotaMedico 
         Height          =   285
         Left            =   4920
         TabIndex        =   5
         Top             =   1080
         Width           =   1575
      End
      Begin VB.TextBox txtDonaciones 
         Height          =   285
         Left            =   1440
         TabIndex        =   6
         Top             =   1680
         Width           =   1575
      End
      Begin VB.Label Label4 
         Caption         =   "Alquileres"
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
         TabIndex        =   20
         Top             =   1080
         Width           =   1455
      End
      Begin VB.Label Label5 
         Caption         =   "Seguro de Vida"
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
         TabIndex        =   14
         Top             =   480
         Width           =   1455
      End
      Begin VB.Label Label9 
         Caption         =   "Serv. Domest."
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
         TabIndex        =   13
         Top             =   480
         Width           =   1455
      End
      Begin VB.Label Label10 
         Caption         =   "H. Médicos"
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
         TabIndex        =   12
         Top             =   1680
         Width           =   1455
      End
      Begin VB.Label Label11 
         Caption         =   "Cuota Médico Asistencial"
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
         Height          =   495
         Left            =   3600
         TabIndex        =   11
         Top             =   1080
         Width           =   1455
      End
      Begin VB.Label Label12 
         Caption         =   "Donaciones"
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
   End
   Begin VB.Frame Frame2 
      Caption         =   "Datos Identificatorios"
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
      Height          =   1575
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6615
      Begin VB.TextBox txtFecha 
         Height          =   285
         Left            =   4920
         TabIndex        =   1
         Top             =   360
         Width           =   1575
      End
      Begin VB.TextBox txtPuestoLaboral 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1440
         TabIndex        =   17
         Top             =   360
         Width           =   1575
      End
      Begin VB.TextBox txtDescripcion 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1440
         TabIndex        =   15
         Top             =   960
         Width           =   5055
      End
      Begin VB.Label Label3 
         Caption         =   "Fecha DDJJ"
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
         TabIndex        =   19
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label Label2 
         Caption         =   "Pueto Laboral"
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
         TabIndex        =   18
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label Label1 
         Caption         =   "Descripción"
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
         Height          =   495
         Left            =   120
         TabIndex        =   16
         Top             =   960
         Width           =   1455
      End
   End
End
Attribute VB_Name = "CargaDeduccionesGenerales"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()

    Call CenterMe(CargaDeduccionesGenerales, 6950, 5000)

End Sub

Private Sub cmdAgregar_Click()

    GenerarDeduccionesGenerales

End Sub

Private Sub Form_Unload(Cancel As Integer)

    strEditandoDeduccionesGenerales = ""

End Sub
