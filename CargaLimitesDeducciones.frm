VERSION 5.00
Begin VB.Form CargaLimitesDeducciones 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Límites Deducciones 4ta Categoría Ganancias"
   ClientHeight    =   7395
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   6855
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7395
   ScaleWidth      =   6855
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame4 
      Caption         =   "Deducciones Generales (Porcentaje s/ Gcia. Neta)"
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
      TabIndex        =   26
      Top             =   5160
      Width           =   6615
      Begin VB.TextBox txtDonaciones 
         Height          =   285
         Left            =   1440
         TabIndex        =   13
         Top             =   1080
         Width           =   1575
      End
      Begin VB.TextBox txtCuotaMedico 
         Height          =   285
         Left            =   4920
         TabIndex        =   12
         Top             =   480
         Width           =   1575
      End
      Begin VB.TextBox txtHonorariosMedicos 
         Height          =   285
         Left            =   1440
         TabIndex        =   11
         Top             =   480
         Width           =   1575
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
         TabIndex        =   29
         Top             =   1080
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
         TabIndex        =   28
         Top             =   480
         Width           =   1455
      End
      Begin VB.Label Label10 
         Caption         =   "Honorarios Médicos"
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
         TabIndex        =   27
         Top             =   480
         Width           =   1455
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Deducciones Generales (Importe Fijo Anual)"
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
      TabIndex        =   23
      Top             =   3480
      Width           =   6615
      Begin VB.TextBox txtAlquileres 
         Height          =   285
         Left            =   1440
         TabIndex        =   10
         Top             =   1080
         Width           =   1575
      End
      Begin VB.TextBox txtServicioDomestico 
         Height          =   285
         Left            =   1440
         TabIndex        =   8
         Top             =   480
         Width           =   1575
      End
      Begin VB.TextBox txtSeguroDeVida 
         Height          =   285
         Left            =   4920
         TabIndex        =   9
         Top             =   480
         Width           =   1575
      End
      Begin VB.Label Label13 
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
         TabIndex        =   30
         Top             =   1080
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
         TabIndex        =   25
         Top             =   480
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
         TabIndex        =   24
         Top             =   480
         Width           =   1455
      End
   End
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
      TabIndex        =   20
      Top             =   120
      Width           =   6615
      Begin VB.TextBox txtFecha 
         Height          =   285
         Left            =   4920
         TabIndex        =   2
         Top             =   480
         Width           =   1575
      End
      Begin VB.TextBox txtNormaLegal 
         Height          =   285
         Left            =   1440
         TabIndex        =   1
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
         TabIndex        =   22
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
         TabIndex        =   21
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
      TabIndex        =   14
      Top             =   6840
      Width           =   1815
   End
   Begin VB.Frame Frame3 
      Caption         =   "Deducciones Personales (Importe Fijo Anual)"
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
      Top             =   1200
      Width           =   6615
      Begin VB.TextBox txtOtrasCargas 
         Height          =   285
         Left            =   1440
         TabIndex        =   7
         Top             =   1680
         Width           =   1575
      End
      Begin VB.TextBox txtHijo 
         Height          =   285
         Left            =   1440
         TabIndex        =   5
         Top             =   1080
         Width           =   1575
      End
      Begin VB.TextBox txtMinimoNoImponible 
         Height          =   285
         Left            =   1440
         TabIndex        =   3
         Top             =   480
         Width           =   1575
      End
      Begin VB.TextBox txtDeduccionEspecial 
         Height          =   285
         Left            =   4920
         TabIndex        =   4
         Top             =   480
         Width           =   1575
      End
      Begin VB.TextBox txtConyuge 
         Height          =   285
         Left            =   4920
         TabIndex        =   6
         Top             =   1080
         Width           =   1575
      End
      Begin VB.Label Label6 
         Caption         =   "Otras Cargas"
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
         TabIndex        =   19
         Top             =   1680
         Width           =   1455
      End
      Begin VB.Label Label3 
         Caption         =   "Hijo/a"
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
         Top             =   1080
         Width           =   1455
      End
      Begin VB.Label Label1 
         Caption         =   "Mínimo no Imponible"
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
         TabIndex        =   17
         Top             =   480
         Width           =   1455
      End
      Begin VB.Label Label2 
         Caption         =   "Deducción Especial"
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
         TabIndex        =   16
         Top             =   480
         Width           =   1455
      End
      Begin VB.Label Label4 
         Caption         =   "Conyuge"
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
         TabIndex        =   15
         Top             =   1080
         Width           =   1455
      End
   End
End
Attribute VB_Name = "CargaLimitesDeducciones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()

    Call CenterMe(CargaLimitesDeducciones, 6950, 7750)

End Sub

Private Sub cmdAgregar_Click()

    GenerarLimitesDeducciones

End Sub

Private Sub Form_Unload(Cancel As Integer)

    strEditandoLimitesDeducciones = ""

End Sub

