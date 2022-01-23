VERSION 5.00
Begin VB.Form CargaFamiliar 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Carga de Familiar"
   ClientHeight    =   6435
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   6870
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6435
   ScaleWidth      =   6870
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame2 
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
      Height          =   1575
      Left            =   120
      TabIndex        =   15
      Top             =   120
      Width           =   6615
      Begin VB.TextBox txtDescripcionAgente 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1440
         TabIndex        =   0
         Top             =   960
         Width           =   5055
      End
      Begin VB.TextBox txtPuestoLaboral 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1440
         TabIndex        =   1
         Top             =   360
         Width           =   1575
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
         Height          =   375
         Left            =   120
         TabIndex        =   17
         Top             =   960
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
         TabIndex        =   16
         Top             =   360
         Width           =   1455
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Datos del Familiar"
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
      Height          =   3975
      Left            =   120
      TabIndex        =   11
      Top             =   1800
      Width           =   6615
      Begin VB.CheckBox chkGanancias 
         Alignment       =   1  'Right Justify
         Caption         =   "Deducible Ganancias"
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
         TabIndex        =   10
         Top             =   3360
         Width           =   2350
      End
      Begin VB.CheckBox chkAdherente 
         Alignment       =   1  'Right Justify
         Caption         =   "Adherente Obra Social"
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
         TabIndex        =   8
         Top             =   2760
         Width           =   2350
      End
      Begin VB.CheckBox chkDiscapacitado 
         Alignment       =   1  'Right Justify
         Caption         =   "Discapacitado"
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
         Left            =   120
         TabIndex        =   9
         Top             =   3360
         Width           =   2350
      End
      Begin VB.CheckBox chkCobraSalario 
         Alignment       =   1  'Right Justify
         Caption         =   "Cobra Salario"
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
         Left            =   120
         TabIndex        =   7
         Top             =   2760
         Width           =   2350
      End
      Begin VB.ComboBox cmbNivelDeEstudio 
         Height          =   315
         Left            =   4920
         TabIndex        =   6
         Top             =   2160
         Width           =   1575
      End
      Begin VB.ComboBox cmbParentesco 
         Height          =   315
         Left            =   1440
         TabIndex        =   4
         Top             =   1560
         Width           =   5055
      End
      Begin VB.TextBox txtDescripcionFamiliar 
         Height          =   285
         Left            =   1440
         TabIndex        =   3
         Top             =   960
         Width           =   5055
      End
      Begin VB.TextBox txtDNI 
         Height          =   285
         Left            =   1440
         TabIndex        =   2
         Top             =   360
         Width           =   1575
      End
      Begin VB.TextBox txtFechaAlta 
         Height          =   285
         Left            =   1440
         TabIndex        =   5
         Top             =   2160
         Width           =   1575
      End
      Begin VB.Label Label5 
         Caption         =   "Nivel Estudio"
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
         TabIndex        =   20
         Top             =   2160
         Width           =   1455
      End
      Begin VB.Label Label4 
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
         Height          =   375
         Left            =   120
         TabIndex        =   19
         Top             =   960
         Width           =   1455
      End
      Begin VB.Label Label3 
         Caption         =   "DNI"
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
      Begin VB.Label Label12 
         Caption         =   "Fecha Alta"
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
         TabIndex        =   14
         Top             =   2160
         Width           =   1455
      End
      Begin VB.Label Label9 
         Caption         =   "Parentesco"
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
         Top             =   1560
         Width           =   1455
      End
   End
   Begin VB.CommandButton cmdAgregar 
      BackColor       =   &H008080FF&
      Caption         =   "Guardar Datos"
      Height          =   495
      Left            =   2520
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   5880
      Width           =   1815
   End
End
Attribute VB_Name = "CargaFamiliar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()

    Call CenterMe(CargaFamiliar, 6950, 6800)

End Sub

Private Sub cmdAgregar_Click()

    GenerarFamiliar

End Sub

Private Sub Form_Unload(Cancel As Integer)

    strEditandoFamiliar = ""

End Sub
