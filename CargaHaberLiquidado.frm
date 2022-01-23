VERSION 5.00
Begin VB.Form CargaHaberLiquidado 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Agentes"
   ClientHeight    =   4665
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   6855
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4665
   ScaleWidth      =   6855
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Caption         =   "Datos del Período y Concepto a Incorporar"
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
      TabIndex        =   8
      Top             =   1800
      Width           =   6615
      Begin VB.TextBox txtPeriodo 
         Height          =   285
         Left            =   1440
         TabIndex        =   11
         Top             =   480
         Width           =   1575
      End
      Begin VB.ComboBox cmbConcepto 
         Height          =   315
         Left            =   1440
         TabIndex        =   1
         Top             =   1080
         Width           =   5055
      End
      Begin VB.TextBox txtImporte 
         Height          =   285
         Left            =   1440
         TabIndex        =   2
         Top             =   1680
         Width           =   1575
      End
      Begin VB.Label Label6 
         Caption         =   "Período"
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
         TabIndex        =   12
         Top             =   480
         Width           =   1455
      End
      Begin VB.Label Label7 
         Caption         =   "Concepto"
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
         Top             =   1080
         Width           =   1455
      End
      Begin VB.Label Label4 
         Caption         =   "Importe"
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
         TabIndex        =   9
         Top             =   1680
         Width           =   1455
      End
   End
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
      Height          =   1575
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6615
      Begin VB.TextBox txtPuestoLaboral 
         Height          =   285
         Left            =   1440
         TabIndex        =   4
         Top             =   480
         Width           =   1575
      End
      Begin VB.TextBox txtNombreCompleto 
         Height          =   285
         Left            =   1440
         TabIndex        =   6
         Top             =   1080
         Width           =   5055
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
         Caption         =   "Nombre"
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
         TabIndex        =   5
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
      TabIndex        =   3
      Top             =   4080
      Width           =   1815
   End
End
Attribute VB_Name = "CargaHaberLiquidado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()

    Call CenterMe(CargaHaberLiquidado, 6950, 5100)
    Me.txtPuestoLaboral.Enabled = False
    Me.txtNombreCompleto.Enabled = False
    Me.txtPeriodo.Enabled = False
    CargarCmbCargaHaberLiquidado

End Sub

Private Sub cmdAgregar_Click()

    GenerarHaberLiquidado

End Sub
