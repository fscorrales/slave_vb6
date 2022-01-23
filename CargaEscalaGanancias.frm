VERSION 5.00
Begin VB.Form CargaEscalaGanancias 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Tramos de Escala (Art.90) - Gcias. 4ta Categoría"
   ClientHeight    =   2970
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   6885
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2970
   ScaleWidth      =   6885
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdAgregar 
      BackColor       =   &H008080FF&
      Caption         =   "Guardar Datos"
      Height          =   495
      Left            =   2520
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   2400
      Width           =   1815
   End
   Begin VB.Frame Frame3 
      Caption         =   "Importes Acumulados a Diciembre por Tramo"
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
      Begin VB.TextBox txtFecha 
         Height          =   285
         Left            =   4920
         TabIndex        =   11
         Top             =   480
         Width           =   1575
      End
      Begin VB.TextBox txtNormaLegal 
         Height          =   285
         Left            =   1440
         TabIndex        =   10
         Top             =   480
         Width           =   1575
      End
      Begin VB.TextBox txtImporteVariable 
         Height          =   285
         Left            =   4920
         TabIndex        =   4
         Top             =   1680
         Width           =   1575
      End
      Begin VB.TextBox txtTramo 
         Height          =   285
         Left            =   1440
         TabIndex        =   1
         Top             =   1080
         Width           =   1575
      End
      Begin VB.TextBox txtImporteMaximo 
         Height          =   285
         Left            =   4920
         TabIndex        =   2
         Top             =   1080
         Width           =   1575
      End
      Begin VB.TextBox txtImporteFijo 
         Height          =   285
         Left            =   1440
         TabIndex        =   3
         Top             =   1680
         Width           =   1575
      End
      Begin VB.Label Label6 
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
         TabIndex        =   13
         Top             =   480
         Width           =   1455
      End
      Begin VB.Label Label5 
         Caption         =   "Norma Legal"
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
      Begin VB.Label Label3 
         Caption         =   "% Variable"
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
         Top             =   1680
         Width           =   1455
      End
      Begin VB.Label Label1 
         Caption         =   "Nro. Tramo"
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
         TabIndex        =   8
         Top             =   1080
         Width           =   1455
      End
      Begin VB.Label Label2 
         Caption         =   "Monto Max."
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
         TabIndex        =   7
         Top             =   1080
         Width           =   1455
      End
      Begin VB.Label Label4 
         Caption         =   "Monto Fijo"
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
         TabIndex        =   6
         Top             =   1680
         Width           =   1455
      End
   End
End
Attribute VB_Name = "CargaEscalaGanancias"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()

    Call CenterMe(CargaEscalaGanancias, 6950, 3400)

End Sub

Private Sub cmdAgregar_Click()

    GenerarEscalaGanancias
    
End Sub

Private Sub Form_Unload(Cancel As Integer)

    strEditandoEscalaGanancias = ""

End Sub

