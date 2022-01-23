VERSION 5.00
Begin VB.Form CargaDeduccionesPersonales 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Deducciones Personales"
   ClientHeight    =   7440
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   6870
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7440
   ScaleWidth      =   6870
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Caption         =   "Deducciones Personales a partir de Cargas de Familia"
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
      Height          =   4935
      Left            =   120
      TabIndex        =   10
      Top             =   1800
      Width           =   6615
      Begin VB.ListBox lstFamiliaresDeduciblesGanancias 
         Height          =   3570
         ItemData        =   "CargaDeduccionesPersonales.frx":0000
         Left            =   3480
         List            =   "CargaDeduccionesPersonales.frx":0002
         TabIndex        =   3
         Top             =   720
         Width           =   3000
      End
      Begin VB.ListBox lstFamiliaresACargo 
         Height          =   3570
         ItemData        =   "CargaDeduccionesPersonales.frx":0004
         Left            =   120
         List            =   "CargaDeduccionesPersonales.frx":0006
         TabIndex        =   1
         Top             =   720
         Width           =   3000
      End
      Begin VB.CommandButton cmdAgregar 
         BackColor       =   &H008080FF&
         Caption         =   "Agregar Familiar Deducible"
         Height          =   375
         Left            =   480
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   4440
         Width           =   2175
      End
      Begin VB.CommandButton cmdEliminar 
         BackColor       =   &H008080FF&
         Caption         =   "Qutar Familiar Deducible"
         Height          =   375
         Left            =   3960
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   4440
         Width           =   2175
      End
      Begin VB.Label Label5 
         Caption         =   "Familiares a Cargo"
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
         Top             =   360
         Width           =   2895
      End
      Begin VB.Label Label3 
         Caption         =   "Familiares Deducibles Ganancias"
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
         Left            =   3480
         TabIndex        =   11
         Top             =   360
         Width           =   2895
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
      TabIndex        =   5
      Top             =   120
      Width           =   6615
      Begin VB.TextBox txtDescripcion 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1440
         TabIndex        =   7
         Top             =   960
         Width           =   5055
      End
      Begin VB.TextBox txtPuestoLaboral 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1440
         TabIndex        =   6
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
         Height          =   495
         Left            =   120
         TabIndex        =   9
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
         TabIndex        =   8
         Top             =   360
         Width           =   1455
      End
   End
   Begin VB.CommandButton cmdGuardar 
      BackColor       =   &H008080FF&
      Caption         =   "Guardar Datos"
      Height          =   495
      Left            =   2520
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   6840
      Width           =   1815
   End
End
Attribute VB_Name = "CargaDeduccionesPersonales"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()

    Call CenterMe(CargaDeduccionesPersonales, 6950, 7800)

End Sub

Private Sub cmdGuardar_Click()

    If Not Me.lstFamiliaresDeduciblesGanancias.ListCount = "0" Then
        GenerarDeduccionesPersonales
    End If

End Sub

Private Sub Form_Unload(Cancel As Integer)

    'strEditandoDeduccionesGenerales = ""

End Sub

Private Sub cmdAgregar_Click()
    
    With CargaDeduccionesPersonales
        Call PasarDatosListBox(.lstFamiliaresACargo, .lstFamiliaresDeduciblesGanancias)
    End With
    
End Sub

Private Sub cmdEliminar_Click()

    With CargaDeduccionesPersonales
        Call PasarDatosListBox(.lstFamiliaresDeduciblesGanancias, .lstFamiliaresACargo)
    End With
End Sub
