VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form ListadoConceptos 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Conceptos"
   ClientHeight    =   7185
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4665
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7185
   ScaleWidth      =   4665
   ShowInTaskbar   =   0   'False
   Begin VB.Frame frmMultiuso 
      Caption         =   "Datos del Concepto de Liquidación"
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
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   4649
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid dgConceptos 
         Height          =   4455
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   4395
         _ExtentX        =   7752
         _ExtentY        =   7858
         _Version        =   393216
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Acciones Posibles"
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
      Height          =   2055
      Left            =   1080
      TabIndex        =   0
      Top             =   5040
      Width           =   2415
      Begin VB.CommandButton cmdAgregar 
         BackColor       =   &H008080FF&
         Caption         =   "Agregar Concepto"
         Height          =   375
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   360
         Width           =   2175
      End
      Begin VB.CommandButton cmdEditar 
         BackColor       =   &H008080FF&
         Caption         =   "Editar Concepto"
         Height          =   375
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   960
         Width           =   2175
      End
      Begin VB.CommandButton cmdEliminar 
         BackColor       =   &H008080FF&
         Caption         =   "Eliminar Concepto"
         Height          =   375
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   1560
         Width           =   2175
      End
   End
End
Attribute VB_Name = "ListadoConceptos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()

    Call CenterMe(ListadoConceptos, 4750, 7500)

End Sub

Private Sub cmdAgregar_Click()

    CargaConcepto.Show
    Unload ListadoConceptos

End Sub

Private Sub cmdEditar_Click()

    EditarConcepto

End Sub

Private Sub cmdEliminar_Click()

    EliminarConcepto

End Sub
