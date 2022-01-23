VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form ListadoAgentes 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Agentes"
   ClientHeight    =   6225
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   7830
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6225
   ScaleWidth      =   7830
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Caption         =   "Listado de Agentes"
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
      TabIndex        =   4
      Top             =   120
      Width           =   7650
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid dgAgentes 
         Height          =   4455
         Left            =   120
         TabIndex        =   5
         Top             =   360
         Width           =   7400
         _ExtentX        =   13044
         _ExtentY        =   7858
         _Version        =   393216
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
   End
   Begin VB.Frame Frame3 
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
      Height          =   975
      Left            =   120
      TabIndex        =   0
      Top             =   5160
      Width           =   7650
      Begin VB.CommandButton cmdAgregar 
         BackColor       =   &H008080FF&
         Caption         =   "Agregar Agente"
         Height          =   375
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   360
         Width           =   2175
      End
      Begin VB.CommandButton cmdEditar 
         BackColor       =   &H008080FF&
         Caption         =   "Editar Agente"
         Height          =   375
         Left            =   2700
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   360
         Width           =   2175
      End
      Begin VB.CommandButton cmdEliminar 
         BackColor       =   &H008080FF&
         Caption         =   "Eliminar Agente"
         Height          =   375
         Left            =   5280
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   360
         Width           =   2175
      End
   End
End
Attribute VB_Name = "ListadoAgentes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAgregar_Click()

    CargaAgente.Show
    Unload ListadoAgentes

End Sub

Private Sub cmdEditar_Click()

    EditarAgente

End Sub

Private Sub cmdEliminar_Click()

    EliminarAgente

End Sub

Private Sub Form_Load()

    Call CenterMe(ListadoAgentes, 7950, 6550)

End Sub
