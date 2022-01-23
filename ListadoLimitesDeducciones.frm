VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form ListadoLimitesDeducciones 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Límites Deducciones 4ta Categoría Ganancias"
   ClientHeight    =   6195
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   16515
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6195
   ScaleWidth      =   16515
   ShowInTaskbar   =   0   'False
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
      TabIndex        =   2
      Top             =   5160
      Width           =   16290
      Begin VB.CommandButton cmdEliminar 
         BackColor       =   &H008080FF&
         Caption         =   "Eliminar Norma Legal"
         Height          =   375
         Left            =   14040
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   360
         Width           =   2175
      End
      Begin VB.CommandButton cmdEditar 
         BackColor       =   &H008080FF&
         Caption         =   "Editar Norma Legal"
         Height          =   375
         Left            =   7080
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   360
         Width           =   2175
      End
      Begin VB.CommandButton cmdAgregar 
         BackColor       =   &H008080FF&
         Caption         =   "Agregar Norma Legal"
         Height          =   375
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   360
         Width           =   2175
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Listado de Deducciones Ordenados Por Norma Legal"
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
      TabIndex        =   0
      Top             =   120
      Width           =   16305
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid dgLimitesDeducciones 
         Height          =   4455
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   16065
         _ExtentX        =   28337
         _ExtentY        =   7858
         _Version        =   393216
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
   End
End
Attribute VB_Name = "ListadoLimitesDeducciones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAgregar_Click()

    CargaLimitesDeducciones.Show
    Unload ListadoLimitesDeducciones

End Sub

Private Sub cmdEditar_Click()

    EditarLimitesDeducciones

End Sub

Private Sub cmdEliminar_Click()

    EliminarLimitesDeducciones

End Sub

Private Sub Form_Load()

    Call CenterMe(ListadoLimitesDeducciones, 16600, 6550)

End Sub

