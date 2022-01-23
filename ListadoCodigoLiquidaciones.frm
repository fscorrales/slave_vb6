VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form ListadoCodigoLiquidaciones 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Código Liquidaciones"
   ClientHeight    =   7275
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   5655
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7275
   ScaleWidth      =   5655
   ShowInTaskbar   =   0   'False
   Begin VB.Frame frmMultiuso 
      Caption         =   "Códigos Liquidaciones"
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
      Width           =   5445
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid dgCodigoLiquidacion 
         Height          =   4455
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   5175
         _ExtentX        =   9128
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
      Left            =   1560
      TabIndex        =   0
      Top             =   5160
      Width           =   2415
      Begin VB.CommandButton cmdAgregar 
         BackColor       =   &H008080FF&
         Caption         =   "Agregar Código"
         Height          =   375
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   360
         Width           =   2175
      End
      Begin VB.CommandButton cmdEditar 
         BackColor       =   &H008080FF&
         Caption         =   "Editar Código"
         Height          =   375
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   960
         Width           =   2175
      End
      Begin VB.CommandButton cmdEliminar 
         BackColor       =   &H008080FF&
         Caption         =   "Eliminar Código"
         Height          =   375
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   1560
         Width           =   2175
      End
   End
End
Attribute VB_Name = "ListadoCodigoLiquidaciones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()

    Call CenterMe(ListadoCodigoLiquidaciones, 5750, 7650)

End Sub

Private Sub cmdAgregar_Click()

    CargaCodigoLiquidacion.Show
    Unload ListadoCodigoLiquidaciones

End Sub

Private Sub cmdEditar_Click()

    EditarCodigoLiquidacion

End Sub

Private Sub cmdEliminar_Click()

    EliminarCodigoLiquidacion

End Sub
