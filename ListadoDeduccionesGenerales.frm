VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form ListadoDeduccionesGenerales 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Deducciones Generales"
   ClientHeight    =   6945
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   9060
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6945
   ScaleWidth      =   9060
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
      TabIndex        =   4
      Top             =   5880
      Width           =   8850
      Begin VB.CommandButton cmdEliminar 
         BackColor       =   &H008080FF&
         Caption         =   "Eliminar Deducción"
         Height          =   375
         Left            =   6480
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   360
         Width           =   2175
      End
      Begin VB.CommandButton cmdEditar 
         BackColor       =   &H008080FF&
         Caption         =   "Editar Deducción"
         Height          =   375
         Left            =   3360
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   360
         Width           =   2175
      End
      Begin VB.CommandButton cmdAgregar 
         BackColor       =   &H008080FF&
         Caption         =   "Agregar Deducción"
         Height          =   375
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   360
         Width           =   2175
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Listado de Deducciones Generales  Informadas por DDJJ"
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
      Left            =   120
      TabIndex        =   2
      Top             =   3720
      Width           =   8850
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid dgDeduccionesGenerales 
         Height          =   1575
         Left            =   120
         TabIndex        =   3
         Top             =   360
         Width           =   8595
         _ExtentX        =   15161
         _ExtentY        =   2778
         _Version        =   393216
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
   End
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
      Height          =   3495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   8850
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid dgAgentes 
         Height          =   3015
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   8595
         _ExtentX        =   15161
         _ExtentY        =   5318
         _Version        =   393216
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
   End
End
Attribute VB_Name = "ListadoDeduccionesGenerales"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAgregar_Click()

    Dim i As String

    With ListadoDeduccionesGenerales.dgAgentes
        i = .Row
        CargaDeduccionesGenerales.Show
        CargaDeduccionesGenerales.txtPuestoLaboral = .TextMatrix(i, 2)
        CargaDeduccionesGenerales.txtDescripcion = .TextMatrix(i, 1)
        i = ""
    End With
    Unload ListadoDeduccionesGenerales

End Sub

Private Sub cmdEditar_Click()

    EditarDeduccionesGenerales

End Sub

Private Sub cmdEliminar_Click()

    EliminarDeduccionesGenerales

End Sub

Private Sub dgAgentes_RowColChange()

    Dim i As Integer
    i = dgAgentes.Row
    ConfigurardgDeduccionesGenerales
    Call CargardgDeduccionesGenerales(dgAgentes.TextMatrix(i, 2))
    i = 0

End Sub

Private Sub Form_Load()

    Call CenterMe(ListadoDeduccionesGenerales, 9150, 7300)

End Sub
