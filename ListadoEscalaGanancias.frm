VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form ListadoEscalaGanancias 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Tramos de Escala (Art.90) - Gcias. 4ta Categoría"
   ClientHeight    =   7185
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   8655
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7185
   ScaleWidth      =   8655
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame5 
      Caption         =   "Normas Aplicables"
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
      TabIndex        =   16
      Top             =   0
      Width           =   3090
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid dgNormasEscalaGanancias 
         Height          =   4455
         Left            =   120
         TabIndex        =   17
         Top             =   240
         Width           =   2750
         _ExtentX        =   4842
         _ExtentY        =   7858
         _Version        =   393216
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
   End
   Begin VB.Frame Frame4 
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
      Left            =   360
      TabIndex        =   12
      Top             =   5040
      Width           =   2415
      Begin VB.CommandButton cmdAgregarNorma 
         BackColor       =   &H008080FF&
         Caption         =   "Agregar Norma Legal"
         Height          =   375
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   360
         Width           =   2175
      End
      Begin VB.CommandButton cmdEditarNorma 
         BackColor       =   &H008080FF&
         Caption         =   "Editar Norma Legal"
         Height          =   375
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   960
         Width           =   2175
      End
      Begin VB.CommandButton cmdEliminarNorma 
         BackColor       =   &H008080FF&
         Caption         =   "Eliminar Norma Legal"
         Height          =   375
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   1560
         Width           =   2175
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Importes acumulados a Diciembre"
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
      TabIndex        =   10
      Top             =   0
      Width           =   15
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
         Height          =   4455
         Left            =   120
         TabIndex        =   11
         Top             =   240
         Width           =   4395
         _ExtentX        =   7752
         _ExtentY        =   7858
         _Version        =   393216
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
   End
   Begin VB.Frame Frame1 
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
      TabIndex        =   6
      Top             =   5040
      Width           =   15
      Begin VB.CommandButton Command3 
         BackColor       =   &H008080FF&
         Caption         =   "Agregar Tramo Escala"
         Height          =   375
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   360
         Width           =   2175
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H008080FF&
         Caption         =   "Editar Tramo Escala"
         Height          =   375
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   960
         Width           =   2175
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H008080FF&
         Caption         =   "Eliminar Tramo Escala"
         Height          =   375
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   1560
         Width           =   2175
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
      Left            =   4680
      TabIndex        =   2
      Top             =   5040
      Width           =   2415
      Begin VB.CommandButton cmdEliminar 
         BackColor       =   &H008080FF&
         Caption         =   "Eliminar Tramo Escala"
         Height          =   375
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   1560
         Width           =   2175
      End
      Begin VB.CommandButton cmdEditar 
         BackColor       =   &H008080FF&
         Caption         =   "Editar Tramo Escala"
         Height          =   375
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   960
         Width           =   2175
      End
      Begin VB.CommandButton cmdAgregar 
         BackColor       =   &H008080FF&
         Caption         =   "Agregar Tramo Escala"
         Height          =   375
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   360
         Width           =   2175
      End
   End
   Begin VB.Frame frmMultiuso 
      Caption         =   "Importes acumulados a Diciembre"
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
      Left            =   3240
      TabIndex        =   0
      Top             =   0
      Width           =   5370
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid dgEscalaGanancias 
         Height          =   4455
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   5000
         _ExtentX        =   8811
         _ExtentY        =   7858
         _Version        =   393216
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
   End
End
Attribute VB_Name = "ListadoEscalaGanancias"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAgregarNorma_Click()

    CargaNormaEscalaGanancias.Show
    Unload ListadoEscalaGanancias

End Sub

Private Sub cmdEditarNorma_Click()

    EditarNormaEscalaGanancias
    
End Sub

Private Sub cmdEliminarNorma_Click()

    EliminarNormaEscalaGanancias

End Sub

Private Sub Form_Load()

    Call CenterMe(ListadoEscalaGanancias, 8750, 7550)

End Sub

Private Sub cmdAgregar_Click()

    Dim strNormaLegal As String
    Dim strFecha As String
    
    With Me.dgNormasEscalaGanancias
        i = .Row
        strNormaLegal = .TextMatrix(i, 0)
        strFecha = .TextMatrix(i, 1)
    End With
            
    CargaEscalaGanancias.Show
    CargaEscalaGanancias.txtNormaLegal.Text = strNormaLegal
    CargaEscalaGanancias.txtNormaLegal.Enabled = False
    CargaEscalaGanancias.txtFecha.Text = strFecha
    CargaEscalaGanancias.txtFecha.Enabled = False
    Unload ListadoEscalaGanancias

End Sub

Private Sub cmdEditar_Click()

    EditarEscalaGanancias

End Sub

Private Sub cmdEliminar_Click()

    EliminarEscalaGanancias

End Sub

Private Sub dgNormasEscalaGanancias_RowColChange()

    Dim i As Integer
    i = Me.dgNormasEscalaGanancias.Row
    ConfigurardgEscalaGanancias
    Call CargardgEscalaGanancias(Me.dgNormasEscalaGanancias.TextMatrix(i, 0))
    i = 0

End Sub
