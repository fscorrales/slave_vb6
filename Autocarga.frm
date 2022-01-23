VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form Autocarga 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Asistente de Carga"
   ClientHeight    =   6015
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   11415
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6015
   ScaleWidth      =   11415
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Caption         =   "Listado de Comprobantes a Cargar en SIIF"
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
      Height          =   5775
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   11175
      Begin VB.CommandButton cmdAgregar 
         BackColor       =   &H008080FF&
         Caption         =   "Cargar Comprobante SIIF"
         Height          =   495
         Left            =   2400
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   5160
         Width           =   1695
      End
      Begin VB.CommandButton cmdEliminar 
         BackColor       =   &H008080FF&
         Caption         =   "Eliminar Comprobante"
         Height          =   495
         Left            =   6840
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   5160
         Width           =   1695
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid dgListadoAutocarga 
         Height          =   4815
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   10935
         _ExtentX        =   19288
         _ExtentY        =   8493
         _Version        =   393216
         BandDisplay     =   1
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
   End
End
Attribute VB_Name = "Autocarga"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAgregar_Click()

    EditarAutocarga
    
End Sub

Private Sub cmdEliminar_Click()

    EliminarAutocarga

End Sub


Private Sub Form_Load()

    Call CenterMe(Autocarga, 11500, 6350)

End Sub

'Private Sub dgAutocargaCertificado_KeyDown(KeyCode As Integer, Shift As Integer)
'
'    If KeyCode = vbKeyF6 Then
'        LlenarVariablesAutocargaCertificado
'        AgregarCertificado
'    End If
'    If KeyCode = vbKeyF9 Then
'        EliminarCertificado
'    End If
'
'End Sub
'
'Private Sub Form_Unload(Cancel As Integer)
'    'dgCertificado.Clear
'    'dgEPAM.Clear
'End Sub
'
