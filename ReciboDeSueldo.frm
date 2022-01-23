VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form ReciboDeSueldo 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Recibo de Sueldo SISPER por Agente y Período"
   ClientHeight    =   6465
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   12840
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6465
   ScaleWidth      =   12840
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame2 
      Caption         =   "Identificación del Agente y Período"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   855
      Left            =   120
      TabIndex        =   5
      Top             =   0
      Width           =   12615
      Begin VB.CommandButton cmdActualizar 
         BackColor       =   &H008080FF&
         Caption         =   "Actualizar"
         Height          =   495
         Left            =   10440
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   240
         Width           =   1695
      End
      Begin VB.ComboBox cmbAgente 
         Height          =   315
         Left            =   1560
         TabIndex        =   1
         Top             =   360
         Width           =   3615
      End
      Begin VB.ComboBox cmbPeriodo 
         Height          =   315
         Left            =   7320
         TabIndex        =   2
         Top             =   360
         Width           =   2775
      End
      Begin VB.Label Label7 
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
         Left            =   5880
         TabIndex        =   7
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label Label8 
         Caption         =   "Agente"
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
         Top             =   360
         Width           =   1455
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Recibo de Sueldo del Agente y Período Seleccionado"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   5415
      Left            =   120
      TabIndex        =   0
      Top             =   960
      Width           =   12615
      Begin VB.TextBox txtSueldoNeto 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   300
         Left            =   7200
         TabIndex        =   15
         Top             =   4920
         Width           =   1605
      End
      Begin VB.CommandButton cmdEliminarHaber 
         BackColor       =   &H008080FF&
         Caption         =   "Eliminar Concepto/Haber"
         Height          =   495
         Left            =   2160
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   4680
         Width           =   1695
      End
      Begin VB.CommandButton cmdAgregarHaber 
         BackColor       =   &H008080FF&
         Caption         =   "Agregar Concepto/Haber"
         Height          =   495
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   4680
         Width           =   1695
      End
      Begin VB.TextBox txtTotalDescuentos 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   300
         Left            =   10440
         TabIndex        =   10
         Top             =   4920
         Width           =   1605
      End
      Begin VB.TextBox txtTotalHaberes 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   300
         Left            =   4080
         TabIndex        =   9
         Top             =   4920
         Width           =   1605
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid dgHaberesLiquidados 
         Height          =   4215
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   6015
         _ExtentX        =   10610
         _ExtentY        =   7435
         _Version        =   393216
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid dgDescuentosLiquidados 
         Height          =   4215
         Left            =   6480
         TabIndex        =   8
         Top             =   240
         Width           =   6015
         _ExtentX        =   10610
         _ExtentY        =   7435
         _Version        =   393216
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Caption         =   "Sueldo NETO"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000080FF&
         Height          =   255
         Left            =   7200
         TabIndex        =   16
         Top             =   4560
         Width           =   1575
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   "Total Descuentos"
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
         Left            =   10440
         TabIndex        =   12
         Top             =   4560
         Width           =   1575
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "Total Haberes"
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
         Left            =   4080
         TabIndex        =   11
         Top             =   4560
         Width           =   1575
      End
   End
End
Attribute VB_Name = "ReciboDeSueldo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdActualizar_Click()

    Dim strPuestoLaboral As String
    Dim strCodigoLiquidacion As String
    
    If ValidarActualizarReciboDeSueldo = True Then
        strPuestoLaboral = "Select PUESTOLABORAL From AGENTES" _
        & " Where NOMBRECOMPLETO = '" & Me.cmbAgente.Text & "'"
        Set rstBuscarSlave = New ADODB.Recordset
        rstBuscarSlave.Open strPuestoLaboral, dbSlave, adOpenDynamic, adLockOptimistic
        strPuestoLaboral = rstBuscarSlave!PuestoLaboral
        rstBuscarSlave.Close
        Set rstBuscarSlave = Nothing
        strCodigoLiquidacion = Left(Me.cmbPeriodo.Text, 4)
        Call CargardgHaberesLiquidados(strPuestoLaboral, strCodigoLiquidacion)
        Call CargardgDescuentosLiquidados(strPuestoLaboral, strCodigoLiquidacion)
    End If

End Sub

Private Sub cmdAgregarHaber_Click()
    
    Dim SQL As String
    
    With CargaHaberLiquidado
        .Show
        SQL = "Select PUESTOLABORAL From AGENTES" _
        & " Where NOMBRECOMPLETO = '" & Me.cmbAgente.Text & "'"
        Set rstBuscarSlave = New ADODB.Recordset
        rstBuscarSlave.Open SQL, dbSlave, adOpenDynamic, adLockOptimistic
        .txtPuestoLaboral.Text = rstBuscarSlave!PuestoLaboral
        rstBuscarSlave.Close
        Set rstBuscarSlave = Nothing
        .txtNombreCompleto.Text = Me.cmbAgente.Text
        .txtPeriodo.Text = Me.cmbPeriodo.Text
    End With
    
    Unload ReciboDeSueldo

End Sub

Private Sub cmdEliminarHaber_Click()

    EliminarHaberLiquidado

End Sub

Private Sub Form_Load()

    Call CenterMe(ReciboDeSueldo, 13000, 7000)
    CargarCmbReciboDeSueldo
    
End Sub

