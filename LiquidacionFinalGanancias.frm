VERSION 5.00
Begin VB.Form LiquidacionFinalGanancias 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Liquidación Anual / Final"
   ClientHeight    =   2970
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   6870
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2970
   ScaleWidth      =   6870
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame3 
      Caption         =   "Datos del Agente"
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
      TabIndex        =   3
      Top             =   120
      Width           =   6615
      Begin VB.OptionButton optLiquidacionFinal 
         Caption         =   "Liquidación Final"
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
         Width           =   2295
      End
      Begin VB.OptionButton optLiquidacionAnual 
         Caption         =   "Liquidación Anual"
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
         Left            =   720
         TabIndex        =   8
         Top             =   1680
         Value           =   -1  'True
         Width           =   2295
      End
      Begin VB.TextBox txtDescripcion 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1440
         TabIndex        =   2
         Top             =   1080
         Width           =   5055
      End
      Begin VB.TextBox txtPuestoLaboral 
         Height          =   285
         Left            =   1440
         TabIndex        =   0
         Top             =   480
         Width           =   1575
      End
      Begin VB.TextBox txtPeriodo 
         Height          =   285
         Left            =   4920
         TabIndex        =   1
         Top             =   480
         Width           =   1575
      End
      Begin VB.Label Label3 
         Caption         =   "Descripcion"
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
         TabIndex        =   7
         Top             =   1080
         Width           =   1455
      End
      Begin VB.Label Label1 
         Caption         =   "Puesto Laboral"
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
         Top             =   480
         Width           =   1455
      End
      Begin VB.Label Año 
         Caption         =   "Año (aaaa)"
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
         TabIndex        =   5
         Top             =   480
         Width           =   1455
      End
   End
   Begin VB.CommandButton cmdAgregar 
      BackColor       =   &H008080FF&
      Caption         =   "Generar F. 649"
      Height          =   495
      Left            =   2520
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   2400
      Width           =   1815
   End
End
Attribute VB_Name = "LiquidacionFinalGanancias"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()

    Call CenterMe(LiquidacionFinalGanancias, 6950, 3250)

End Sub

Private Sub cmdAgregar_Click()

    Call GenerarF649(App.Path & "\F649.xls")

End Sub

Private Sub txtPuestoLaboral_LostFocus()

    Dim SQL As String

    SQL = "Select * From AGENTES Where PUESTOLABORAL = '" & Me.txtPuestoLaboral.Text & "'"
    If SQLNoMatch(SQL) = True Then
        MsgBox "Debe ingresar un Nro de Puesto Laboral válido", vbCritical + vbOKOnly, "NRO PUESTO LABORAL INEXISTENTE"
        Me.txtDescripcion.Text = ""
        Me.txtPuestoLaboral.SetFocus
    Else
        Set rstBuscarSlave = New ADODB.Recordset
        rstBuscarSlave.Open SQL, dbSlave, adOpenDynamic, adLockOptimistic
        Me.txtDescripcion.Text = rstBuscarSlave!NombreCompleto
        rstBuscarSlave.Close
        Set rstBuscarSlave = Nothing
    End If
    
End Sub


