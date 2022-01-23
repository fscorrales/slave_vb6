VERSION 5.00
Begin VB.Form CargaComprobanteSIIF 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Carga Gasto"
   ClientHeight    =   4920
   ClientLeft      =   4755
   ClientTop       =   285
   ClientWidth     =   6885
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4920
   ScaleWidth      =   6885
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Caption         =   "Número de Comprobante y Ejercicio de Imputación"
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
      TabIndex        =   17
      Top             =   120
      Width           =   6615
      Begin VB.TextBox txtFecha 
         Height          =   285
         Left            =   4920
         TabIndex        =   1
         Top             =   360
         Width           =   1575
      End
      Begin VB.TextBox txtComprobante 
         Height          =   285
         Left            =   1560
         TabIndex        =   0
         Top             =   360
         Width           =   1575
      End
      Begin VB.Label Label3 
         Caption         =   "Fecha (dd/mm/aaaa)"
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
         Left            =   3480
         TabIndex        =   19
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label Label1 
         Caption         =   "Comprobante"
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
         TabIndex        =   18
         Top             =   360
         Width           =   1455
      End
   End
   Begin VB.CommandButton cmdAgregar 
      BackColor       =   &H008080FF&
      Caption         =   "Guardar Datos"
      Height          =   495
      Left            =   2520
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   4320
      Width           =   1815
   End
   Begin VB.Frame Frame3 
      Caption         =   "Datos Numéricos del Comprobante"
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
      TabIndex        =   12
      Top             =   2160
      Width           =   6615
      Begin VB.TextBox txtFuente 
         Height          =   285
         Left            =   1560
         TabIndex        =   7
         Top             =   1560
         Width           =   1575
      End
      Begin VB.TextBox txtCuenta 
         Height          =   285
         Left            =   1560
         TabIndex        =   5
         Top             =   960
         Width           =   1575
      End
      Begin VB.TextBox txtRetenciones 
         Height          =   285
         Left            =   4920
         TabIndex        =   4
         Top             =   360
         Width           =   1575
      End
      Begin VB.TextBox txtImporte 
         Height          =   285
         Left            =   1560
         TabIndex        =   3
         Top             =   360
         Width           =   1575
      End
      Begin VB.TextBox txtDescripcionCuenta 
         Height          =   285
         Left            =   3240
         TabIndex        =   6
         Top             =   960
         Width           =   3255
      End
      Begin VB.TextBox txtDescripcionFuente 
         Height          =   285
         Left            =   3240
         TabIndex        =   8
         Top             =   1560
         Width           =   3255
      End
      Begin VB.Label Label5 
         Caption         =   "Importe Bruto"
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
         TabIndex        =   16
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label Label7 
         Caption         =   "Retenciones"
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
         TabIndex        =   15
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label Label9 
         Caption         =   "Cuenta"
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
         TabIndex        =   14
         Top             =   960
         Width           =   1455
      End
      Begin VB.Label Label10 
         Caption         =   "Fuente"
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
         TabIndex        =   13
         Top             =   1560
         Width           =   1455
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Tipo de Imputación"
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
      Height          =   855
      Left            =   120
      TabIndex        =   9
      Top             =   1200
      Width           =   6615
      Begin VB.ComboBox txtImputacion 
         Height          =   315
         Left            =   1560
         TabIndex        =   2
         Top             =   360
         Width           =   4935
      End
      Begin VB.Label Label4 
         Caption         =   "Imputación"
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
         TabIndex        =   11
         Top             =   360
         Width           =   1455
      End
   End
End
Attribute VB_Name = "CargaComprobanteSIIF"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAgregar_Click()

    GenerarComprobanteSIIF

End Sub

Private Sub Form_Load()

    Call CenterMe(CargaComprobanteSIIF, 7000, 5350)
    CargarcmbComprobanteSIIF
    
    With Me
        .txtImporte.Enabled = False
        .txtRetenciones.Enabled = False
        .txtCuenta.Enabled = False
        .txtFuente.Enabled = False
        .txtDescripcionCuenta.Enabled = False
        .txtDescripcionFuente.Enabled = False
    End With

End Sub

Private Sub Form_Unload(Cancel As Integer)

'    If bolAutocargaCertificado = True Then
'        VaciarVariablesAutocarga
'        AutocargaCertificados.Show
'        ConfigurardgAutocargaCertificados
'        CargardgAutocargaCertificados
'    ElseIf bolAutocargaEPAM = True Then
'        VaciarVariablesAutocarga
'        AutocargaEPAM.Show
'        ConfigurardgAutocargaEPAM
'        CargardgAutocargaEPAM
'    End If
    strEditandoAutocarga = ""

End Sub
