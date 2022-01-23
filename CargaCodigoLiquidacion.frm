VERSION 5.00
Begin VB.Form CargaCodigoLiquidacion 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Liquidaciones"
   ClientHeight    =   3435
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   5925
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3435
   ScaleWidth      =   5925
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Caption         =   "Incremento Deduccion Especial (Exenciones)"
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
      TabIndex        =   8
      Top             =   1800
      Width           =   5655
      Begin VB.TextBox txtMontoExento 
         Height          =   285
         Left            =   1440
         TabIndex        =   9
         Text            =   "0"
         Top             =   480
         Width           =   1000
      End
      Begin VB.Label Label4 
         Caption         =   "Máx. Exento"
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
         TabIndex        =   10
         Top             =   480
         Width           =   1455
      End
   End
   Begin VB.CommandButton cmdAgregar 
      BackColor       =   &H008080FF&
      Caption         =   "Guardar Datos"
      Height          =   495
      Left            =   2040
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   2880
      Width           =   1815
   End
   Begin VB.Frame Frame3 
      Caption         =   "Datos de la Liquidación"
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
      Height          =   1575
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5655
      Begin VB.TextBox txtCodigo 
         Height          =   285
         Left            =   1440
         TabIndex        =   1
         Top             =   480
         Width           =   1000
      End
      Begin VB.TextBox txtDescripcion 
         Height          =   285
         Left            =   1440
         TabIndex        =   3
         Top             =   1080
         Width           =   4000
      End
      Begin VB.TextBox txtPeriodo 
         Height          =   285
         Left            =   4440
         TabIndex        =   2
         Top             =   480
         Width           =   1000
      End
      Begin VB.Label Label1 
         Caption         =   "Código"
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
         Top             =   480
         Width           =   1455
      End
      Begin VB.Label Label3 
         Caption         =   "Descripción"
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
         Top             =   1080
         Width           =   1455
      End
      Begin VB.Label Label2 
         Caption         =   "Período (mm/aaaa)"
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
         Left            =   3120
         TabIndex        =   5
         Top             =   480
         Width           =   1455
      End
   End
End
Attribute VB_Name = "CargaCodigoLiquidacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAgregar_Click()

    GenerarCodigoLiquidacion

End Sub

Private Sub Form_Load()

    Call CenterMe(CargaCodigoLiquidacion, 6000, 3900)

End Sub

Private Sub Form_Unload(Cancel As Integer)

    strEditandoCodigoLiquidacion = ""

End Sub

