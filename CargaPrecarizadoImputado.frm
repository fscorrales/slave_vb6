VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form CargaPrecarizadoImputado 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Precarizado"
   ClientHeight    =   4890
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   6405
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4890
   ScaleWidth      =   6405
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame2 
      Caption         =   "Datos del Comprobante"
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
      TabIndex        =   10
      Top             =   0
      Width           =   6200
      Begin VB.TextBox txtComprobante 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1800
         TabIndex        =   12
         Top             =   360
         Width           =   1215
      End
      Begin VB.TextBox txtPeriodo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   4680
         TabIndex        =   11
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label Label9 
         Caption         =   "Nro CyO SIIF"
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
         Top             =   360
         Width           =   1815
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
         Left            =   3360
         TabIndex        =   13
         Top             =   360
         Width           =   1455
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Estructura y Monto Bruto Imputado"
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
      TabIndex        =   6
      Top             =   2640
      Width           =   6200
      Begin VB.TextBox txtMontoBruto 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1800
         TabIndex        =   15
         Top             =   1080
         Width           =   1215
      End
      Begin MSMask.MaskEdBox mskEstructuraImputada 
         Height          =   285
         Left            =   1800
         TabIndex        =   7
         Top             =   480
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   503
         _Version        =   393216
         MaxLength       =   12
         Mask            =   "##-##-##-###"
         PromptChar      =   "_"
      End
      Begin VB.Label Label8 
         Caption         =   "Monto Bruto"
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
         Top             =   1080
         Width           =   1815
      End
      Begin VB.Label Label6 
         Caption         =   "Estructura"
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
         TabIndex        =   9
         Top             =   480
         Width           =   1815
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         Caption         =   "Prog-Proy-Act-Partida"
         BeginProperty Font 
            Name            =   "Gill Sans MT"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000040C0&
         Height          =   255
         Left            =   3480
         TabIndex        =   8
         Top             =   480
         Width           =   2175
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Datos del Precarizado"
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
      Top             =   960
      Width           =   6200
      Begin VB.ComboBox cmbNombreCompleto 
         Height          =   315
         Left            =   1800
         TabIndex        =   17
         Top             =   480
         Width           =   4095
      End
      Begin MSMask.MaskEdBox mskEstructuraPrevista 
         Height          =   285
         Left            =   1800
         TabIndex        =   4
         Top             =   1080
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   503
         _Version        =   393216
         MaxLength       =   12
         Mask            =   "##-##-##-###"
         PromptChar      =   "_"
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   "Prog-Proy-Act-Partida"
         BeginProperty Font 
            Name            =   "Gill Sans MT"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000040C0&
         Height          =   255
         Left            =   3480
         TabIndex        =   5
         Top             =   1080
         Width           =   2175
      End
      Begin VB.Label Label1 
         Caption         =   "Nombre Completo"
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
         TabIndex        =   3
         Top             =   480
         Width           =   1815
      End
      Begin VB.Label Label3 
         Caption         =   "Estructura Pres."
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
         TabIndex        =   2
         Top             =   1080
         Width           =   1815
      End
   End
   Begin VB.CommandButton cmdAgregar 
      BackColor       =   &H008080FF&
      Caption         =   "Guardar Datos"
      Height          =   495
      Left            =   2280
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   4320
      Width           =   1815
   End
End
Attribute VB_Name = "CargaPrecarizadoImputado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()

    Call CenterMe(CargaPrecarizadoImputado, 6500, 5250)

End Sub

Private Sub cmdAgregar_Click()

    GenerarPrecarizadoImputado

End Sub

Private Sub Form_Unload(Cancel As Integer)

    strEditandoPrecarizadoImputado = ""

End Sub
