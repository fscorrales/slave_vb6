VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form CargaPrecarizado 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Precarizado"
   ClientHeight    =   2385
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   5940
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2385
   ScaleWidth      =   5940
   ShowInTaskbar   =   0   'False
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
      TabIndex        =   1
      Top             =   120
      Width           =   5655
      Begin MSMask.MaskEdBox mskEstructura 
         Height          =   285
         Left            =   1800
         TabIndex        =   5
         Top             =   1080
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   503
         _Version        =   393216
         MaxLength       =   12
         Mask            =   "##-##-##-###"
         PromptChar      =   "_"
      End
      Begin VB.TextBox txtNombreCompleto 
         Height          =   285
         Left            =   1800
         TabIndex        =   0
         Top             =   480
         Width           =   3675
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
         Left            =   3240
         TabIndex        =   6
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
         TabIndex        =   4
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
         TabIndex        =   3
         Top             =   1080
         Width           =   1815
      End
   End
   Begin VB.CommandButton cmdAgregar 
      BackColor       =   &H008080FF&
      Caption         =   "Guardar Datos"
      Height          =   495
      Left            =   2040
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1800
      Width           =   1815
   End
End
Attribute VB_Name = "CargaPrecarizado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()

    Call CenterMe(CargaPrecarizado, 6000, 2700)

End Sub

Private Sub cmdAgregar_Click()

    GenerarPrecarizado

End Sub

Private Sub Form_Unload(Cancel As Integer)

    strEditandoPrecarizado = ""

End Sub
