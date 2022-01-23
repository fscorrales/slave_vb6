VERSION 5.00
Begin VB.Form CargaCodigoSIRADIG 
   BorderStyle     =   4  'Fixed ToolWindow
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
   Begin VB.Frame frRecuadro 
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
      TabIndex        =   2
      Top             =   120
      Width           =   5655
      Begin VB.TextBox txtCodigo 
         Height          =   285
         Left            =   1440
         TabIndex        =   0
         Top             =   480
         Width           =   800
      End
      Begin VB.TextBox txtDenominacion 
         Height          =   285
         Left            =   1440
         TabIndex        =   1
         Top             =   1080
         Width           =   4000
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
         TabIndex        =   5
         Top             =   480
         Width           =   1455
      End
      Begin VB.Label Label3 
         Caption         =   "Denominación"
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
         Top             =   1080
         Width           =   1455
      End
   End
   Begin VB.CommandButton cmdAgregar 
      BackColor       =   &H008080FF&
      Caption         =   "Guardar Datos"
      Height          =   495
      Left            =   2040
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1800
      Width           =   1815
   End
End
Attribute VB_Name = "CargaCodigoSIRADIG"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()

    Call CenterMe(CargaCodigoSIRADIG, 6000, 2700)
    
    Select Case strCargaCodigoSIRADIG
    Case "Parentesco"
        Me.Caption = "Parentesco SIRADIG"
        Me.frRecuadro.Caption = "Datos del Parentesco SIRADIG"
    Case "Deducciones"
        Me.Caption = "Deducciones SIRADIG"
        Me.frRecuadro.Caption = "Datos de las Deducciones SIRADIG"
    Case "OtrasDeducciones"
        Me.Caption = "Otras Deducciones SIRADIG"
        Me.frRecuadro.Caption = "Datos de las Otras Deducciones SIRADIG"
    End Select

End Sub

Private Sub cmdAgregar_Click()

    Select Case strCargaCodigoSIRADIG
    Case "Parentesco"
        Call GenerarCodigoSIRADIG("ParentescoSIRADIG")
    Case "Deducciones"
        Call GenerarCodigoSIRADIG("DeduccionesSIRADIG")
    Case "OtrasDeducciones"
        Call GenerarCodigoSIRADIG("OtrasDeduccionesSIRADIG")
    End Select

End Sub

Private Sub Form_Unload(Cancel As Integer)

    strEditandoCodigoSIRADIG = ""
    strCargaCodigoSIRADIG = ""

End Sub
