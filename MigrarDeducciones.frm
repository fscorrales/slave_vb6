VERSION 5.00
Begin VB.Form MigrarDeducciones 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Migrar Deducciones"
   ClientHeight    =   6930
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4815
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6930
   ScaleWidth      =   4815
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame3 
      Caption         =   "Datos a Migrar"
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
      Height          =   1095
      Left            =   120
      TabIndex        =   9
      Top             =   2520
      Width           =   4575
      Begin VB.ComboBox cmbTipoDatos 
         Height          =   315
         Left            =   1320
         TabIndex        =   10
         Top             =   480
         Width           =   3165
      End
      Begin VB.Label Label3 
         Caption         =   "Tipo Datos"
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
         Top             =   480
         Width           =   1095
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Resumen de Proceso"
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
      Height          =   2535
      Left            =   120
      TabIndex        =   7
      Top             =   3720
      Width           =   4575
      Begin VB.TextBox txtInforme 
         Height          =   2055
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   8
         Top             =   360
         Width           =   4335
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Período Destino"
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
      Height          =   1095
      Left            =   120
      TabIndex        =   4
      Top             =   1320
      Width           =   4575
      Begin VB.ComboBox cmbPeriodoDDJJDestino 
         Height          =   315
         Left            =   1320
         TabIndex        =   5
         Top             =   480
         Width           =   3165
      End
      Begin VB.Label Label2 
         Caption         =   "Año DDJJ"
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
         Width           =   1095
      End
   End
   Begin VB.CommandButton cmdMigrar 
      BackColor       =   &H008080FF&
      Caption         =   "Migrar DDJJ"
      Height          =   495
      Left            =   1440
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   6360
      Width           =   1935
   End
   Begin VB.Frame frmMultiuso 
      Caption         =   "Périodo Origen"
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
      Height          =   1095
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   4575
      Begin VB.ComboBox cmbPeriodoDDJJOrigen 
         Height          =   315
         Left            =   1320
         TabIndex        =   0
         Top             =   480
         Width           =   3165
      End
      Begin VB.Label Label1 
         Caption         =   "Año DDJJ"
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
         Width           =   1095
      End
   End
End
Attribute VB_Name = "MigrarDeducciones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdMigrar_Click()

    Dim strPeriodoOrigen As String
    Dim strPeriodoDestino As String
    Dim strTipoDato As String
    
    strPeriodoOrigen = Me.cmbPeriodoDDJJOrigen.Text
    strPeriodoDestino = Me.cmbPeriodoDDJJDestino.Text
    strTipoDato = Me.cmbTipoDatos.Text
    
    If ValidarMigrarDeducciones(strPeriodoOrigen, _
    strPeriodoDestino, strTipoDato) = True Then
        Call MigrarDeduccionesSIRADIG(strPeriodoOrigen, strPeriodoDestino, strTipoDato)
    End If

End Sub

Private Sub Form_Load()

    Call CenterMe(MigrarDeducciones, 4900, 7350)

End Sub

