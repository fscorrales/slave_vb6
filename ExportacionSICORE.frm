VERSION 5.00
Begin VB.Form ExportacionSICORE 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Generación de Archivo txt para SICORE"
   ClientHeight    =   2385
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   6870
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2385
   ScaleWidth      =   6870
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame3 
      Caption         =   "Datos de Generación txt"
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
      TabIndex        =   5
      Top             =   120
      Width           =   6615
      Begin VB.TextBox txtDecimal 
         Height          =   285
         Left            =   4920
         TabIndex        =   3
         Top             =   1080
         Width           =   1575
      End
      Begin VB.TextBox txtFecha 
         Height          =   285
         Left            =   1440
         TabIndex        =   2
         Top             =   1080
         Width           =   1575
      End
      Begin VB.TextBox txtCodigoLiquidacion 
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
      Begin VB.Label Label2 
         Caption         =   "Signo Decimal"
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
         Top             =   1080
         Width           =   1455
      End
      Begin VB.Label Label3 
         Caption         =   "Fecha Emisión"
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
         TabIndex        =   8
         Top             =   1080
         Width           =   1455
      End
      Begin VB.Label Label1 
         Caption         =   "Código Liq."
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
      Begin VB.Label Año 
         Caption         =   "Periodo"
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
         TabIndex        =   6
         Top             =   480
         Width           =   1455
      End
   End
   Begin VB.CommandButton cmdGenerarSICORE 
      BackColor       =   &H008080FF&
      Caption         =   "Generar Archivo SICORE"
      Height          =   495
      Left            =   2520
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1800
      Width           =   1815
   End
End
Attribute VB_Name = "ExportacionSICORE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdGenerarSICORE_Click()

    If GenerarArchivoSICORE(App.Path & "\SICORE.txt", Me.txtCodigoLiquidacion.Text, _
    Me.txtPeriodo.Text, Me.txtFecha.Text, Me.txtDecimal.Text) Then
        MsgBox " Datos exportados en " & App.Path, vbInformation
    End If

End Sub

Private Sub Form_Load()

    Call CenterMe(ExportacionSICORE, 6950, 2750)

End Sub

