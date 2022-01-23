VERSION 5.00
Begin VB.Form ExportacionSLAVE 
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   2370
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   6870
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2370
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
      Height          =   1575
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   6615
      Begin VB.OptionButton optAgentes 
         Caption         =   "Exportar Precarizados"
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
         TabIndex        =   7
         Top             =   1080
         Width           =   2535
      End
      Begin VB.OptionButton optComprobantes 
         Caption         =   "Exportar Comprobantes"
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
         TabIndex        =   6
         Top             =   1080
         Value           =   -1  'True
         Width           =   2535
      End
      Begin VB.TextBox txtAno 
         Height          =   285
         Left            =   1440
         TabIndex        =   0
         Top             =   480
         Width           =   1575
      End
      Begin VB.TextBox txtDecimal 
         Height          =   285
         Left            =   4920
         TabIndex        =   1
         Top             =   480
         Width           =   1575
      End
      Begin VB.Label Label1 
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
         Left            =   120
         TabIndex        =   5
         Top             =   480
         Width           =   1455
      End
      Begin VB.Label Año 
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
         TabIndex        =   4
         Top             =   480
         Width           =   1455
      End
   End
   Begin VB.CommandButton cmdGenerarSLAVE 
      BackColor       =   &H008080FF&
      Caption         =   "Generar Archivo CSV"
      Height          =   495
      Left            =   2520
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1800
      Width           =   1815
   End
End
Attribute VB_Name = "ExportacionSLAVE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()

    Call CenterMe(ExportacionSLAVE, 6950, 2750)

End Sub

Private Sub cmdGenerarSLAVE_Click()

    If Me.optComprobantes.Value = True Then
        If GenerarArchivoLiquidacionHonorarios(App.Path & "\HonorariosSIIF.csv", _
        Me.txtAno.Text, Me.txtDecimal.Text) Then
            MsgBox " Datos exportados en " & App.Path, vbInformation
        End If
    ElseIf Me.optAgentes.Value = True Then
        If GenerarArchivoPrecarizados(App.Path & "\PrecarizadosSIIF.csv") Then
            MsgBox " Datos exportados en " & App.Path, vbInformation
        End If
    End If

End Sub
