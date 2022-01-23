VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form ImportacionLiquidacionSueldo 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Importar Liquidaci�n de Haberes SISPER"
   ClientHeight    =   5520
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4815
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5520
   ScaleWidth      =   4815
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdImportar 
      BackColor       =   &H008080FF&
      Caption         =   "Importar Liquidaci�n de Haberse SISPER"
      Height          =   495
      Left            =   1440
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   4920
      Width           =   1935
   End
   Begin VB.Frame frmMultiuso 
      Caption         =   "Identificaci�n Liquidaci�n"
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
      Top             =   120
      Width           =   4575
      Begin VB.ComboBox cmbCodigoLiquidacion 
         Height          =   315
         Left            =   1440
         TabIndex        =   1
         Top             =   480
         Width           =   2925
      End
      Begin VB.Label Label1 
         Caption         =   "C�digo"
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
   End
   Begin VB.Frame Frame1 
      Caption         =   "Resumen de Descarga"
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
      Height          =   3375
      Left            =   120
      TabIndex        =   0
      Top             =   1320
      Width           =   4575
      Begin VB.TextBox txtInforme 
         Height          =   3015
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   3
         Top             =   360
         Width           =   4335
      End
   End
   Begin MSComDlg.CommonDialog dlgMultifuncion 
      Left            =   4080
      Top             =   4920
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "ImportacionLiquidacionSueldo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdImportar_Click()

    Dim SQL As String
    
    SQL = "Select * From CODIGOLIQUIDACIONES Where CODIGO= '" & Left(Me.cmbCodigoLiquidacion.Text, 4) & "'"
    
    If SQLNoMatch(SQL) Then
        MsgBox "Debe seleccionar un C�digo de Liquidaci�n del Listado", vbCritical + vbOKOnly, "C�DIGO LIQUIDACI�N INEXISTENTE"
        Me.cmbCodigoLiquidacion.SetFocus
    Else
        ImportarLiquidacionSueldosSchema (Left(Me.cmbCodigoLiquidacion.Text, 4))
    End If
End Sub

Private Sub Form_Load()

    Call CenterMe(ImportacionLiquidacionSueldo, 4900, 5850)

End Sub
