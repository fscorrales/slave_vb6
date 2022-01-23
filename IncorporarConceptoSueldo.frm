VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form IncorporarConceptoSueldo 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Incorporar Concepto a Liquidación de Haberes SISPER"
   ClientHeight    =   6345
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4815
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6345
   ScaleWidth      =   4815
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdMontoFijo 
      BackColor       =   &H008080FF&
      Caption         =   "Incorporar Monto Fijo"
      Height          =   495
      Left            =   2760
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   5760
      Width           =   1935
   End
   Begin VB.Frame Frame2 
      Caption         =   "Identificación Concepto a Incorporar"
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
      Top             =   1320
      Width           =   4575
      Begin VB.CheckBox chkRemunerativo 
         Alignment       =   1  'Right Justify
         Caption         =   "¿Remunerativo?"
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
         Left            =   2640
         TabIndex        =   11
         Top             =   960
         Value           =   1  'Checked
         Width           =   1815
      End
      Begin VB.TextBox txtMonto 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1320
         TabIndex        =   10
         Top             =   1080
         Width           =   1215
      End
      Begin VB.ComboBox cmbConceptoLiquidacion 
         Height          =   315
         Left            =   1320
         TabIndex        =   7
         Top             =   480
         Width           =   3165
      End
      Begin VB.Label Label3 
         Caption         =   "Monto"
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
         Top             =   1080
         Width           =   1095
      End
      Begin VB.Label Label2 
         Caption         =   "Concepto"
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
         Top             =   480
         Width           =   1095
      End
   End
   Begin VB.CommandButton cmdImportar 
      BackColor       =   &H008080FF&
      Caption         =   "Incorporar por Archivo Plano SISPER"
      Height          =   495
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   5760
      Width           =   1935
   End
   Begin VB.Frame frmMultiuso 
      Caption         =   "Identificación Liquidación a Modificar"
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
         Left            =   1320
         TabIndex        =   1
         Top             =   480
         Width           =   3165
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
      TabIndex        =   0
      Top             =   3000
      Width           =   4575
      Begin VB.TextBox txtInforme 
         Height          =   2055
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   3
         Top             =   360
         Width           =   4335
      End
   End
   Begin MSComDlg.CommonDialog dlgMultifuncion 
      Left            =   2160
      Top             =   5760
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "IncorporarConceptoSueldo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdMontoFijo_Click()

    Dim SQL As String
    Dim SQL2 As String
    Dim strCodigoLiquidacion As String
    Dim strCodigoConcepto As String
    
    strCodigoLiquidacion = Left(Me.cmbCodigoLiquidacion.Text, 4)
    strCodigoConcepto = Left(Me.cmbConceptoLiquidacion.Text, 4)
    
    SQL = "Select * From CODIGOLIQUIDACIONES Where CODIGO= '" & Left(Me.cmbCodigoLiquidacion.Text, 4) & "'"
    SQL2 = "Select * From CONCEPTOS Where CODIGO= '" & Left(Me.cmbConceptoLiquidacion.Text, 4) & "'"
    
    
    If SQLNoMatch(SQL) Then
        MsgBox "Debe seleccionar un Código de Liquidación del Listado", vbCritical + vbOKOnly, "CÓDIGO LIQUIDACIÓN INEXISTENTE"
        Me.cmbCodigoLiquidacion.SetFocus
    ElseIf SQLNoMatch(SQL2) Then
        MsgBox "Debe seleccionar un Código de Concepto del Listado", vbCritical + vbOKOnly, "CONCEPTO LIQUIDACIÓN INEXISTENTE"
        Me.cmbConceptoLiquidacion.SetFocus
    Else
        Call IncorporarConceptoMontoFijo(strCodigoLiquidacion, strCodigoConcepto, Me.txtMonto.Text, Me.chkRemunerativo.Value)
    End If

End Sub

Private Sub cmdImportar_Click()

    Dim SQL As String
    
    SQL = "Select * From CODIGOLIQUIDACIONES Where CODIGO= '" & Left(Me.cmbCodigoLiquidacion.Text, 4) & "'"
    
    If SQLNoMatch(SQL) Then
        MsgBox "Debe seleccionar un Código de Liquidación del Listado", vbCritical + vbOKOnly, "CÓDIGO LIQUIDACIÓN INEXISTENTE"
        Me.cmbCodigoLiquidacion.SetFocus
    Else
        IncorporarConceptoPorArchivoSchema (Left(Me.cmbCodigoLiquidacion.Text, 4))
    End If

End Sub

Private Sub Form_Load()

    Call CenterMe(IncorporarConceptoSueldo, 4900, 6700)

End Sub

