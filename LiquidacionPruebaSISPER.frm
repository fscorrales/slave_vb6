VERSION 5.00
Begin VB.Form LiquidacionPruebaSISPER 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Liquidación Prueba SISPER"
   ClientHeight    =   7545
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4815
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7545
   ScaleWidth      =   4815
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame3 
      Caption         =   "Liquidación Destino de la Prueba"
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
      TabIndex        =   11
      Top             =   1320
      Width           =   4575
      Begin VB.ComboBox cmbCodigoLiquidacionDestino 
         Height          =   315
         Left            =   1320
         TabIndex        =   12
         Top             =   480
         Width           =   3165
      End
      Begin VB.Label Label4 
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
         TabIndex        =   13
         Top             =   480
         Width           =   1095
      End
   End
   Begin VB.CommandButton cmdLiquidar 
      BackColor       =   &H008080FF&
      Caption         =   "Inicio Liquidación Prueba"
      Height          =   495
      Left            =   1440
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   6840
      Width           =   1935
   End
   Begin VB.Frame Frame2 
      Caption         =   "Identificación Concepto a Liquidar"
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
      Top             =   2520
      Width           =   4575
      Begin VB.TextBox txtPorcentaje 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   3600
         TabIndex        =   14
         Text            =   "0"
         Top             =   1080
         Width           =   855
      End
      Begin VB.TextBox txtMontoFijo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1320
         TabIndex        =   9
         Text            =   "0"
         Top             =   1080
         Width           =   855
      End
      Begin VB.ComboBox cmbConceptoLiquidacion 
         Height          =   315
         Left            =   1320
         TabIndex        =   6
         Top             =   480
         Width           =   3165
      End
      Begin VB.Label Label5 
         Caption         =   "Porcentaje"
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
         Left            =   2400
         TabIndex        =   15
         Top             =   1080
         Width           =   1095
      End
      Begin VB.Label Label3 
         Caption         =   "Monto Fijo"
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
         TabIndex        =   7
         Top             =   480
         Width           =   1095
      End
   End
   Begin VB.Frame frmMultiuso 
      Caption         =   "Liquidación Base sobre cual trabajar"
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
      TabIndex        =   3
      Top             =   120
      Width           =   4575
      Begin VB.ComboBox cmbCodigoLiquidacionBase 
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
         TabIndex        =   4
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
      Top             =   4200
      Width           =   4575
      Begin VB.TextBox txtInforme 
         Height          =   2055
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   2
         Top             =   360
         Width           =   4335
      End
   End
End
Attribute VB_Name = "LiquidacionPruebaSISPER"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdLiquidar_Click()

    Dim SQL As String
    Dim SQL2 As String
    Dim SQL3 As String
    Dim strCodigoBase As String
    Dim strCodigoDestino As String
    Dim strConcepto As String
    
    strCodigoBase = Left(Me.cmbCodigoLiquidacionBase.Text, 4)
    strCodigoDestino = Left(Me.cmbCodigoLiquidacionDestino.Text, 4)
    strConcepto = Left(Me.cmbConceptoLiquidacion.Text, 4)
    
    SQL = "Select * From CODIGOLIQUIDACIONES Where CODIGO= '" & strCodigoBase & "'"
    SQL2 = "Select * From CODIGOLIQUIDACIONES Where CODIGO= '" & strCodigoDestino & "'"
    SQL3 = "Select * From CONCEPTOS Where CODIGO= '" & strConcepto & "'"
    
    If SQLNoMatch(SQL) Then
        MsgBox "Debe seleccionar un Código de Liquidación del Listado", vbCritical + vbOKOnly, "CÓDIGO LIQUIDACIÓN DE ORIGEN INEXISTENTE"
        Me.cmbCodigoLiquidacionBase.SetFocus
    ElseIf SQLNoMatch(SQL2) Then
        MsgBox "Debe seleccionar un Código de Liquidación del Listado", vbCritical + vbOKOnly, "CÓDIGO LIQUIDACIÓN DE DESTINO INEXISTENTE"
        Me.cmbCodigoLiquidacionDestino.SetFocus
    ElseIf strCodigoBase = strCodigoDestino Then
        MsgBox "Los códigos de liquidación origen y destino deben ser distintos", vbCritical + vbOKOnly, "CÓDIGOS DUPLICADOS"
        Me.cmbCodigoLiquidacionDestino.SetFocus
    ElseIf SQLNoMatch(SQL3) Then
        MsgBox "Debe seleccionar un Código de Concepto del Listado", vbCritical + vbOKOnly, "CONCEPTO LIQUIDACIÓN INEXISTENTE"
        Me.cmbConceptoLiquidacion.SetFocus
    Else
        If IsNumeric(Me.txtMontoFijo.Text) And Me.txtMontoFijo.Text <> 0 Then
            Call LiquidacionPrueba(strCodigoBase, strCodigoDestino, strConcepto, Me.txtMontoFijo.Text)
        ElseIf IsNumeric(Me.txtPorcentaje.Text) And Me.txtPorcentaje.Text <> 0 Then
            Dim Porcenaje As Double
            Porcenaje = Me.txtPorcentaje.Text
            Porcenaje = 1 + Porcenaje / 100
            Call LiquidacionPrueba(strCodigoBase, strCodigoDestino, strConcepto, , Porcenaje)
            Porcenaje = 0
        Else
            Call LiquidacionPrueba(strCodigoBase, strCodigoDestino, strConcepto)
        End If
    
    End If

End Sub
'
'    Dim SQL As String
'    Dim SQL2 As String
'
'    SQL = "Select * From CODIGOLIQUIDACIONES Where CODIGO= '" & Left(Me.cmbCodigoLiquidacion.Text, 4) & "'"
'    SQL2 = "Select * From CONCEPTOS Where CODIGO= '" & Left(Me.cmbConceptoLiquidacion.Text, 4) & "'"
'
'
'    If SQLNoMatch(SQL) Then
'        MsgBox "Debe seleccionar un Código de Liquidación del Listado", vbCritical + vbOKOnly, "CÓDIGO LIQUIDACIÓN INEXISTENTE"
'        Me.cmbCodigoLiquidacion.SetFocus
'    ElseIf SQLNoMatch(SQL2) Then
'        MsgBox "Debe seleccionar un Código de Concepto del Listado", vbCritical + vbOKOnly, "CONCEPTO LIQUIDACIÓN INEXISTENTE"
'        Me.cmbConceptoLiquidacion.SetFocus
'    Else
'        'ImportarLiquidacionSueldos (Left(Me.cmbCodigoLiquidacion.Text, 4))
'    End If
'
'End Sub

'Private Sub cmdImportar_Click()
'
'    Dim SQL As String
'
'    SQL = "Select * From CODIGOLIQUIDACIONES Where CODIGO= '" & Left(Me.cmbCodigoLiquidacion.Text, 4) & "'"
'
'    If SQLNoMatch(SQL) Then
'        MsgBox "Debe seleccionar un Código de Liquidación del Listado", vbCritical + vbOKOnly, "CÓDIGO LIQUIDACIÓN INEXISTENTE"
'        Me.cmbCodigoLiquidacion.SetFocus
'    Else
'        IncorporarConceptoPorArchivo (Left(Me.cmbCodigoLiquidacion.Text, 4))
'    End If
'
'End Sub

Private Sub Form_Load()

    Call CenterMe(LiquidacionPruebaSISPER, 4900, 8000)

End Sub

