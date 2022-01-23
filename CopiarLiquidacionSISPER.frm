VERSION 5.00
Begin VB.Form CopiarLiquidacionSisper 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Copiar Liquidación de Haberes SISPER"
   ClientHeight    =   5715
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4815
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5715
   ScaleWidth      =   4815
   ShowInTaskbar   =   0   'False
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
      Top             =   2520
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
      Caption         =   "Liquidación Destino"
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
      Begin VB.ComboBox cmbCodigoLiquidacionDestino 
         Height          =   315
         Left            =   1320
         TabIndex        =   5
         Top             =   480
         Width           =   3165
      End
      Begin VB.Label Label2 
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
         TabIndex        =   6
         Top             =   480
         Width           =   1095
      End
   End
   Begin VB.CommandButton cmdCopiar 
      BackColor       =   &H008080FF&
      Caption         =   "Copiar Liquidación"
      Height          =   495
      Left            =   1440
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   5160
      Width           =   1935
   End
   Begin VB.Frame frmMultiuso 
      Caption         =   "Liquidación Origen"
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
      Begin VB.ComboBox cmbCodigoLiquidacionOrigen 
         Height          =   315
         Left            =   1320
         TabIndex        =   0
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
         TabIndex        =   3
         Top             =   480
         Width           =   1095
      End
   End
End
Attribute VB_Name = "CopiarLiquidacionSisper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCopiar_Click()

    Dim SQL As String
    Dim SQL2 As String
    Dim strCodigoOrigen As String
    Dim strCodigoDestino As String
    
    strCodigoOrigen = Left(Me.cmbCodigoLiquidacionOrigen.Text, 4)
    strCodigoDestino = Left(Me.cmbCodigoLiquidacionDestino.Text, 4)
    
    SQL = "Select * From CODIGOLIQUIDACIONES Where CODIGO= '" & strCodigoOrigen & "'"
    SQL2 = "Select * From CODIGOLIQUIDACIONES Where CODIGO= '" & strCodigoDestino & "'"
    
    
    If SQLNoMatch(SQL) Then
        MsgBox "Debe seleccionar un Código de Liquidación del Listado", vbCritical + vbOKOnly, "CÓDIGO LIQUIDACIÓN DE ORIGEN INEXISTENTE"
        Me.cmbCodigoLiquidacionOrigen.SetFocus
    ElseIf SQLNoMatch(SQL2) Then
        MsgBox "Debe seleccionar un Código de Liquidación del Listado", vbCritical + vbOKOnly, "CÓDIGO LIQUIDACIÓN DE DESTINO INEXISTENTE"
        Me.cmbCodigoLiquidacionDestino.SetFocus
    ElseIf strCodigoOrigen = strCodigoDestino Then
        MsgBox "Los códigos de liquidación origen y destino deben ser distintos", vbCritical + vbOKOnly, "CÓDIGOS DUPLICADOS"
        Me.cmbCodigoLiquidacionDestino.SetFocus
    Else
        Call LiquidacionOrigenALiquidacionDestino(strCodigoOrigen, strCodigoDestino)
    End If

End Sub

Private Sub Form_Load()

    Call CenterMe(CopiarLiquidacionSisper, 4900, 6150)

End Sub

