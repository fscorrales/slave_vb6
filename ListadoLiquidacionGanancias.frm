VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form ListadoLiquidacionGanancias 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Código Liquidación (Gcias. 4ta Categoría)"
   ClientHeight    =   7800
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   7860
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7800
   ScaleWidth      =   7860
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame4 
      Caption         =   "Liquidación Mensual Ganancias"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   1815
      Left            =   120
      TabIndex        =   8
      Top             =   5880
      Width           =   2955
      Begin VB.CommandButton cmdSICORE 
         BackColor       =   &H008080FF&
         Caption         =   "Exportar SICORE TXT"
         Height          =   375
         Left            =   360
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   1320
         Width           =   2300
      End
      Begin VB.CommandButton cmdReporte 
         BackColor       =   &H008080FF&
         Caption         =   "Generar Reporte XLS"
         Height          =   375
         Left            =   360
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   360
         Width           =   2300
      End
      Begin VB.CommandButton cmdExportar 
         BackColor       =   &H008080FF&
         Caption         =   "Exportar SISPER XLS"
         Height          =   375
         Left            =   360
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   840
         Width           =   2300
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Agente Retenido"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   1815
      Left            =   3240
      TabIndex        =   4
      Top             =   5880
      Width           =   4485
      Begin VB.CommandButton cmdAgregarSIRADIG 
         BackColor       =   &H008080FF&
         Caption         =   "Nuevo Agente SIRADIG"
         Height          =   375
         Left            =   2280
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   360
         Width           =   2000
      End
      Begin VB.CommandButton cmdEditarSIRADIG 
         BackColor       =   &H008080FF&
         Caption         =   "Editar SIRADIG"
         Height          =   375
         Left            =   2280
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   840
         Width           =   2000
      End
      Begin VB.CommandButton cmdEliminar 
         BackColor       =   &H008080FF&
         Caption         =   "Eliminar Retención"
         Height          =   375
         Left            =   1080
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   1320
         Width           =   2300
      End
      Begin VB.CommandButton cmdEditar 
         BackColor       =   &H008080FF&
         Caption         =   "Editar Retención"
         Height          =   375
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   840
         Width           =   2000
      End
      Begin VB.CommandButton cmdAgregar 
         BackColor       =   &H008080FF&
         Caption         =   "Nuevo Agente Retenido"
         Height          =   375
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   360
         Width           =   2000
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Retención Ganancias por Liquidación"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   5655
      Left            =   3240
      TabIndex        =   2
      Top             =   120
      Width           =   4530
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid dgAgentesRetenidos 
         Height          =   5175
         Left            =   120
         TabIndex        =   3
         Top             =   360
         Width           =   4275
         _ExtentX        =   7541
         _ExtentY        =   9128
         _Version        =   393216
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Códigos de Liquidaciones"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   5655
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2970
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid dgCodigosLiquidacionesGanancias 
         Height          =   5175
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   2715
         _ExtentX        =   4789
         _ExtentY        =   9128
         _Version        =   393216
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
   End
End
Attribute VB_Name = "ListadoLiquidacionGanancias"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAgregar_Click()

    Dim i As String
    Dim SQL As String
    Dim strCL As String
    Dim strPeriodo As String
    
    
    With Me.dgCodigosLiquidacionesGanancias
        i = .Row
        strCL = .TextMatrix(i, 0)
        strPeriodo = .TextMatrix(i, 1)
    End With
    
    With LiquidacionGanancia4ta
        .Show
        .txtCodigoLiquidacion.Text = strCL & " - (" & BuscarPeriodoLiquidacion(strCL) & ")"
        .txtDescripcionPeriodo.Text = strPeriodo
        '.txtPeriodo = BuscarPeriodoLiquidacion(strCL)
        .txtPuestoLaboral.Enabled = True
        .txtPuestoLaboral.SetFocus
    End With
    i = ""
    strCL = ""
    strPeriodo = ""
    Unload ListadoLiquidacionGanancias

End Sub

Private Sub cmdAgregarSIRADIG_Click()

    Dim i As String
    Dim SQL As String
    Dim strCL As String
    Dim strPeriodo As String
    
    
    With Me.dgCodigosLiquidacionesGanancias
        i = .Row
        strCL = .TextMatrix(i, 0)
        strPeriodo = .TextMatrix(i, 1)
    End With
    
    With LiquidacionGanancia4taSIRADIG
        .Show
        .txtCodigoLiquidacion.Text = strCL & " - (" & BuscarPeriodoLiquidacion(strCL) & ")"
        .txtDescripcionPeriodo.Text = strPeriodo
        '.txtPeriodo = BuscarPeriodoLiquidacion(strCL)
        .txtPuestoLaboral.Enabled = True
        .txtPuestoLaboral.SetFocus
    End With
    i = ""
    strCL = ""
    strPeriodo = ""
    Unload ListadoLiquidacionGanancias

End Sub

Private Sub cmdEditar_Click()

    EditarRetencionGanancias

End Sub

Private Sub cmdEliminar_Click()

    EliminarRetencionGanancias

End Sub

Private Sub cmdExportar_Click()

    Dim i As Integer
    i = Me.dgCodigosLiquidacionesGanancias.Row
    
    If IsNumeric(Me.dgAgentesRetenidos.TextMatrix(1, 2)) = True Then
        If ExportarLiquidacionGananciasSISPER(App.Path & "\ReporteSISPER.xls", Me.dgCodigosLiquidacionesGanancias.TextMatrix(i, 0)) Then
            MsgBox " Datos exportados en " & App.Path, vbInformation
        End If
    End If
    i = 0
    
'    If IsNumeric(Me.dgAgentesRetenidos.TextMatrix(1, 2)) = True Then
'        If ExportarLiquidacionGananciasSISPERViejo(App.Path & "\ReporteSISPER.xls", Me.dgCodigosLiquidacionesGanancias.TextMatrix(i, 0), Me.dgAgentesRetenidos) Then
'            MsgBox " Datos exportados en " & App.Path, vbInformation
'        End If
'    End If
'    i = 0

End Sub

Private Sub cmdSICORE_Click()

    Dim i As Integer
    Dim strCL As String
    Dim strPL As String
    i = Me.dgCodigosLiquidacionesGanancias.Row
    
    If IsNumeric(Me.dgAgentesRetenidos.TextMatrix(1, 2)) = True Then
        With ExportacionSICORE
            .Show
            strCL = Me.dgCodigosLiquidacionesGanancias.TextMatrix(i, 0)
            .txtCodigoLiquidacion.Text = strCL
            .txtCodigoLiquidacion.Enabled = False
            strPL = BuscarPeriodoLiquidacion(strCL)
            .txtPeriodo.Text = strPL
            .txtPeriodo.Enabled = False
            .txtFecha.Text = "26/" & strPL
            .txtFecha.Enabled = True
            .txtDecimal.Text = ","
            .txtDecimal.Enabled = True
        End With
        Unload ListadoLiquidacionGanancias
    End If
    i = 0

End Sub

Private Sub cmdReporte_Click()

    Dim i As Integer
    i = Me.dgCodigosLiquidacionesGanancias.Row
    
    If IsNumeric(Me.dgAgentesRetenidos.TextMatrix(1, 2)) = True Then
        If GenerarReporteMensualGanancias(App.Path & "\Ganancias4taCategoria.xls", Me.dgCodigosLiquidacionesGanancias.TextMatrix(i, 0)) Then
            MsgBox " Datos exportados en " & App.Path, vbInformation
        End If
    End If
    i = 0

End Sub

Private Sub dgCodigosLiquidacionesGanancias_RowColChange()

    Dim i As Integer
    i = Me.dgCodigosLiquidacionesGanancias.Row
    ConfigurardgAgentesRetenidos
    Call CargardgAgentesRetenidos(Me.dgCodigosLiquidacionesGanancias.TextMatrix(i, 0))
    i = 0

End Sub

Private Sub Form_Load()

    Call CenterMe(ListadoLiquidacionGanancias, 7950, 8130)

End Sub

Private Sub cmdEditarSIRADIG_Click()

    EditarRetencionGananciasSIRADIG

End Sub
