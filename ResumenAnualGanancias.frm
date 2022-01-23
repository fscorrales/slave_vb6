VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form ResumenAnualGanancias 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Resumen de Liquidación Ganancias 4ta Categoría"
   ClientHeight    =   7035
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   12915
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7039.958
   ScaleMode       =   0  'User
   ScaleWidth      =   12915
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdExportarTodo 
      BackColor       =   &H008080FF&
      Caption         =   "Exportar Todo"
      Height          =   495
      Left            =   9960
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   6480
      Width           =   1695
   End
   Begin VB.CommandButton cmdExportarSelección 
      BackColor       =   &H008080FF&
      Caption         =   "Exportar Selección"
      Height          =   495
      Left            =   1920
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   6480
      Width           =   1695
   End
   Begin VB.Frame Frame2 
      Caption         =   "Identificación del Período y Agente"
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
      Height          =   855
      Left            =   120
      TabIndex        =   2
      Top             =   0
      Width           =   12615
      Begin VB.CommandButton cmdActualizar 
         BackColor       =   &H008080FF&
         Caption         =   "Actualizar"
         Height          =   495
         Left            =   10440
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   240
         Width           =   1695
      End
      Begin VB.ComboBox cmbAgente 
         Height          =   315
         Left            =   4440
         TabIndex        =   6
         Top             =   360
         Width           =   5535
      End
      Begin VB.ComboBox cmbPeriodo 
         Height          =   315
         Left            =   1560
         TabIndex        =   5
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label Label7 
         Caption         =   "Período"
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
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label Label8 
         Caption         =   "Agente"
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
         Left            =   3000
         TabIndex        =   3
         Top             =   360
         Width           =   1455
      End
   End
   Begin VB.Frame Frame1 
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
      Height          =   5415
      Left            =   120
      TabIndex        =   0
      Top             =   960
      Width           =   12615
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid dgLiquidacionMensualGanancias 
         Height          =   5055
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   12375
         _ExtentX        =   21828
         _ExtentY        =   8916
         _Version        =   393216
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
   End
End
Attribute VB_Name = "ResumenAnualGanancias"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Dim intMeses As Integer

Private Sub cmbPeriodo_LostFocus()
    
    If ValidarCmbPeriodoResumenAnualGanancias = True Then
        CargarCmbAgenteResumenAnualGanancias (Me.cmbPeriodo.Text)
    End If
    
End Sub

Private Sub cmdActualizar_Click()

    Dim strPuestoLaboral As String
    
    If ValidarActualizarResumenAnualGanancias = True Then
        strPuestoLaboral = "Select PUESTOLABORAL From AGENTES" _
        & " Where NOMBRECOMPLETO = '" & Me.cmbAgente.Text & "'"
        Set rstBuscarSlave = New ADODB.Recordset
        rstBuscarSlave.Open strPuestoLaboral, dbSlave, adOpenDynamic, adLockOptimistic
        strPuestoLaboral = rstBuscarSlave!PuestoLaboral
        rstBuscarSlave.Close
        Set rstBuscarSlave = Nothing
        Call CargardgLiquidacionMensualGanancias(strPuestoLaboral, Me.cmbPeriodo.Text)
    End If

End Sub

Private Sub cmdExportarSelección_Click()

    Dim strAgente As String
    
    strResumenGcias = Me.cmbPeriodo.Text
    
    strResumenGcias = "\ResumenAnualGcias " & strResumenGcias & ".xls"
        
    strAgente = Me.cmbAgente.Text
        
    If Exportar_Excel(App.Path & strResumenGcias, Me.dgLiquidacionMensualGanancias, strAgente) Then
        MsgBox " Datos exportados en " & App.Path, vbInformation
    End If

End Sub

Private Sub cmdExportarTodo_Click()

    strResumenGcias = Me.cmbPeriodo.Text
    
    strResumenGcias = "\ResumenAnualGcias " & strResumenGcias & ".xls"
        
    If ExportarResumenAnualGcias(App.Path & strResumenGcias) Then
        MsgBox " Datos exportados en " & App.Path, vbInformation
    End If

End Sub

Private Sub Form_Load()

    Call CenterMe(ResumenAnualGanancias, 13000, 7400)
    CargarCmbPeriodoResumenAnualGanancias
    CargarCmbAgenteResumenAnualGanancias (Me.cmbPeriodo.Text)
'    EstablecerMeses
'    Call ConfigurardgEjecucionMensual(intMeses)
'    Call CargardgEjecucionMensual(intMeses)
'    intMeses = 0
'    ConfigurardgComprobantesEjecucionMensual
'    With Me.dgEjecucion
'        If .TextMatrix(1, 0) <> "" Then
'            ConfigurardgComprobantesEjecucionMensual
'            CargardgComprobantesEjecucionMensual (.TextMatrix(1, 0) & "-" & .TextMatrix(1, 1))
'        End If
'    End With
    
End Sub

'Sub EstablecerMeses()
'
'    Dim SQL As String
'
'    Set rstBuscarIcaro = New ADODB.Recordset
'    SQL = "Select COMPROBANTE, FECHA From CARGA Where Right(COMPROBANTE,2) = " & Right(strEjercicio, 2) & " Order by Fecha Desc"
'    rstBuscarIcaro.Open SQL, dbIcaro, adOpenForwardOnly, adLockOptimistic
'    If rstBuscarIcaro.BOF = False Then
'        rstBuscarIcaro.MoveFirst
'        intMeses = Format(rstBuscarIcaro!Fecha, "mm")
'    Else
'        intMeses = 12
'    End If
'
'    Set rstBuscarIcaro = Nothing
'    SQL = ""
'
'End Sub

Private Sub dgEjecucion_RowColChange()

'    Dim i As Integer
'
'    With Me.dgEjecucion
'        i = .Row
'        If i <> 0 Then
'            ConfigurardgComprobantesEjecucionMensual
'            CargardgComprobantesEjecucionMensual (.TextMatrix(i, 0) & "-" & .TextMatrix(i, 1))
'        End If
'    End With
'
'    i = 0
    
End Sub

