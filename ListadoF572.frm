VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form ListadoF572 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Resumen de F572 Presentados "
   ClientHeight    =   6465
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   12840
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6465
   ScaleWidth      =   12840
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame4 
      Caption         =   "Deducciones Generales"
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
      Height          =   2535
      Left            =   3240
      TabIndex        =   8
      Top             =   3840
      Width           =   9495
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid dgDeduccionesGenerales 
         Height          =   2175
         Left            =   120
         TabIndex        =   9
         Top             =   240
         Width           =   9255
         _ExtentX        =   16325
         _ExtentY        =   3836
         _Version        =   393216
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Presentaciones Realizadas"
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
      TabIndex        =   6
      Top             =   960
      Width           =   2970
      Begin VB.CommandButton cmdEliminar 
         BackColor       =   &H008080FF&
         Caption         =   "Eliminar DDJJ seleccionada"
         Height          =   495
         Left            =   600
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   4800
         Width           =   1695
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid dgPresentacionesF572 
         Height          =   4455
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   2715
         _ExtentX        =   4789
         _ExtentY        =   7858
         _Version        =   393216
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Identificación del Agente"
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
         Left            =   10680
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   240
         Width           =   1695
      End
      Begin VB.ComboBox cmbAgente 
         Height          =   315
         Left            =   1560
         TabIndex        =   4
         Top             =   360
         Width           =   5535
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
         Left            =   120
         TabIndex        =   3
         Top             =   360
         Width           =   1455
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Cargas de Familia"
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
      Height          =   2535
      Left            =   3240
      TabIndex        =   0
      Top             =   960
      Width           =   9495
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid dgCargasDeFamilia 
         Height          =   2175
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   9255
         _ExtentX        =   16325
         _ExtentY        =   3836
         _Version        =   393216
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
   End
End
Attribute VB_Name = "ListadoF572"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Dim intMeses As Integer

Private Sub cmbAgente_Change()
    
    ConfigurardgPresentacionesF572
    ConfigurardgCargasDeFamiliaF572
    ConfigurardgDeduccionesGeneralesF572
    
End Sub

Private Sub cmdActualizar_Click()

    Dim strCUIL As String
    
    strCUIL = "Select CUIL From AGENTES" _
    & " Where NOMBRECOMPLETO = '" & Me.cmbAgente.Text & "'"
    Set rstBuscarSlave = New ADODB.Recordset
    rstBuscarSlave.Open strCUIL, dbSlave, adOpenDynamic, adLockOptimistic
    strCUIL = rstBuscarSlave!CUIL
    rstBuscarSlave.Close
    Set rstBuscarSlave = Nothing
    ConfigurardgPresentacionesF572
    Call CargardgPresentacionesF572(strCUIL)
    ConfigurardgCargasDeFamiliaF572
    Call CargardgCargasDeFamiliaF572(Me.dgPresentacionesF572.TextMatrix(1, 0))
    ConfigurardgDeduccionesGeneralesF572
    Call CargardgDeduccionesGeneralesF572(Me.dgPresentacionesF572.TextMatrix(1, 0))


End Sub

Private Sub cmdEliminar_Click()

    EliminarF572

End Sub

Private Sub Form_Load()

    Call CenterMe(ListadoF572, 13000, 7000)
    CargarCmbAgenteListadoF572
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

Private Sub dgPresentacionesF572_RowColChange()

    Dim i As Integer
    i = Me.dgPresentacionesF572.Row
    ConfigurardgCargasDeFamiliaF572
    Call CargardgCargasDeFamiliaF572(Me.dgPresentacionesF572.TextMatrix(i, 0))
    Call CargardgDeduccionesGeneralesF572(Me.dgPresentacionesF572.TextMatrix(i, 0))
    i = 0

End Sub
