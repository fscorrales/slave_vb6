VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form ListadoHonorariosImputados 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Precarizados"
   ClientHeight    =   7395
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   6405
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7395
   ScaleWidth      =   6405
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame2 
      Caption         =   "Datos del Comprobante"
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
      Height          =   855
      Left            =   120
      TabIndex        =   6
      Top             =   0
      Width           =   6200
      Begin VB.TextBox txtPeriodo 
         Height          =   285
         Left            =   4680
         TabIndex        =   9
         Top             =   360
         Width           =   1335
      End
      Begin VB.TextBox txtComprobante 
         Height          =   285
         Left            =   1440
         TabIndex        =   7
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label Label1 
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
         Left            =   3360
         TabIndex        =   10
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label Label9 
         Caption         =   "Nro CyO SIIF"
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
         Top             =   360
         Width           =   1455
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Listado de Precarizados"
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
      Height          =   4935
      Left            =   120
      TabIndex        =   3
      Top             =   960
      Width           =   6200
      Begin VB.TextBox txtTotalImputacion 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   300
         Left            =   4560
         TabIndex        =   11
         Top             =   4560
         Width           =   1125
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid dgListadoHonorariosImputados 
         Height          =   4095
         Left            =   120
         TabIndex        =   4
         Top             =   360
         Width           =   5865
         _ExtentX        =   10345
         _ExtentY        =   7223
         _Version        =   393216
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Acciones Posibles"
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
      Height          =   1335
      Left            =   120
      TabIndex        =   0
      Top             =   6000
      Width           =   6200
      Begin VB.CommandButton cmdExportar 
         BackColor       =   &H008080FF&
         Caption         =   "Exportar Listado"
         Height          =   375
         Left            =   2160
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   840
         Width           =   1850
      End
      Begin VB.CommandButton cmdEliminar 
         BackColor       =   &H008080FF&
         Caption         =   "Eliminar Imputación"
         Height          =   375
         Left            =   4200
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   360
         Width           =   1850
      End
      Begin VB.CommandButton cmdAgregar 
         BackColor       =   &H008080FF&
         Caption         =   "Agregar Imputación"
         Height          =   375
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   360
         Width           =   1850
      End
      Begin VB.CommandButton cmdEditar 
         BackColor       =   &H008080FF&
         Caption         =   "Editar Imputación"
         Height          =   375
         Left            =   2160
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   360
         Width           =   1850
      End
   End
End
Attribute VB_Name = "ListadoHonorariosImputados"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Private Sub cmdAgregar_Click()
'
'    CargaPrecarizado.Show
'    Unload ListadoPrecarizados
'
'End Sub

Private Sub cmdEditar_Click()

    EditarPrecarizadoImputado

End Sub

Private Sub cmdEliminar_Click()

    EliminarPrecarizadoImputado

End Sub

Private Sub cmdExportar_Click()

    Dim strPeriodo As String
    Dim strCYO As String
    Dim strNombreArchivo As String
    
    strPeriodo = Replace(Me.txtPeriodo.Text, "/", "-")
    strCYO = Replace(Me.txtComprobante.Text, "/", "-")
    
    strNombreArchivo = "\Listado Factureros del Periodo " & strPeriodo & " CYO " & strCYO & ".xls"
        
    If Exportar_Excel(App.Path & strNombreArchivo, Me.dgListadoHonorariosImputados) Then
        MsgBox " Datos exportados en " & App.Path, vbInformation
    End If

End Sub

Private Sub Form_Load()

    Call CenterMe(ListadoHonorariosImputados, 6500, 7830)
    Me.txtComprobante.Enabled = False
    Me.txtPeriodo.Enabled = False
    Me.cmdAgregar.Enabled = False


End Sub

