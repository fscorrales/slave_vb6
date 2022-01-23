VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form ListadoLiquidacionesSISPER 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Liquidaciones SISPER"
   ClientHeight    =   5610
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   4905
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5610
   ScaleWidth      =   4905
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdEliminar 
      BackColor       =   &H008080FF&
      Caption         =   "Eliminar Liquidación"
      Height          =   375
      Left            =   1320
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   5160
      Width           =   2175
   End
   Begin VB.Frame frmMultiuso 
      Caption         =   "Resumen Liquidaciones de Haberes"
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
      TabIndex        =   0
      Top             =   120
      Width           =   4649
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid dgCodigoLiquidacion 
         Height          =   4455
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   4395
         _ExtentX        =   7752
         _ExtentY        =   7858
         _Version        =   393216
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
   End
End
Attribute VB_Name = "ListadoLiquidacionesSISPER"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()

    Call CenterMe(ListadoLiquidacionesSISPER, 5000, 6050)

End Sub


Private Sub cmdEliminar_Click()

    EliminarLiquidacionSISPER

End Sub
