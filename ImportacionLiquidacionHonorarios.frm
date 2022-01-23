VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form ImportacionLiquidacionHonorarios 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Importar Liquidación Honorarios"
   ClientHeight    =   5175
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   4590
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5175
   ScaleWidth      =   4590
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtInforme 
      Height          =   4335
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   1
      Top             =   120
      Width           =   4335
   End
   Begin VB.CommandButton cmdImportar 
      BackColor       =   &H008080FF&
      Caption         =   "Importar Honorarios"
      Height          =   495
      Left            =   1320
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   4560
      Width           =   1935
   End
   Begin MSComDlg.CommonDialog dlgMultifuncion 
      Left            =   3960
      Top             =   4560
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "ImportacionLiquidacionHonorarios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdImportar_Click()
    
    If bolImportandoLiquidacionHonorariosSLAVE = True Then
        ImportarLiquidacionHonorariosSLAVE
    ElseIf bolImportandoPrecarizadosSLAVE = True Then
        ImportarPrecarizadosSLAVE
    Else
        ImportarLiquidacionHonorarios
    End If
    
    
End Sub

Private Sub Form_Load()

    Call CenterMe(ImportacionLiquidacionHonorarios, 4700, 5600)
    If bolImportandoLiquidacionHonorariosSLAVE = True Then
        Me.Caption = "Importar Liquidación Honorarios (SLAVE)"
        Me.cmdImportar.Caption = "Importar Honorarios SLAVE"
    ElseIf bolImportandoPrecarizadosSLAVE = True Then
        Me.Caption = "Importar Precarizados (SLAVE)"
        Me.cmdImportar.Caption = "Importar Precarizados SLAVE"
    Else
        Me.Caption = "Importar Liquidación Honorarios (Gestión Financiera)"
        Me.cmdImportar.Caption = "Importar Honorarios G.Fciera."
    End If

End Sub

Private Sub Form_Unload(Cancel As Integer)

    bolImportandoLiquidacionHonorariosSLAVE = False
    bolImportandoPrecarizadosSLAVE = False

End Sub
