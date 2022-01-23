VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form ImportacionPadronFamiliares 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Importar Padrón Familiares"
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
      Caption         =   "Importar Padrón Familiares"
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
Attribute VB_Name = "ImportacionPadronFamiliares"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdImportar_Click()

    ImportarPadronFamiliares

End Sub

Private Sub Form_Load()

    Call CenterMe(ImportacionPadronFamiliares, 4700, 5600)

End Sub
