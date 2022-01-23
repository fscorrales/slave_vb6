VERSION 5.00
Begin VB.Form Registro 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "Registro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    Call RegisterComponent(App.Path & "\Componentes\" & "dao360.dll", True, True)
    Call RegisterComponent(App.Path & "\Componentes\" & "dao360.dll", False, True)
    Call RegisterComponent(App.Path & "\Componentes\" & "MSADODC.OCX", False, True)
    Call RegisterComponent(App.Path & "\Componentes\" & "MSADODC.OCX", True, True)
    Call RegisterComponent(App.Path & "\Componentes\" & "msado27.tlb", False, True)
    Call RegisterComponent(App.Path & "\Componentes\" & "msado27.tlb", True, True)
End Sub
