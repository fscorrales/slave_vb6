Attribute VB_Name = "SubMain"
    Call RegisterComponent(App.Path & "\Componentes\" & "dao360.dll", True, True)
    Call RegisterComponent(App.Path & "\Componentes\" & "dao360.dll", False, True)
    Call RegisterComponent(App.Path & "\Componentes\" & "MSADODC.OCX", False, True)
    Call RegisterComponent(App.Path & "\Componentes\" & "MSADODC.OCX", True, True)
    Call RegisterComponent(App.Path & "\Componentes\" & "msado27.tlb", False, True)
    Call RegisterComponent(App.Path & "\Componentes\" & "msado27.tlb", True, True)

