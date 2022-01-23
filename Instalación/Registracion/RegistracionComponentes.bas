Attribute VB_Name = "RegistracionComponentes"
'Registering components using regsvr32

'Below is a routine for registering components using Regsvr32.exe.

'Purpose     :  Registers components (DLLs and OCXs)
'Inputs      :  sFileName               The path and file name of the component to register.
'               [bUnRegister]           If True unregisters the component, else registers the component.
'               [bHideResults]          If True the confirmation dialog will be hidden, else
'                                       a modal dialog will display the results.
'Outputs     :  N/A


Public Sub RegisterComponent(sFileName As String, Optional bUnRegister As Boolean = False, Optional bHideResults As Boolean = True)
    If Len(Dir$(sFileName)) = 0 Then
        'File is missing
        MsgBox "Unable to locate file "" & sFileName & """, vbCritical
    Else
        If bUnRegister Then
            'Unregister a component
            If bHideResults Then
                'Hide results
                Shell "regsvr32 /s /u " & """" & sFileName & """"
            Else
                'Show results
                Shell "regsvr32 /u " & """" & sFileName & """"
            End If
        Else
            'Register a component
            If bHideResults Then
                'Hide results
                Shell "regsvr32 /s " & """" & sFileName & """"
            Else
                'Show results
                Shell "regsvr32 " & """" & sFileName & """"
            End If
        End If
    End If
End Sub


