Attribute VB_Name = "ConfigurarListBox"
Public Sub PasarDatosListBox(ListOrigen As ListBox, ListDestino As ListBox)

    If Not ListOrigen.Text = "" Then
        ListDestino.AddItem ListOrigen.Text
        ListOrigen.RemoveItem ListOrigen.ListIndex
    End If

End Sub

Public Sub CargarlstDeduccionesPersonales(PuestoLaboral As String)
    
    Set rstListadoSlave = New ADODB.Recordset
    rstListadoSlave.Open "Select * From CARGASDEFAMILIA Where PUESTOLABORAL = " & "'" & PuestoLaboral & "' Order By FechaAlta", dbSlave, adOpenDynamic, adLockOptimistic
    With CargaDeduccionesPersonales
        If rstListadoSlave.BOF = False Then
            rstListadoSlave.MoveFirst
            While rstListadoSlave.EOF = False
                If rstListadoSlave!DeducibleGanancias = True Then
                    .lstFamiliaresDeduciblesGanancias.AddItem rstListadoSlave!NombreCompleto
                ElseIf rstListadoSlave!DeducibleGanancias = False Then
                    .lstFamiliaresACargo.AddItem rstListadoSlave!NombreCompleto
                End If
                rstListadoSlave.MoveNext
            Wend
        End If
    End With
    rstListadoSlave.Close
    Set rstListadoSlave = Nothing
    
End Sub

