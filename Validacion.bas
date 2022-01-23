Attribute VB_Name = "Validacion"
Function ValidarAgente() As Boolean

    Dim SQL As String
    
    With CargaAgente
    
        If Trim(.txtPuestoLaboral.Text) = "" Or IsNumeric(.txtPuestoLaboral.Text) = False Or Len(.txtPuestoLaboral.Text) > "6" Then
            MsgBox "Debe ingresar un Nro de Puesto Laboral de hasta 6 cifras", vbCritical + vbOKOnly, "NRO PUESTO LABORAL INCORRECTO"
            .txtPuestoLaboral.SetFocus
            ValidarAgente = False
            Exit Function
        End If
        If strEditandoAgente <> "" And .txtPuestoLaboral.Text <> strEditandoAgente Then
            If SQLNoMatch("Select * from AGENTES Where PUESTOLABORAL= '" & .txtPuestoLaboral.Text & "'") = False Then
                MsgBox "Debe ingresar un Nro de Puesto Laboral �NICO", vbCritical + vbOKOnly, "NRO PUESTO LABORAL DUPLICADO"
                .txtPuestoLaboral.SetFocus
                ValidarAgente = False
                Exit Function
            End If
        ElseIf strEditandoAgente = "" Then
            If SQLNoMatch("Select * from AGENTES Where PUESTOLABORAL= '" & .txtPuestoLaboral.Text & "'") = False Then
                MsgBox "Debe ingresar un Nro de Puesto Laboral �NICO", vbCritical + vbOKOnly, "NRO PUESTO LABORAL DUPLICADO"
                .txtPuestoLaboral.SetFocus
                ValidarAgente = False
                Exit Function
            End If
        End If
        
        If Trim(.txtCUIL.Text) = "" Or IsNumeric(.txtCUIL.Text) = False Or Len(.txtCUIL.Text) > "11" Then
            MsgBox "Debe ingresar un Nro de CUIL/DNI de hasta 11 cifras", vbCritical + vbOKOnly, "NRO CUIL/DNI INCORRECTO"
            .txtCUIL.SetFocus
            ValidarAgente = False
            Exit Function
        End If
        If strEditandoAgente <> "" Then
            Set rstRegistroSlave = New ADODB.Recordset
            SQL = "Select * from AGENTES Where PUESTOLABORAL = " & "'" & strEditandoAgente & "'"
            rstRegistroSlave.Open SQL, dbSlave, adOpenForwardOnly, adLockReadOnly
            SQL = ""
            If .txtCUIL.Text <> rstRegistroSlave.Fields("CUIL") Then
                If SQLNoMatch("Select * from AGENTES Where CUIL= '" & .txtCUIL.Text & "'") = False Then
                    MsgBox "Debe ingresar un Nro de CUIL/DNI �NICO", vbCritical + vbOKOnly, "NRO CUIL/DNI DUPLICADO"
                    .txtCUIL.SetFocus
                    rstRegistroSlave.Close
                    Set rstRegistroSlave = Nothing
                    ValidarAgente = False
                    Exit Function
                End If
            End If
            rstRegistroSlave.Close
            Set rstRegistroSlave = Nothing
        ElseIf strEditandoAgente = "" Then
            If SQLNoMatch("Select * from AGENTES Where CUIL= '" & .txtCUIL.Text & "'") = False Then
                MsgBox "Debe ingresar un Nro de CUIL/DNI �NICO", vbCritical + vbOKOnly, "NRO CUIL/DNI DUPLICADO"
                .txtCUIL.SetFocus
                ValidarAgente = False
                Exit Function
            End If
        End If

        If Trim(.txtDescripcion.Text) = "" Or IsNumeric(.txtDescripcion.Text) = True Or Len(.txtDescripcion.Text) > "50" Then
            MsgBox "Debe ingresar el Nombre Completo del Agente de hasta 50 caracteres", vbCritical + vbOKOnly, "NOMBRE AGENTE INCORRECTO"
            .txtDescripcion.SetFocus
            ValidarAgente = False
            Exit Function
        End If
        If strEditandoAgente <> "" Then
            Set rstRegistroSlave = New ADODB.Recordset
            SQL = "Select * from AGENTES Where PUESTOLABORAL = " & "'" & strEditandoAgente & "'"
            rstRegistroSlave.Open SQL, dbSlave, adOpenForwardOnly, adLockReadOnly
            SQL = ""
            If .txtDescripcion.Text <> rstRegistroSlave.Fields("NOMBRECOMPLETO") Then
                If SQLNoMatch("Select * from AGENTES Where NOMBRECOMPLETO= '" & .txtDescripcion.Text & "'") = False Then
                    MsgBox "Debe ingresar un Nombre del Agente �NICO", vbCritical + vbOKOnly, "NOMBRE DEL AGENTE DUPLICADO"
                    .txtDescripcion.SetFocus
                    rstRegistroSlave.Close
                    Set rstRegistroSlave = Nothing
                    ValidarAgente = False
                    Exit Function
                End If
            End If
            rstRegistroSlave.Close
            Set rstRegistroSlave = Nothing
        ElseIf strEditandoAgente = "" Then
            If SQLNoMatch("Select * from AGENTES Where NOMBRECOMPLETO= '" & .txtDescripcion.Text & "'") = False Then
                MsgBox "Debe ingresar un Nombre del Agente �NICO", vbCritical + vbOKOnly, "NOMBRE DEL AGENTE DUPLICADO"
                .txtDescripcion.SetFocus
                ValidarAgente = False
                Exit Function
            End If
        End If

        If Trim(.txtLegajo.Text) = "" Or IsNumeric(.txtLegajo.Text) = False Or Len(.txtLegajo.Text) > "4" Then
            MsgBox "Debe ingresar un Nro de Legajo de hasta 4 cifras", vbCritical + vbOKOnly, "NRO LEGAJO INCORRECTO"
            .txtLegajo.SetFocus
            ValidarAgente = False
            Exit Function
        End If
        If strEditandoAgente <> "" Then
            Set rstRegistroSlave = New ADODB.Recordset
            SQL = "Select * from AGENTES Where PUESTOLABORAL = " & "'" & strEditandoAgente & "'"
            rstRegistroSlave.Open SQL, dbSlave, adOpenForwardOnly, adLockReadOnly
            SQL = ""
            If .txtLegajo.Text <> rstRegistroSlave.Fields("LEGAJO") Then
                If SQLNoMatch("Select * from AGENTES Where LEGAJO= '" & .txtLegajo.Text & "'") = False Then
                    MsgBox "Debe ingresar un Nro de Legajo �NICO", vbCritical + vbOKOnly, "NRO LEGAJO DUPLICADO"
                    .txtLegajo.SetFocus
                    rstRegistroSlave.Close
                    Set rstRegistroSlave = Nothing
                    ValidarAgente = False
                    Exit Function
                End If
            End If
            rstRegistroSlave.Close
            Set rstRegistroSlave = Nothing
        ElseIf strEditandoAgente = "" Then
            If SQLNoMatch("Select * from AGENTES Where LEGAJO= '" & .txtLegajo.Text & "'") = False Then
                MsgBox "Debe ingresar un Nro de Legajo �NICO", vbCritical + vbOKOnly, "NRO LEGAJO DUPLICADO"
                .txtLegajo.SetFocus
                ValidarAgente = False
                Exit Function
            End If
        End If

    End With
    ValidarAgente = True
    
End Function

Function ValidarConcepto() As Boolean

    Dim SQL As String
    Dim strCodigo As String
    
    With CargaConcepto
    
        If Trim(.txtCodigo.Text) = "" Or IsNumeric(.txtCodigo.Text) = False Or Len(.txtCodigo.Text) > "4" Then
            MsgBox "Debe ingresar un Nro de C�digo de hasta 4 cifras", vbCritical + vbOKOnly, "NRO C�DIGO INCORRECTO"
            .txtCodigo.SetFocus
            ValidarConcepto = False
            Exit Function
        End If
        strCodigo = Format(.txtCodigo.Text, "0000")
        If strEditandoConcepto <> "" And strCodigo <> strEditandoConcepto Then
            If SQLNoMatch("Select * from CONCEPTOS Where CODIGO= '" & strCodigo & "'") = False Then
                MsgBox "Debe ingresar un Nro de C�DIGO DE CONCEPTO �NICO", vbCritical + vbOKOnly, "NRO C�DIGO DUPLICADO"
                .txtCodigo.SetFocus
                ValidarConcepto = False
                Exit Function
            End If
        ElseIf strEditandoConcepto = "" Then
            If SQLNoMatch("Select * from CONCEPTOS Where CODIGO= '" & strCodigo & "'") = False Then
                MsgBox "Debe ingresar un Nro de C�DIGO DE CONCEPTO �NICO", vbCritical + vbOKOnly, "NRO C�DIGO DUPLICADO"
                .txtCodigo.SetFocus
                ValidarConcepto = False
                Exit Function
            End If
        End If
        
        If Trim(.txtDenominacion.Text) = "" Or IsNumeric(.txtDenominacion.Text) = True Or Len(.txtDenominacion.Text) > "30" Then
            MsgBox "Debe ingresar una Denomiaci�n de hasta 30 caracteres", vbCritical + vbOKOnly, "DENOMINACI�N INCORRECTA"
            .txtDenominacion.SetFocus
            ValidarConcepto = False
            Exit Function
        End If
        
    End With
    ValidarConcepto = True

End Function

Function ValidarNormaEscalaGanancias() As Boolean

    Dim SQL As String
    Dim datFecha As Date
    
    With CargaNormaEscalaGanancias
    
        If Trim(.txtNormaLegal.Text) = "" Or Len(.txtNormaLegal.Text) > "10" Then
            MsgBox "Debe ingresar una Norma Legal de hasta 10 caracteres", vbCritical + vbOKOnly, "NORMA LEGAL INCORRECTA"
            .txtNormaLegal.SetFocus
            ValidarNormaEscalaGanancias = False
            Exit Function
        End If
        If strEditandoNormaEscalaGanancias <> "" And .txtNormaLegal.Text <> strEditandoNormaEscalaGanancias Then
            If SQLNoMatch("Select NORMALEGAL from ESCALAAPLICABLEGANANCIAS Where NORMALEGAL= '" & .txtNormaLegal.Text & "' Group by NORMALEGAL") = False Then
                MsgBox "Debe ingresar una Norma Legal �NICA", vbCritical + vbOKOnly, "NORMA LEGAL DUPLICADA"
                .txtNormaLegal.SetFocus
                ValidarNormaEscalaGanancias = False
                Exit Function
            End If
        ElseIf strEditandoNormaEscalaGanancias = "" Then
            If SQLNoMatch("Select NORMALEGAL from ESCALAAPLICABLEGANANCIAS Where NORMALEGAL= '" & .txtNormaLegal.Text & "' Group by NORMALEGAL") = False Then
                MsgBox "Debe ingresar una Norma Legal �NICA", vbCritical + vbOKOnly, "NORMA LEGAL DUPLICADA"
                .txtNormaLegal.SetFocus
                ValidarNormaEscalaGanancias = False
                Exit Function
            End If
        End If
        
        If Trim(.txtFecha.Text) = "" Or IsDate(.txtFecha.Text) = False Then
            MsgBox "Debe ingresar una Fecha de la Norma Legal adecuada", vbCritical + vbOKOnly, "FECHA INCORRECTA"
            .txtFecha.SetFocus
            ValidarNormaEscalaGanancias = False
            Exit Function
        End If
        If .txtNormaLegal.Text = strEditandoNormaEscalaGanancias Then
            'Quiere decir que lo unico que estoy editando es la fecha y, por lo tanto, no puede existir otra fecha igual
            datFecha = DateTime.DateSerial(Right(.txtFecha.Text, 4), Mid(.txtFecha.Text, 4, 2), Left(.txtFecha.Text, 2))
            If SQLNoMatch("Select NORMALEGAL from ESCALAAPLICABLEGANANCIAS Where FECHA= #" & Format(datFecha, "MM/DD/YYYY") & "# Group by NORMALEGAL") = False Then
                MsgBox "Debe ingresar una Fecha �NICA s� es lo �nico que pretende editar", vbCritical + vbOKOnly, "FECHA DUPLICADA"
                .txtFecha.SetFocus
                ValidarNormaEscalaGanancias = False
                Exit Function
            End If
        End If
        
    End With
    ValidarNormaEscalaGanancias = True
    
End Function

Function ValidarEscalaGanancias() As Boolean

    Dim SQL As String
    Dim strNormaLegal As String
    
    With CargaEscalaGanancias
        
        strNormaLegal = .txtNormaLegal.Text
    
        If Trim(.txtTramo.Text) = "" Or IsNumeric(.txtTramo.Text) = False Or Len(.txtTramo.Text) > "2" Then
            MsgBox "Debe ingresar un Nro de Tramo de hasta 2 cifras", vbCritical + vbOKOnly, "NRO TRAMO INCORRECTO"
            .txtTramo.SetFocus
            ValidarEscalaGanancias = False
            Exit Function
        End If
        SQL = "Select * from ESCALAAPLICABLEGANANCIAS" _
        & " Where TRAMO = '" & .txtTramo.Text _
        & "' And NORMALEGAL = '" & strNormaLegal & "'"
        If strEditandoEscalaGanancias <> "" And .txtTramo.Text <> strEditandoEscalaGanancias Then
            If SQLNoMatch(SQL) = False Then
                MsgBox "Debe ingresar un Nro de Tramo �NICO", vbCritical + vbOKOnly, "NRO TRAMO DUPLICADO"
                .txtTramo.SetFocus
                ValidarEscalaGanancias = False
                Exit Function
            End If
        ElseIf strEditandoEscalaGanancias = "" Then
            If SQLNoMatch(SQL) = False Then
                MsgBox "Debe ingresar un Nro de Tramo �NICO", vbCritical + vbOKOnly, "NRO TRAMO DUPLICADO"
                .txtTramo.SetFocus
                ValidarEscalaGanancias = False
                Exit Function
            End If
        End If
        
        If Trim(.txtImporteMaximo.Text) = "" Or IsNumeric(.txtImporteMaximo.Text) = False Then
            MsgBox "Debe ingresar un Importe M�ximo adecuado", vbCritical + vbOKOnly, "IMPORTE M�XIMO INCORRECTO"
            .txtImporteMaximo.SetFocus
            ValidarEscalaGanancias = False
            Exit Function
        End If

        If Trim(.txtImporteFijo.Text) = "" Or IsNumeric(.txtImporteFijo.Text) = False Then
            MsgBox "Debe ingresar un Importe Fijo adecuado", vbCritical + vbOKOnly, "IMPORTE FIJO INCORRECTO"
            .txtImporteFijo.SetFocus
            ValidarEscalaGanancias = False
            Exit Function
        End If

        If Trim(.txtImporteVariable.Text) = "" Or IsNumeric(.txtImporteVariable.Text) = False Or .txtImporteVariable.Text < 1 Then
            MsgBox "Debe ingresar un Porcentaje Variable adecuado MAYOR a 1", vbCritical + vbOKOnly, "PORCENTAJE VARIABLE INCORRECTO"
            .txtImporteVariable.SetFocus
            ValidarEscalaGanancias = False
            Exit Function
        End If

    End With
    ValidarEscalaGanancias = True
    
End Function

Function ValidarLimitesDeducciones() As Boolean

    Dim SQL As String
    
    With CargaLimitesDeducciones
    
        If Trim(.txtNormaLegal.Text) = "" Or Len(.txtNormaLegal.Text) > "10" Then
            MsgBox "Debe ingresar una Norma Legal de hasta 10 caracteres", vbCritical + vbOKOnly, "NORMA LEGAL INCORRECTA"
            .txtNormaLegal.SetFocus
            ValidarLimitesDeducciones = False
            Exit Function
        End If
        If strEditandoLimitesDeducciones <> "" And .txtNormaLegal.Text <> strEditandoLimitesDeducciones Then
            If SQLNoMatch("Select * from DEDUCCIONES4TACATEGORIA Where NORMALEGAL= '" & .txtNormaLegal.Text & "'") = False Then
                MsgBox "Debe ingresar una Norma Legal �NICA", vbCritical + vbOKOnly, "NORMA LEGAL DUPLICADA"
                .txtNormaLegal.SetFocus
                ValidarLimitesDeducciones = False
                Exit Function
            End If
        ElseIf strEditandoLimitesDeducciones = "" Then
            If SQLNoMatch("Select * from DEDUCCIONES4TACATEGORIA Where NORMALEGAL= '" & .txtNormaLegal.Text & "'") = False Then
                MsgBox "Debe ingresar una Norma Legal �NICA", vbCritical + vbOKOnly, "NORMA LEGAL DUPLICADA"
                .txtNormaLegal.SetFocus
                ValidarLimitesDeducciones = False
                Exit Function
            End If
        End If
        
        If Trim(.txtFecha.Text) = "" Or IsDate(.txtFecha.Text) = False Then
            MsgBox "Debe ingresar una Fecha de la Norma Legal adecuada", vbCritical + vbOKOnly, "FECHA INCORRECTA"
            .txtFecha.SetFocus
            ValidarLimitesDeducciones = False
            Exit Function
        End If

        If Trim(.txtMinimoNoImponible.Text) = "" Or IsNumeric(.txtMinimoNoImponible.Text) = False Then
            MsgBox "Debe ingresar un monto de M�nimo no Imponible adecuado", vbCritical + vbOKOnly, "M�NIMO NO IMPONIBLE INCORRECTO"
            .txtMinimoNoImponible.SetFocus
            ValidarLimitesDeducciones = False
            Exit Function
        End If

        If Trim(.txtDeduccionEspecial.Text) = "" Or IsNumeric(.txtDeduccionEspecial.Text) = False Then
            MsgBox "Debe ingresar un monto de Deducci�n Especial adecuado", vbCritical + vbOKOnly, "DEDUCCI�N ESPECIAL INCORRECTA"
            .txtDeduccionEspecial.SetFocus
            ValidarLimitesDeducciones = False
            Exit Function
        End If

        If Trim(.txtHijo.Text) = "" Or IsNumeric(.txtHijo.Text) = False Then
            MsgBox "Debe ingresar un monto de Deducci�n por Hijo/a adecuado", vbCritical + vbOKOnly, "DEDUCCI�N POR HIJO/A INCORRECTA"
            .txtHijo.SetFocus
            ValidarLimitesDeducciones = False
            Exit Function
        End If

        If Trim(.txtConyuge.Text) = "" Or IsNumeric(.txtConyuge.Text) = False Then
            MsgBox "Debe ingresar un monto de Deducci�n por Esposa/o adecuado", vbCritical + vbOKOnly, "DEDUCCI�N POR ESPOSA/O INCORRECTA"
            .txtConyuge.SetFocus
            ValidarLimitesDeducciones = False
            Exit Function
        End If

        If Trim(.txtOtrasCargas.Text) = "" Or IsNumeric(.txtOtrasCargas.Text) = False Then
            MsgBox "Debe ingresar un monto de Deducci�n por Otras Cargas de Familia adecuado", vbCritical + vbOKOnly, "DEDUCCI�N POR OTRAS CARGAS DE FAMILIA INCORRECTA"
            .txtOtrasCargas.SetFocus
            ValidarLimitesDeducciones = False
            Exit Function
        End If

        If Trim(.txtServicioDomestico.Text) = "" Or IsNumeric(.txtServicioDomestico.Text) = False Then
            MsgBox "Debe ingresar un monto de Servicio Dom�stico adecuado", vbCritical + vbOKOnly, "SERVICIO DOM�STICO INCORRECTO"
            .txtServicioDomestico.SetFocus
            ValidarLimitesDeducciones = False
            Exit Function
        End If
        
        If Trim(.txtSeguroDeVida.Text) = "" Or IsNumeric(.txtSeguroDeVida.Text) = False Then
            MsgBox "Debe ingresar un monto de Seguro de Vida adecuado", vbCritical + vbOKOnly, "SEGURO DE VIDA INCORRECTO"
            .txtSeguroDeVida.SetFocus
            ValidarLimitesDeducciones = False
            Exit Function
        End If
        
        If Trim(.txtAlquileres.Text) = "" Or IsNumeric(.txtAlquileres.Text) = False Then
            MsgBox "Debe ingresar un monto de Alquileres adecuado", vbCritical + vbOKOnly, "ALQUILERES INCORRECTO"
            .txtAlquileres.SetFocus
            ValidarLimitesDeducciones = False
            Exit Function
        End If
        
        If Trim(.txtHonorariosMedicos.Text) = "" Or IsNumeric(.txtHonorariosMedicos.Text) = False Or .txtHonorariosMedicos.Text < 1 Then
            MsgBox "Debe ingresar un Porcentaje de Deducci�n por Honorarios M�dicos adecuado MAYOR a 1", vbCritical + vbOKOnly, "HONORARIOS M�DICOS INCORRECTO"
            .txtHonorariosMedicos.SetFocus
            ValidarLimitesDeducciones = False
            Exit Function
        End If

        If Trim(.txtCuotaMedico.Text) = "" Or IsNumeric(.txtCuotaMedico.Text) = False Or .txtCuotaMedico.Text < 1 Then
            MsgBox "Debe ingresar un Porcentaje de Deducci�n por Cuota M�dico Asistencial adecuado MAYOR a 1", vbCritical + vbOKOnly, "CUOTA M�DICO ASISTENCIAL INCORRECTA"
            .txtCuotaMedico.SetFocus
            ValidarLimitesDeducciones = False
            Exit Function
        End If
        
        If Trim(.txtDonaciones.Text) = "" Or IsNumeric(.txtDonaciones.Text) = False Or .txtDonaciones.Text < 1 Then
            MsgBox "Debe ingresar un Porcentaje de Deducci�n por Donaciones adecuado MAYOR a 1", vbCritical + vbOKOnly, "DONACIONES INCORRECTO"
            .txtDonaciones.SetFocus
            ValidarLimitesDeducciones = False
            Exit Function
        End If
        
    End With
    ValidarLimitesDeducciones = True
    
End Function

Function ValidarParentesco() As Boolean

    Dim SQL As String
    Dim strCodigo As String
    
    With CargaParentesco
    
        If Trim(.txtCodigo.Text) = "" Or IsNumeric(.txtCodigo.Text) = False Or Len(.txtCodigo.Text) > "2" Then
            MsgBox "Debe ingresar un Nro de C�digo de hasta 2 cifras", vbCritical + vbOKOnly, "NRO C�DIGO INCORRECTO"
            .txtCodigo.SetFocus
            ValidarParentesco = False
            Exit Function
        End If
        strCodigo = Format(.txtCodigo.Text, "00")
        If strEditandoParentesco <> "" And strCodigo <> strEditandoParentesco Then
            If SQLNoMatch("Select * from ASIGNACIONESFAMILIARES Where CODIGO= '" & strCodigo & "'") = False Then
                MsgBox "Debe ingresar un Nro de C�DIGO DE PARENTESCO �NICO", vbCritical + vbOKOnly, "NRO C�DIGO DUPLICADO"
                .txtCodigo.SetFocus
                ValidarParentesco = False
                Exit Function
            End If
        ElseIf strEditandoParentesco = "" Then
            If SQLNoMatch("Select * from ASIGNACIONESFAMILIARES Where CODIGO= '" & strCodigo & "'") = False Then
                MsgBox "Debe ingresar un Nro de C�DIGO DE PARENTESCO �NICO", vbCritical + vbOKOnly, "NRO C�DIGO DUPLICADO"
                .txtCodigo.SetFocus
                ValidarParentesco = False
                Exit Function
            End If
        End If
        
        If Trim(.txtDenominacion.Text) = "" Or IsNumeric(.txtDenominacion.Text) = True Or Len(.txtDenominacion.Text) > "40" Then
            MsgBox "Debe ingresar una Denomiaci�n de hasta 40 caracteres", vbCritical + vbOKOnly, "DENOMINACI�N INCORRECTA"
            .txtDenominacion.SetFocus
            ValidarParentesco = False
            Exit Function
        End If
        
        If Trim(.txtImporte.Text) = "" Or IsNumeric(.txtImporte.Text) = False Then
            MsgBox "Debe ingresar un monto de Asignaci�n Familiar adecuado", vbCritical + vbOKOnly, "IMPORTE INCORRECTO"
            .txtImporte.SetFocus
            ValidarParentesco = False
            Exit Function
        End If
        
    End With
    ValidarParentesco = True

End Function

Function ValidarDeduccionesGenerales() As Boolean

    Dim SQL As String
    
    With CargaDeduccionesGenerales
    
        If Trim(.txtFecha.Text) = "" Or IsDate(.txtFecha.Text) = False Then
            MsgBox "Debe ingresar una Fecha de alta adecuada de las Deducciones Generales", vbCritical + vbOKOnly, "FECHA INCORRECTA"
            .txtFecha.SetFocus
            ValidarDeduccionesGenerales = False
            Exit Function
        End If
        If strEditandoDeduccionesGenerales <> "" And .txtFecha.Text <> strEditandoDeduccionesGenerales Then
            If SQLNoMatch("Select * from IMPORTEDEDUCCIONESGENERALES Where PUESTOLABORAL= '" & .txtPuestoLaboral.Text & "' And FECHA = #" & Format(.txtFecha, "YYYY/MM/DD") & "#") = False Then
                MsgBox "Debe ingresar una Fecha de alta �NICA para el Agente en cuesti�n", vbCritical + vbOKOnly, "FECHA DUPLICADA"
                .txtFecha.SetFocus
                ValidarDeduccionesGenerales = False
                Exit Function
            End If
        ElseIf strEditandoDeduccionesGenerales = "" Then
            If SQLNoMatch("Select * from IMPORTEDEDUCCIONESGENERALES Where PUESTOLABORAL= '" & .txtPuestoLaboral.Text & "' And FECHA = #" & Format(.txtFecha, "YYYY/MM/DD") & "#") = False Then
                MsgBox "Debe ingresar una Norma Legal �NICA", vbCritical + vbOKOnly, "NORMA LEGAL DUPLICADA"
                .txtFecha.SetFocus
                ValidarDeduccionesGenerales = False
                Exit Function
            End If
        End If

        If Trim(.txtServicioDomestico.Text) = "" Or IsNumeric(.txtServicioDomestico.Text) = False Then
            MsgBox "Debe ingresar un monto de Servicio Dom�stico adecuado", vbCritical + vbOKOnly, "SERVICIO DOM�STICO INCORRECTO"
            .txtServicioDomestico.SetFocus
            ValidarDeduccionesGenerales = False
            Exit Function
        End If
        
        If Trim(.txtSeguroDeVida.Text) = "" Or IsNumeric(.txtSeguroDeVida.Text) = False Then
            MsgBox "Debe ingresar un monto de Seguro de Vida adecuado", vbCritical + vbOKOnly, "SEGURO DE VIDA INCORRECTO"
            .txtSeguroDeVida.SetFocus
            ValidarDeduccionesGenerales = False
            Exit Function
        End If
        
        If Trim(.txtAlquileres.Text) = "" Or IsNumeric(.txtAlquileres.Text) = False Then
            MsgBox "Debe ingresar un monto de Alquiler adecuado", vbCritical + vbOKOnly, "ALQUILER INCORRECTO"
            .txtAlquileres.SetFocus
            ValidarDeduccionesGenerales = False
            Exit Function
        End If
        
        If Trim(.txtHonorariosMedicos.Text) = "" Or IsNumeric(.txtHonorariosMedicos.Text) = False Then
            MsgBox "Debe ingresar un monto de Deducci�n por Honorarios M�dicos adecuado", vbCritical + vbOKOnly, "HONORARIOS M�DICOS INCORRECTO"
            .txtHonorariosMedicos.SetFocus
            ValidarDeduccionesGenerales = False
            Exit Function
        End If

        If Trim(.txtCuotaMedico.Text) = "" Or IsNumeric(.txtCuotaMedico.Text) = False Then
            MsgBox "Debe ingresar un monto de Deducci�n por Cuota M�dico Asistencial adecuado", vbCritical + vbOKOnly, "CUOTA M�DICO ASISTENCIAL INCORRECTA"
            .txtCuotaMedico.SetFocus
            ValidarDeduccionesGenerales = False
            Exit Function
        End If
        
        If Trim(.txtDonaciones.Text) = "" Or IsNumeric(.txtDonaciones.Text) = False Then
            MsgBox "Debe ingresar un monto de Deducci�n por Donaciones adecuado", vbCritical + vbOKOnly, "DONACIONES INCORRECTO"
            .txtDonaciones.SetFocus
            ValidarDeduccionesGenerales = False
            Exit Function
        End If
        
    End With
    ValidarDeduccionesGenerales = True
    
End Function

Function ValidarCodigoLiquidacion() As Boolean

    Dim SQL As String
    
    With CargaCodigoLiquidacion
    
        If Trim(.txtCodigo.Text) = "" Or IsNumeric(.txtCodigo.Text) = False Or Len(.txtCodigo.Text) > "4" Then
            MsgBox "Debe ingresar un Nro de C�digo de hasta 4 cifras", vbCritical + vbOKOnly, "NRO C�DIGO INCORRECTO"
            .txtCodigo.SetFocus
            ValidarCodigoLiquidacion = False
            Exit Function
        End If
        strCodigo = .txtCodigo.Text
        If strEditandoCodigoLiquidacion <> "" And strCodigo <> strEditandoCodigoLiquidacion Then
            If SQLNoMatch("Select * from CODIGOLIQUIDACIONES Where CODIGO= '" & strCodigo & "'") = False Then
                MsgBox "Debe ingresar un Nro de C�DIGO DE LIQUIDACI�N �NICO", vbCritical + vbOKOnly, "NRO C�DIGO DUPLICADO"
                .txtCodigo.SetFocus
                ValidarCodigoLiquidacion = False
                Exit Function
            End If
        ElseIf strEditandoCodigoLiquidacion = "" Then
            If SQLNoMatch("Select * from ASIGNACIONESFAMILIARES Where CODIGO= '" & strCodigo & "'") = False Then
                MsgBox "Debe ingresar un Nro de C�DIGO DE LIQUIDACI�N �NICO", vbCritical + vbOKOnly, "NRO C�DIGO DUPLICADO"
                .txtCodigo.SetFocus
                ValidarCodigoLiquidacion = False
                Exit Function
            End If
        End If
        
        If Trim(.txtPeriodo.Text) = "" Or IsNumeric(Left(.txtPeriodo.Text, 2)) = False Or IsNumeric(Right(.txtPeriodo.Text, 4)) = False Or Mid(.txtPeriodo.Text, 3, 1) <> "/" Then
            MsgBox "Debe ingresar un Per�odo de Liquidaci�n acorde al Formato mm/aaaa", vbCritical + vbOKOnly, "PER�ODO LIQUIDACI�N INCORRECTO"
            .txtPeriodo.SetFocus
            ValidarCodigoLiquidacion = False
            Exit Function
        End If
        
        If Trim(.txtDescripcion.Text) = "" Or IsNumeric(.txtDescripcion.Text) = True Or Len(.txtDescripcion.Text) > "20" Then
            MsgBox "Debe ingresar una Descripci�n de hasta 20 caracteres", vbCritical + vbOKOnly, "DESCRIPCI�N INCORRECTA"
            .txtDescripcion.SetFocus
            ValidarCodigoLiquidacion = False
            Exit Function
        End If
        
        If Trim(.txtMontoExento.Text) = "" Or IsNumeric(.txtMontoExento.Text) = False Then
            MsgBox "Debe ingresar un Monto Exento adecuado", vbCritical + vbOKOnly, "MONTO EXENTO INCORRECTO"
            .txtMontoExento.SetFocus
            ValidarCodigoLiquidacion = False
            Exit Function
        End If
        
    End With
    ValidarCodigoLiquidacion = True

End Function

Function ValidarFamiliar() As Boolean

    Dim SQL As String
    Dim strFecha As String
    Dim datFecha As Date
    
    With CargaFamiliar
    
        If Trim(.txtDNI.Text) = "" Or IsNumeric(.txtDNI.Text) = False Or Len(.txtDNI.Text) > "8" Then
            MsgBox "Debe ingresar un Nro de DNI de hasta 8 cifras", vbCritical + vbOKOnly, "NRO DNI INCORRECTO"
            .txtDNI.SetFocus
            ValidarFamiliar = False
            Exit Function
        End If
        If strEditandoFamiliar <> "" And .txtDNI.Text <> strEditandoFamiliar Then
            If SQLNoMatch("Select * from CARGASDEFAMILIA Where PUESTOLABORAL= '" & .txtPuestoLaboral.Text & "' and DNI= '" & .txtDNI.Text & "'") = False Then
                MsgBox "Debe ingresar un Nro de DNI �NICO", vbCritical + vbOKOnly, "NRO DNI DUPLICADO"
                .txtDNI.SetFocus
                ValidarFamiliar = False
                Exit Function
            End If
        ElseIf strEditandoFamiliar = "" Then
            If SQLNoMatch("Select * from CARGASDEFAMILIA Where PUESTOLABORAL= '" & .txtPuestoLaboral.Text & "' and DNI= '" & .txtDNI.Text & "'") = False Then
                MsgBox "Debe ingresar un Nro de DNI �NICO", vbCritical + vbOKOnly, "NRO DNI DUPLICADO"
                .txtDNI.SetFocus
                ValidarFamiliar = False
                Exit Function
            End If
        End If
        
        If Trim(.txtDescripcionFamiliar.Text) = "" Or IsNumeric(.txtDescripcionFamiliar.Text) = True Or Len(.txtDescripcionFamiliar.Text) > "100" Then
            MsgBox "Debe ingresar el Nombre Completo del Familiar de hasta 100 caracteres", vbCritical + vbOKOnly, "NOMBRE FAMILIAR INCORRECTO"
            .txtDescripcionFamiliar.SetFocus
            ValidarFamiliar = False
            Exit Function
        End If
        
        If SQLNoMatch("Select * from ASIGNACIONESFAMILIARES Where PARENTESCO= '" & .cmbParentesco.Text & "'") = True Then
            MsgBox "Debe ingresar un PARENTESCO de la lista desplegable", vbCritical + vbOKOnly, "PARENTESCO INEXISTENTE"
            .cmbParentesco.SetFocus
            ValidarFamiliar = False
            Exit Function
        End If
        
        strFecha = .txtFechaAlta.Text
        If IsDate(strFecha) = False Then
            MsgBox "Debe Ingresarse una Fecha de Alta V�lida", vbCritical + vbOKOnly, "FECHA ALTA INV�LIDA"
            .txtFechaAlta.SetFocus
            ValidarFamiliar = False
            Exit Function
        End If
        
        datFecha = DateTime.DateSerial(Right(strFecha, 4), Mid(strFecha, 4, 2), Left(strFecha, 2))
        
        If datFecha > Date Then
            MsgBox "La Fecha de Alta del Familiar no puede ser superior a la actual", vbCritical + vbOKOnly, "FECHA ALTA INV�LIDA"
            .txtFechaAlta.SetFocus
            ValidarFamiliar = False
            Exit Function
        End If

        If .cmbNivelDeEstudio.Text <> "Sin Estudios" And .cmbNivelDeEstudio.Text <> "Primario" And .cmbNivelDeEstudio.Text <> "Secundario" And .cmbNivelDeEstudio.Text <> "Universitario" Then
            MsgBox "Debe ingresar un NIVEL DE ESTUDIO de la lista desplegable", vbCritical + vbOKOnly, "NIVEL DE ESTUDIO INCORRECTO"
            .cmbNivelDeEstudio.SetFocus
            ValidarFamiliar = False
            Exit Function
        End If
        
    End With
    ValidarFamiliar = True
    
End Function

Function ValidarLiquidacionGanancias() As Boolean

    Dim SQL As String
    
    With LiquidacionGanancia4ta

        If Trim(.txtHaberOptimo.Text) = "" Or IsNumeric(.txtHaberOptimo.Text) = False Then
            MsgBox "Debe ingresar un monto de Haber �ptimo adecuado", vbCritical + vbOKOnly, "HABER �PTIMO INCORRECTO"
            .txtHaberOptimo.SetFocus
            ValidarLiquidacionGanancias = False
            Exit Function
        End If
        
        If Trim(.txtPluriempleo.Text) = "" Or IsNumeric(.txtPluriempleo.Text) = False Then
            MsgBox "Debe ingresar un monto de Pluriempleo adecuado", vbCritical + vbOKOnly, "PLURIEMPLEO INCORRECTO"
            .txtPluriempleo.SetFocus
            ValidarLiquidacionGanancias = False
            Exit Function
        End If
        
        If Trim(.txtAjuste.Text) = "" Or IsNumeric(.txtAjuste.Text) = False Then
            MsgBox "Debe ingresar un monto de Ajuste de Renta Imponible adecuado", vbCritical + vbOKOnly, "AJUSTE DE RENTA IMPONIBLE INCORRECTO"
            .txtAjuste.SetFocus
            ValidarLiquidacionGanancias = False
            Exit Function
        End If

        If Trim(.txtJubilacion.Text) = "" Or IsNumeric(.txtJubilacion.Text) = False Then
            MsgBox "Debe ingresar un monto de Jubilaci�n adecuado", vbCritical + vbOKOnly, "JUBILACI�N INCORRECTA"
            .txtJubilacion.SetFocus
            ValidarLiquidacionGanancias = False
            Exit Function
        End If
        
        If Trim(.txtObraSocial.Text) = "" Or IsNumeric(.txtObraSocial.Text) = False Then
            MsgBox "Debe ingresar un monto de Obra Social adecuado", vbCritical + vbOKOnly, "OBRA SOCIAL INCORRECTA"
            .txtObraSocial.SetFocus
            ValidarLiquidacionGanancias = False
            Exit Function
        End If
        
        If Trim(.txtAdherente.Text) = "" Or IsNumeric(.txtAdherente.Text) = False Then
            MsgBox "Debe ingresar un monto de Adherente adecuado", vbCritical + vbOKOnly, "ADHERENTE INCORRECTO"
            .txtAdherente.SetFocus
            ValidarLiquidacionGanancias = False
            Exit Function
        End If
        
        If Trim(.txtSeguroObligatorio.Text) = "" Or IsNumeric(.txtSeguroObligatorio.Text) = False Then
            MsgBox "Debe ingresar un monto de Seguro Obligatorio adecuado", vbCritical + vbOKOnly, "SEGURO OBLIGATORIO INCORRECTO"
            .txtSeguroObligatorio.SetFocus
            ValidarLiquidacionGanancias = False
            Exit Function
        End If
        
        If Trim(.txtCuotaSindical.Text) = "" Or IsNumeric(.txtCuotaSindical.Text) = False Then
            MsgBox "Debe ingresar un monto de Cuota Sindical adecuado", vbCritical + vbOKOnly, "CUOTA SINDICAL INCORRECTO"
            .txtCuotaSindical.SetFocus
            ValidarLiquidacionGanancias = False
            Exit Function
        End If
        
        If Trim(.txtSeguroOptativo.Text) = "" Or IsNumeric(.txtSeguroOptativo.Text) = False Then
            MsgBox "Debe ingresar un monto de Seguro Optativo adecuado", vbCritical + vbOKOnly, "SEGURO OPTATIVO INCORRECTO"
            .txtSeguroOptativo.SetFocus
            ValidarLiquidacionGanancias = False
            Exit Function
        End If
                
        If Trim(.txtAjuesteRetencion.Text) = "" Or IsNumeric(.txtAjuesteRetencion.Text) = False Then
            MsgBox "Debe ingresar un monto de Ajuste de Retenci�n adecuado", vbCritical + vbOKOnly, "AJUSTE DE RETENCI�N INCORRECTO"
            .txtAjuesteRetencion.SetFocus
            ValidarLiquidacionGanancias = False
            Exit Function
        End If
        
    End With
    ValidarLiquidacionGanancias = True
    
End Function

Function ValidarLiquidacionGananciasSIRADIG() As Boolean

    Dim SQL As String
    
    With LiquidacionGanancia4taSIRADIG

        If Trim(.txtHaberOptimo.Text) = "" Or IsNumeric(.txtHaberOptimo.Text) = False Then
            MsgBox "Debe ingresar un monto de Haber �ptimo adecuado", vbCritical + vbOKOnly, "HABER �PTIMO INCORRECTO"
            .txtHaberOptimo.SetFocus
            ValidarLiquidacionGananciasSIRADIG = False
            Exit Function
        End If
        
        If Trim(.txtPluriempleo.Text) = "" Or IsNumeric(.txtPluriempleo.Text) = False Then
            MsgBox "Debe ingresar un monto de Pluriempleo adecuado", vbCritical + vbOKOnly, "PLURIEMPLEO INCORRECTO"
            .txtPluriempleo.SetFocus
            ValidarLiquidacionGananciasSIRADIG = False
            Exit Function
        End If
        
        If Trim(.txtAjuste.Text) = "" Or IsNumeric(.txtAjuste.Text) = False Then
            MsgBox "Debe ingresar un monto de Ajuste de Renta Imponible adecuado", vbCritical + vbOKOnly, "AJUSTE DE RENTA IMPONIBLE INCORRECTO"
            .txtAjuste.SetFocus
            ValidarLiquidacionGananciasSIRADIG = False
            Exit Function
        End If

        If Trim(.txtJubilacion.Text) = "" Or IsNumeric(.txtJubilacion.Text) = False Then
            MsgBox "Debe ingresar un monto de Jubilaci�n adecuado", vbCritical + vbOKOnly, "JUBILACI�N INCORRECTA"
            .txtJubilacion.SetFocus
            ValidarLiquidacionGananciasSIRADIG = False
            Exit Function
        End If
        
        If Trim(.txtObraSocial.Text) = "" Or IsNumeric(.txtObraSocial.Text) = False Then
            MsgBox "Debe ingresar un monto de Obra Social adecuado", vbCritical + vbOKOnly, "OBRA SOCIAL INCORRECTA"
            .txtObraSocial.SetFocus
            ValidarLiquidacionGananciasSIRADIG = False
            Exit Function
        End If
        
        If Trim(.txtAdherente.Text) = "" Or IsNumeric(.txtAdherente.Text) = False Then
            MsgBox "Debe ingresar un monto de Adherente adecuado", vbCritical + vbOKOnly, "ADHERENTE INCORRECTO"
            .txtAdherente.SetFocus
            ValidarLiquidacionGananciasSIRADIG = False
            Exit Function
        End If
        
        If Trim(.txtSeguroObligatorio.Text) = "" Or IsNumeric(.txtSeguroObligatorio.Text) = False Then
            MsgBox "Debe ingresar un monto de Seguro Obligatorio adecuado", vbCritical + vbOKOnly, "SEGURO OBLIGATORIO INCORRECTO"
            .txtSeguroObligatorio.SetFocus
            ValidarLiquidacionGananciasSIRADIG = False
            Exit Function
        End If
        
        If Trim(.txtCuotaSindical.Text) = "" Or IsNumeric(.txtCuotaSindical.Text) = False Then
            MsgBox "Debe ingresar un monto de Cuota Sindical adecuado", vbCritical + vbOKOnly, "CUOTA SINDICAL INCORRECTO"
            .txtCuotaSindical.SetFocus
            ValidarLiquidacionGananciasSIRADIG = False
            Exit Function
        End If
        
        If Trim(.txtSeguroOptativo.Text) = "" Or IsNumeric(.txtSeguroOptativo.Text) = False Then
            MsgBox "Debe ingresar un monto de Seguro Optativo adecuado", vbCritical + vbOKOnly, "SEGURO OPTATIVO INCORRECTO"
            .txtSeguroOptativo.SetFocus
            ValidarLiquidacionGananciasSIRADIG = False
            Exit Function
        End If
                
        If Trim(.txtAjuesteRetencion.Text) = "" Or IsNumeric(.txtAjuesteRetencion.Text) = False Then
            MsgBox "Debe ingresar un monto de Ajuste de Retenci�n adecuado", vbCritical + vbOKOnly, "AJUSTE DE RETENCI�N INCORRECTO"
            .txtAjuesteRetencion.SetFocus
            ValidarLiquidacionGananciasSIRADIG = False
            Exit Function
        End If
        
    End With
    ValidarLiquidacionGananciasSIRADIG = True
    
End Function

Function ValidarGenerarF649() As Boolean

    Dim SQL As String
    
    With LiquidacionFinalGanancias

        If Trim(.txtPuestoLaboral.Text) = "" Or IsNumeric(.txtPuestoLaboral.Text) = False Then
            MsgBox "Debe ingresar un Nro de Puesto Laboral adecuado", vbCritical + vbOKOnly, "PUESTO LABORAL INCORRECTO"
            .txtPuestoLaboral.SetFocus
            ValidarGenerarF649 = False
            Exit Function
        End If
        
        If Trim(.txtPeriodo.Text) = "" Or IsNumeric(.txtPeriodo.Text) = False Or Len(.txtPeriodo.Text) <> "4" Then
            MsgBox "Debe ingresar un A�o adecuado", vbCritical + vbOKOnly, "A�O INCORRECTO"
            .txtPeriodo.SetFocus
            ValidarGenerarF649 = False
            Exit Function
        End If
        
        SQL = "Select * From CODIGOLIQUIDACIONES Inner Join LIQUIDACIONSUELDOS On CODIGOLIQUIDACIONES.Codigo = LIQUIDACIONSUELDOS.CodigoLiquidacion " _
        & "Where PUESTOLABORAL = '" & .txtPuestoLaboral.Text & "' And Right(PERIODO,4) = '" & .txtPeriodo.Text & "'"
        If SQLNoMatch(SQL) = True Then
            MsgBox "Debe ingresar un A�o decuado", vbCritical + vbOKOnly, "A�O INCORRECTO"
            .txtPeriodo.SetFocus
            ValidarGenerarF649 = False
            Exit Function
        End If

        If Trim(.txtDescripcion.Text) = "" Or IsNumeric(.txtDescripcion.Text) = True Then
            MsgBox "Debe ingresar una Descripci�n del Agente adecuada", vbCritical + vbOKOnly, "DESCRIPCI�N DEL AGENTE INCORRECTO"
            .txtDescripcion.SetFocus
            ValidarGenerarF649 = False
            Exit Function
        End If
        
    End With
    ValidarGenerarF649 = True
    
End Function

Function ValidarComprobanteSIIF() As Boolean

    Dim strValidar As String

    With CargaComprobanteSIIF
    
        If Trim(.txtComprobante.Text) = "" Or IsNumeric(.txtComprobante.Text) = False Or Len(.txtComprobante.Text) > "5" Then
            MsgBox "Debe ingresar un Nro de Comprobante de hasta 5 cifras", vbCritical + vbOKOnly, "NRO COMPROBANTE INCORRECTO"
            .txtComprobante.SetFocus
            ValidarComprobanteSIIF = False
            Exit Function
        End If
        
        If Left(.txtFecha.Text, 2) = "" Or IsDate(.txtFecha.Text) = False Or Len(.txtFecha.Text) <> "10" Then
            MsgBox "Debe ingresar una fecha valida", vbCritical + vbOKOnly, "FECHA INCORRECTA"
            .txtFecha.SetFocus
            ValidarComprobanteSIIF = False
            Exit Function
        End If

        strValidar = Format(.txtComprobante.Text, "00000") & "/" & Right(.txtFecha.Text, 2)
        If SQLNoMatch("Select * from LIQUIDACIONHONORARIOS Where COMPROBANTE= '" & strValidar & "'") = False Then
            MsgBox "El Nro de Comprobante ya existe, verifique el Nro", vbCritical + vbOKOnly, "COMPROBANTE DUPLICADO"
            .txtComprobante.SetFocus
            ValidarComprobanteSIIF = False
            Exit Function
        End If
        strValidar = ""

        strValidar = .txtImputacion.Text
        If strValidar = "" Or EsIgualTextoEspecificado(strValidar, "Honorarios", "Comisiones", "Horas Extras", "Licencia") = False Then
            MsgBox "El Dato Ingresado es incorrecto por no encontrarse en la LISTA ESPECIFICADA", vbCritical + vbOKOnly, "TIPO DE IMPUTACI�N INCORRECTA"
            .txtImputacion.SetFocus
            ValidarComprobanteSIIF = False
            Exit Function
        End If
        strValidar = ""
    
    End With
    
    ValidarComprobanteSIIF = True
    
End Function

Public Function EsIgualTextoEspecificado(TextoAnalizado As String, ParamArray Valores() As Variant) As Boolean
    
    Dim Valor As Variant
    For Each Valor In Valores
        If Valor = TextoAnalizado Then
            EsIgualTextoEspecificado = True
            Exit Function
        End If
    Next Valor
    'MsgBox "El Dato Ingresado es incorrecto por no encontrarse en la LISTA ESPECIFICADA" & vbCrLf & "INTENTE DE NUEVO", vbCritical + vbOKOnly, NombredelCampo & " INCORRECTO"
    EsIgualTextoEspecificado = False
    
End Function

Function ValidarPrecarizado() As Boolean

    With CargaPrecarizado
    
        If Trim(.txtNombreCompleto.Text) = "" Or IsNumeric(.txtNombreCompleto.Text) = True Or Len(.txtNombreCompleto.Text) > 50 Then
            MsgBox "Debe ingresar un Nombre Completo adecuado de hasta 50 caracteres", vbCritical + vbOKOnly, "NOMBRE COMPLETO INCORRECTO"
            .txtNombreCompleto.SetFocus
            ValidarPrecarizado = False
            Exit Function
        End If
        If strEditandoPrecarizado <> "" And .txtNombreCompleto.Text <> strEditandoPrecarizado Then
            If SQLNoMatch("Select * from PRECARIZADOS Where AGENTES= '" & .txtNombreCompleto.Text & "'") = False Then
                MsgBox "Debe ingresar un Nombre Completo �NICO", vbCritical + vbOKOnly, "NOMBRE PRECARIZADO DUPLICADO"
                .txtNombreCompleto.SetFocus
                ValidarPrecarizado = False
                Exit Function
            End If
        ElseIf strEditandoPrecarizado = "" Then
            If SQLNoMatch("Select * from PRECARIZADOS Where AGENTES= '" & .txtNombreCompleto.Text & "'") = False Then
                MsgBox "Debe ingresar un Nombre Completo �NICO", vbCritical + vbOKOnly, "NOMBRE PRECARIZADO DUPLICADO"
                .txtNombreCompleto.SetFocus
                ValidarPrecarizado = False
                Exit Function
            End If
        End If
        
        
        If Trim(.mskEstructura.Text) = "" Or Len(.mskEstructura.Text) <> 12 Then
            MsgBox "Debe ingresar una Estructura Presupuestaria correcta con el siguiente formato: 00-00-00-000 (no se inclye subprograma)", vbCritical + vbOKOnly, "ESTRUCTURA PRESUPUESTARIA INCORRECTA"
            .mskEstructura.SetFocus
            ValidarPrecarizado = False
            Exit Function
        End If
        If IsNumeric(Left(.mskEstructura.Text, 2)) = False Or IsNumeric(Mid(.mskEstructura.Text, 4, 2)) = False Or IsNumeric(Mid(.mskEstructura.Text, 7, 2)) = False Or IsNumeric(Right(.mskEstructura.Text, 3)) = False Then
            MsgBox "Debe ingresar una Estructura Presupuestaria correcta con el siguiente formato: 00-00-00-000 (no se inclye subprograma)", vbCritical + vbOKOnly, "ESTRUCTURA PRESUPUESTARIA INCORRECTA"
            .mskEstructura.SetFocus
            ValidarPrecarizado = False
            Exit Function
        End If
    
    End With
    
    ValidarPrecarizado = True
    
End Function

Function ValidarEstructuraPresupuestaria(strEstructura As String) As Boolean

    If Trim(strEstructura) = "" Or Len(strEstructura) <> 12 Then
        ValidarEstructuraPresupuestaria = False
        Exit Function
    End If
       
    If IsNumeric(Left(strEstructura, 2)) = False Or IsNumeric(Mid(strEstructura, 4, 2)) = False Or IsNumeric(Mid(strEstructura, 7, 2)) = False Or IsNumeric(Right(strEstructura, 3)) = False Then
        ValidarEstructuraPresupuestaria = False
        Exit Function
    End If
    
    If Left(strEstructura, 2) = "00" Or Mid(strEstructura, 7, 2) = "00" Or Right(strEstructura, 3) = "000" Then
        ValidarEstructuraPresupuestaria = False
        Exit Function
    End If
    
    ValidarEstructuraPresupuestaria = True
    
End Function

Function ValidarCmbPeriodoResumenAnualGanancias() As Boolean

    Dim SQL As String
    
    With ResumenAnualGanancias.cmbPeriodo
    
        If Trim(.Text) = "" Or IsNumeric(.Text) = False Or Len(.Text) > "4" Then
            MsgBox "Debe ingresar un Per�odo de Liquidacion del Listado", vbCritical + vbOKOnly, "PER�ODO INCORRECTO"
            .SetFocus
            ValidarCmbPeriodoResumenAnualGanancias = False
            Exit Function
        End If
            
        SQL = "Select Right(PERIODO,4) As PeriodoLiquidacion From" _
        & " LIQUIDACIONGANANCIAS4TACATEGORIA Inner Join CODIGOLIQUIDACIONES On" _
        & " LIQUIDACIONGANANCIAS4TACATEGORIA.CodigoLiquidacion = CODIGOLIQUIDACIONES.Codigo" _
        & " Where Right(PERIODO,4) = '" & .Text & "'" _
        & " Group by Right(PERIODO,4)"
        If SQLNoMatch(SQL) Then
            MsgBox "Debe ingresar un Per�odo de Liquidacion del Listado", vbCritical + vbOKOnly, "PER�ODO INCORRECTO"
            .SetFocus
            ValidarCmbPeriodoResumenAnualGanancias = False
            Exit Function
        End If
        
    End With
    ValidarCmbPeriodoResumenAnualGanancias = True

End Function

Function ValidarActualizarResumenAnualGanancias() As Boolean

    Dim SQL As String
    
    With ResumenAnualGanancias.cmbPeriodo
    
        If Trim(.Text) = "" Or IsNumeric(.Text) = False Or Len(.Text) > "4" Then
            MsgBox "Debe ingresar un Per�odo de Liquidacion del Listado", vbCritical + vbOKOnly, "PER�ODO INCORRECTO"
            .SetFocus
            ValidarActualizarResumenAnualGanancias = False
            Exit Function
        End If
            
        SQL = "Select Right(PERIODO,4) As PeriodoLiquidacion From" _
        & " LIQUIDACIONGANANCIAS4TACATEGORIA Inner Join CODIGOLIQUIDACIONES On" _
        & " LIQUIDACIONGANANCIAS4TACATEGORIA.CodigoLiquidacion = CODIGOLIQUIDACIONES.Codigo" _
        & " Where Right(PERIODO,4) = '" & .Text & "'" _
        & " Group by Right(PERIODO,4)"
        If SQLNoMatch(SQL) Then
            MsgBox "Debe ingresar un Per�odo de Liquidacion del Listado", vbCritical + vbOKOnly, "PER�ODO INCORRECTO"
            .SetFocus
            ValidarActualizarResumenAnualGanancias = False
            Exit Function
        End If
        
    End With
    
    With ResumenAnualGanancias
    
        If Trim(.cmbAgente.Text) = "" Then
            MsgBox "Debe ingresar un Agente Liquidado del Listado", vbCritical + vbOKOnly, "AGENTE INCORRECTO"
            .cmbAgente.SetFocus
            ValidarActualizarResumenAnualGanancias = False
            Exit Function
        End If
            
        SQL = "Select NombreCompleto From" _
        & " (LIQUIDACIONGANANCIAS4TACATEGORIA Inner Join CODIGOLIQUIDACIONES On" _
        & " LIQUIDACIONGANANCIAS4TACATEGORIA.CodigoLiquidacion = CODIGOLIQUIDACIONES.Codigo)" _
        & " Inner Join AGENTES On" _
        & " LIQUIDACIONGANANCIAS4TACATEGORIA.PuestoLaboral = AGENTES.PuestoLaboral" _
        & " Where Right(PERIODO,4) = '" & .cmbPeriodo.Text & "'" _
        & " And NombreCompleto = '" & .cmbAgente.Text & "'" _
        & " Group by NombreCompleto" _
        & " Order by NombreCompleto Asc"
        If SQLNoMatch(SQL) Then
            MsgBox "Debe ingresar un Agente Liquidado del Listado", vbCritical + vbOKOnly, "AGENTE INCORRECTO"
            .cmbAgente.SetFocus
            ValidarActualizarResumenAnualGanancias = False
            Exit Function
        End If
        
    End With
    
    ValidarActualizarResumenAnualGanancias = True

End Function

Function ValidarActualizarReciboDeSueldo() As Boolean

    Dim SQL As String
    
    With ReciboDeSueldo.cmbPeriodo
        
        Dim strCodigoLiquidacion As String
        strCodigoLiquidacion = Left(.Text, 4)
        
        If Trim(strCodigoLiquidacion) = "" Or IsNumeric(strCodigoLiquidacion) = False Then
            MsgBox "Debe ingresar un Per�odo de Liquidacion del Listado", vbCritical + vbOKOnly, "PER�ODO INCORRECTO"
            .SetFocus
            ValidarActualizarReciboDeSueldo = False
            Exit Function
        End If
            
        SQL = "Select CODIGO From CODIGOLIQUIDACIONES" _
        & " Where CODIGO = " & "'" & strCodigoLiquidacion & "'"
        If SQLNoMatch(SQL) Then
            MsgBox "Debe ingresar un Per�odo de Liquidacion del Listado", vbCritical + vbOKOnly, "PER�ODO INCORRECTO"
            .SetFocus
            ValidarActualizarReciboDeSueldo = False
            Exit Function
        End If
        
    End With
    
    With ReciboDeSueldo.cmbAgente
    
        If Trim(.Text) = "" Then
            MsgBox "Debe ingresar un Agente Liquidado del Listado", vbCritical + vbOKOnly, "AGENTE INCORRECTO"
            .SetFocus
            ValidarActualizarReciboDeSueldo = False
            Exit Function
        End If
            
        SQL = "Select NombreCompleto From AGENTES" _
        & " Where NombreCompleto = " & "'" & .Text & "'"
        If SQLNoMatch(SQL) Then
            MsgBox "Debe ingresar un Agente Liquidado del Listado", vbCritical + vbOKOnly, "AGENTE INCORRECTO"
            .SetFocus
            ValidarActualizarReciboDeSueldo = False
            Exit Function
        End If
        
    End With
    
    ValidarActualizarReciboDeSueldo = True

End Function

Function ValidarCargaHaberLiquidado() As Boolean

    Dim SQL As String
    
    With CargaHaberLiquidado.cmbConcepto
        
        Dim strCodigoConcepto As String
        strCodigoConcepto = Left(.Text, 4)
        
        If Trim(strCodigoConcepto) = "" Or IsNumeric(strCodigoConcepto) = False Then
            MsgBox "Debe ingresar un Concepto del Listado", vbCritical + vbOKOnly, "CONCEPTO INCORRECTO"
            .SetFocus
            ValidarCargaHaberLiquidado = False
            Exit Function
        End If
            
        SQL = "Select CODIGO From CONCEPTOS" _
        & " Where CODIGO = " & "'" & strCodigoConcepto & "'"
        If SQLNoMatch(SQL) Then
            MsgBox "Debe ingresar un Concepto del Listado", vbCritical + vbOKOnly, "CONCEPTO INCORRECTO"
            .SetFocus
            ValidarCargaHaberLiquidado = False
            Exit Function
        End If
        
        SQL = "Select CODIGOCONCEPTO From LIQUIDACIONSUELDOS" _
        & " Where CODIGOCONCEPTO = " & "'" & strCodigoConcepto & "'" _
        & " And CODIGOLIQUIDACION = " & "'" & Left(CargaHaberLiquidado.txtPeriodo.Text, 4) & "'" _
        & " And PUESTOLABORAL = " & "'" & CargaHaberLiquidado.txtPuestoLaboral.Text & "'"
        If SQLNoMatch(SQL) = False Then
            MsgBox "Prestar atenci�n a que el concepto ya no se encuentre liquidado", vbCritical + vbOKOnly, "CONCEPTO YA LIQUIDADO"
            .SetFocus
            ValidarCargaHaberLiquidado = False
            Exit Function
        End If
        
    End With
    
    With CargaHaberLiquidado.txtImporte
            
        If Trim(.Text) = "" Or IsNumeric(.Text) = False Then
            MsgBox "Debe ingresar un Importe adecuado", vbCritical + vbOKOnly, "IMPORTE INCORRECTO"
            .SetFocus
            ValidarCargaHaberLiquidado = False
            Exit Function
        End If
    
    End With

    ValidarCargaHaberLiquidado = True

End Function

Function ValidarPrecarizadoImputado() As Boolean

    With CargaPrecarizadoImputado
    
        If Trim(.cmbNombreCompleto.Text) = "" Or IsNumeric(.cmbNombreCompleto.Text) = True Or Len(.cmbNombreCompleto.Text) > 50 Then
            MsgBox "Debe ingresar un Nombre Completo adecuado de hasta 50 caracteres", vbCritical + vbOKOnly, "NOMBRE COMPLETO INCORRECTO"
            .cmbNombreCompleto.SetFocus
            ValidarPrecarizadoImputado = False
            Exit Function
        End If
      
        If Trim(.mskEstructuraImputada.Text) = "" Or Len(.mskEstructuraImputada.Text) <> 12 Then
            MsgBox "Debe ingresar una Estructura Presupuestaria correcta con el siguiente formato: 00-00-00-000 (no se inclye subprograma)", vbCritical + vbOKOnly, "ESTRUCTURA PRESUPUESTARIA INCORRECTA"
            .mskEstructuraImputada.SetFocus
            ValidarPrecarizadoImputado = False
            Exit Function
        End If
        If IsNumeric(Left(.mskEstructuraImputada.Text, 2)) = False Or IsNumeric(Mid(.mskEstructuraImputada.Text, 4, 2)) = False Or IsNumeric(Mid(.mskEstructuraImputada.Text, 7, 2)) = False Or IsNumeric(Right(.mskEstructuraImputada.Text, 3)) = False Then
            MsgBox "Debe ingresar una Estructura Presupuestaria correcta con el siguiente formato: 00-00-00-000 (no se inclye subprograma)", vbCritical + vbOKOnly, "ESTRUCTURA PRESUPUESTARIA INCORRECTA"
            .mskEstructuraImputada.SetFocus
            ValidarPrecarizadoImputado = False
            Exit Function
        End If
    
        If Trim(.txtMontoBruto.Text) = "" Or IsNumeric(.txtMontoBruto.Text) = False Then
            MsgBox "Debe ingresar Monto de Factura adecuado", vbCritical + vbOKOnly, "MONTO BRUTO INCORRECTO"
            .txtMontoBruto.SetFocus
            ValidarPrecarizadoImputado = False
            Exit Function
        End If
    
    End With
    
    ValidarPrecarizadoImputado = True
    
End Function

Function ValidarCodigoSIRADIG(ValidarTabla As String) As Boolean

    Dim SQL As String

    With CargaCodigoSIRADIG
    
        If Trim(.txtCodigo.Text) = "" Or IsNumeric(.txtCodigo.Text) = False Or Len(.txtCodigo.Text) > "2" Then
            MsgBox "Debe ingresar un Nro de C�digo de hasta 2 cifras", vbCritical + vbOKOnly, "NRO C�DIGO INCORRECTO"
            .txtCodigo.SetFocus
            ValidarCodigoSIRADIG = False
            Exit Function
        End If
        If strEditandoConcepto <> "" And .txtCodigo.Text <> strEditandoCodigoSIRADIG Then
            If SQLNoMatch("Select * from " & ValidarTabla & " Where CODIGO= '" & .txtCodigo.Text & "'") = False Then
                MsgBox "Debe ingresar un Nro de C�DIGO �NICO", vbCritical + vbOKOnly, "NRO C�DIGO DUPLICADO"
                .txtCodigo.SetFocus
                ValidarCodigoSIRADIG = False
                Exit Function
            End If
        ElseIf strEditandoConcepto = "" Then
            If SQLNoMatch("Select * from " & ValidarTabla & " Where CODIGO= '" & .txtCodigo.Text & "'") = False Then
                MsgBox "Debe ingresar un Nro de C�DIGO �NICO", vbCritical + vbOKOnly, "NRO C�DIGO DUPLICADO"
                .txtCodigo.SetFocus
                ValidarCodigoSIRADIG = False
                Exit Function
            End If
        End If
        
        If Trim(.txtDenominacion.Text) = "" Or IsNumeric(.txtDenominacion.Text) = True Or Len(.txtDenominacion.Text) > "30" Then
            MsgBox "Debe ingresar una Denomiaci�n de hasta 30 caracteres", vbCritical + vbOKOnly, "DENOMINACI�N INCORRECTA"
            .txtDenominacion.SetFocus
            ValidarCodigoSIRADIG = False
            Exit Function
        End If
        
    End With
    ValidarCodigoSIRADIG = True

End Function

Function ValidarMigrarDeducciones(PeriodoOrigen As String, _
PeriodoDestino As String, TipoDato As String) As Boolean

    Dim SQL As String

    With MigrarDeducciones
    
        SQL = "Select Right(ID,2) as Periodo " _
        & "From PresentacionSIRADIG " _
        & "Where Right(ID,2) = '" & Right(PeriodoOrigen, 2) & "' " _
        & "Group By Right(ID,2)"
        If SQLNoMatch(SQL) Then
            If PeriodoOrigen <> "BD Previa" Then
                MsgBox "Debe seleccionar un Per�odo del Listado", vbCritical + vbOKOnly, "PER�ODO DE ORIGEN INEXISTENTE"
                .cmbPeriodoDDJJOrigen.SetFocus
                ValidarMigrarDeducciones = False
                Exit Function
            End If
        End If
        
        SQL = "Select Right(ID,2) as Periodo " _
        & "From PresentacionSIRADIG " _
        & "Where Right(ID,2) = '" & Right(PeriodoDestino, 2) & "' " _
        & "Group By Right(ID,2) "
        If SQLNoMatch(SQL) Then
            If PeriodoDestino <> Year(Now()) Then
                MsgBox "Debe seleccionar un Per�odo del Listado", vbCritical + vbOKOnly, "PER�ODO DE DESTINO INEXISTENTE"
                .cmbPeriodoDDJJDestino.SetFocus
                ValidarMigrarDeducciones = False
                Exit Function
            End If
        End If
    
        If PeriodoOrigen = PeriodoDestino Then
            MsgBox "Los per�odos de origen y destino deben ser distintos", vbCritical + vbOKOnly, "PER�ODOS DUPLICADOS"
            .cmbPeriodoDDJJDestino.SetFocus
            ValidarMigrarDeducciones = False
            Exit Function
        End If
    
        If EsIgualTextoEspecificado(TipoDato, "Todas las Deducciones", _
        "Solo Deducciones Personales", "Solo Deducciones Generales") = False Then
            MsgBox "Debe selecciones un Tipo de Datos del Listado", vbCritical + vbOKOnly, "TIPO DE DATOS INEXISTENTE"
            .cmbTipoDatos.SetFocus
            ValidarMigrarDeducciones = False
            Exit Function
        End If
        
    End With
    ValidarMigrarDeducciones = True

End Function

