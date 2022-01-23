Attribute VB_Name = "EditarRegistros"
Public strEditandoAgente As String
Public strEditandoPrecarizado As String
Public strEditandoConcepto As String
Public strEditandoNormaEscalaGanancias As String
Public strEditandoEscalaGanancias As String
Public strEditandoLimitesDeducciones As String
Public strEditandoParentesco As String
Public strEditandoDeduccionesGenerales As String
Public strEditandoCodigoLiquidacion As String
Public strEditandoFamiliar As String
Public strEditandoAutocarga As String
Public bolEditandoRetencionGanancias As Boolean
Public strEditandoPrecarizadoImputado As String
Public strEditandoCodigoSIRADIG As String

Public Sub EditarAgente()

    Dim i As String
    With ListadoAgentes.dgAgentes
        i = .Row
        If i <> 0 Then
            strEditandoAgente = .TextMatrix(i, 2)
            CargaAgente.txtPuestoLaboral.Text = .TextMatrix(i, 2)
            CargaAgente.txtCUIL.Text = .TextMatrix(i, 4)
            CargaAgente.txtDescripcion.Text = .TextMatrix(i, 1)
            CargaAgente.txtLegajo.Text = .TextMatrix(i, 3)
            CargaAgente.txtPuestoLaboral.Text = .TextMatrix(i, 2)
            If .TextMatrix(i, 0) = "SI" Then
                CargaAgente.chkActivado.Value = 1
            Else
                CargaAgente.chkActivado.Value = 0
            End If
            Unload ListadoAgentes
        End If
        i = ""
    End With
    
End Sub

Public Sub EditarPrecarizado()

    Dim i As String
    With ListadoPrecarizados.dgPrecarizados
        i = .Row
        If i <> 0 Then
            strEditandoPrecarizado = .TextMatrix(i, 0)
            CargaPrecarizado.txtNombreCompleto.Text = .TextMatrix(i, 0)
            CargaPrecarizado.mskEstructura.Text = .TextMatrix(i, 1)
            Unload ListadoPrecarizados
        End If
        i = ""
    End With
    
End Sub


Public Sub EditarConcepto()

    Dim i As String
    With ListadoConceptos.dgConceptos
        i = .Row
        If i <> 0 Then
            strEditandoConcepto = .TextMatrix(i, 0)
            CargaConcepto.txtCodigo.Text = .TextMatrix(i, 0)
            CargaConcepto.txtDenominacion.Text = .TextMatrix(i, 1)
            Unload ListadoConceptos
        End If
        i = ""
    End With
    
End Sub

Public Sub EditarNormaEscalaGanancias()

    Dim i As String
    With ListadoEscalaGanancias.dgNormasEscalaGanancias
        i = .Row
        If i > 0 Then
            CargaNormaEscalaGanancias.Show
            strEditandoNormaEscalaGanancias = .TextMatrix(i, 0)
            CargaNormaEscalaGanancias.txtNormaLegal.Text = .TextMatrix(i, 0)
            CargaNormaEscalaGanancias.txtFecha.Text = .TextMatrix(i, 1)
            Unload ListadoEscalaGanancias
        End If
        i = ""
    End With
    
End Sub


Public Sub EditarEscalaGanancias()

    Dim i As String
    Dim strNormaLegal As String
    Dim strFecha As String
    
    With ListadoEscalaGanancias.dgNormasEscalaGanancias
        i = .Row
        strNormaLegal = .TextMatrix(i, 0)
        strFecha = .TextMatrix(i, 1)
    End With
    
    With ListadoEscalaGanancias.dgEscalaGanancias
        i = .Row
        If i > 1 Then
            strEditandoEscalaGanancias = .TextMatrix(i, 0)
            CargaEscalaGanancias.txtNormaLegal.Text = strNormaLegal
            CargaEscalaGanancias.txtNormaLegal.Enabled = False
            CargaEscalaGanancias.txtFecha.Text = strFecha
            CargaEscalaGanancias.txtFecha.Enabled = False
            CargaEscalaGanancias.txtTramo.Text = .TextMatrix(i, 0)
            CargaEscalaGanancias.txtImporteMaximo.Text = .TextMatrix(i, 2)
            CargaEscalaGanancias.txtImporteFijo.Text = .TextMatrix(i, 3)
            CargaEscalaGanancias.txtImporteVariable.Text = Left(.TextMatrix(i, 4), Len(.TextMatrix(i, 4)) - 2)
            Unload ListadoEscalaGanancias
        End If
        i = ""
    End With
    
End Sub

Public Sub EditarLimitesDeducciones()

    Dim i As String
    With ListadoLimitesDeducciones.dgLimitesDeducciones
        i = .Row
        If i > 1 Then
            strEditandoLimitesDeducciones = .TextMatrix(i, 0)
            CargaLimitesDeducciones.txtNormaLegal.Text = .TextMatrix(i, 0)
            CargaLimitesDeducciones.txtFecha.Text = .TextMatrix(i, 1)
            CargaLimitesDeducciones.txtMinimoNoImponible.Text = .TextMatrix(i, 2)
            CargaLimitesDeducciones.txtDeduccionEspecial.Text = .TextMatrix(i, 3)
            CargaLimitesDeducciones.txtHijo.Text = .TextMatrix(i, 4)
            CargaLimitesDeducciones.txtConyuge.Text = .TextMatrix(i, 5)
            CargaLimitesDeducciones.txtOtrasCargas.Text = .TextMatrix(i, 6)
            CargaLimitesDeducciones.txtServicioDomestico.Text = .TextMatrix(i, 7)
            CargaLimitesDeducciones.txtSeguroDeVida.Text = .TextMatrix(i, 8)
            CargaLimitesDeducciones.txtAlquileres.Text = .TextMatrix(i, 9)
            CargaLimitesDeducciones.txtHonorariosMedicos.Text = Left(.TextMatrix(i, 10), Len(.TextMatrix(i, 10)) - 2)
            CargaLimitesDeducciones.txtCuotaMedico.Text = Left(.TextMatrix(i, 11), Len(.TextMatrix(i, 11)) - 2)
            CargaLimitesDeducciones.txtDonaciones.Text = Left(.TextMatrix(i, 12), Len(.TextMatrix(i, 12)) - 2)
            Unload ListadoLimitesDeducciones
        End If
        i = ""
    End With
    
End Sub

Public Sub EditarParentesco()

    Dim i As String
    With ListadoParentesco.dgParentesco
        i = .Row
        If i <> 0 Then
            strEditandoParentesco = .TextMatrix(i, 0)
            CargaParentesco.txtCodigo = .TextMatrix(i, 0)
            CargaParentesco.txtDenominacion.Text = .TextMatrix(i, 1)
            CargaParentesco.txtImporte.Text = .TextMatrix(i, 2)
            Unload ListadoParentesco
        End If
        i = ""
    End With
    
End Sub

Public Sub EditarDeduccionesGenerales()

    Dim i As String
    Dim X As String
    X = ListadoDeduccionesGenerales.dgAgentes.Row
    With ListadoDeduccionesGenerales.dgDeduccionesGenerales
        i = .Row
        If i <> 0 Then
            strEditandoDeduccionesGenerales = .TextMatrix(i, 0)
            CargaDeduccionesGenerales.txtPuestoLaboral.Text = ListadoDeduccionesGenerales.dgAgentes.TextMatrix(X, 2)
            CargaDeduccionesGenerales.txtDescripcion.Text = ListadoDeduccionesGenerales.dgAgentes.TextMatrix(X, 1)
            CargaDeduccionesGenerales.txtFecha.Text = .TextMatrix(i, 0)
            CargaDeduccionesGenerales.txtServicioDomestico.Text = .TextMatrix(i, 1)
            CargaDeduccionesGenerales.txtSeguroDeVida.Text = .TextMatrix(i, 2)
            CargaDeduccionesGenerales.txtAlquileres.Text = .TextMatrix(i, 3)
            CargaDeduccionesGenerales.txtCuotaMedico.Text = .TextMatrix(i, 4)
            CargaDeduccionesGenerales.txtDonaciones.Text = .TextMatrix(i, 5)
            CargaDeduccionesGenerales.txtHonorariosMedicos.Text = .TextMatrix(i, 6)
            Unload ListadoDeduccionesGenerales
        End If
        i = ""
        X = ""
    End With
    
End Sub

Public Sub EditarCodigoLiquidacion()

    Dim i As String
    With ListadoCodigoLiquidaciones.dgCodigoLiquidacion
        i = .Row
        If i <> 0 Then
            strEditandoCodigoLiquidacion = .TextMatrix(i, 0)
            CargaCodigoLiquidacion.txtCodigo = .TextMatrix(i, 0)
            CargaCodigoLiquidacion.txtPeriodo.Text = .TextMatrix(i, 1)
            CargaCodigoLiquidacion.txtDescripcion.Text = .TextMatrix(i, 2)
            CargaCodigoLiquidacion.txtMontoExento.Text = .TextMatrix(i, 3)
            Unload ListadoCodigoLiquidaciones
        End If
        i = ""
    End With
    
End Sub

Public Sub EditarFamiliar()

    Dim i As String
    Dim X As String
    
    CargaFamiliar.Show
    CargarcmbFamiliar
    X = ListadoFamiliares.dgAgentes.Row
    With ListadoFamiliares.dgFamiliares
        i = .Row
        If i <> 0 Then
            strEditandoFamiliar = .TextMatrix(i, 3)
            CargaFamiliar.txtPuestoLaboral.Text = ListadoFamiliares.dgAgentes.TextMatrix(X, 2)
            CargaFamiliar.txtDescripcionAgente.Text = ListadoFamiliares.dgAgentes.TextMatrix(X, 1)
            CargaFamiliar.txtDNI.Text = .TextMatrix(i, 3)
            CargaFamiliar.txtDescripcionFamiliar.Text = .TextMatrix(i, 1)
            CargaFamiliar.cmbParentesco.Text = .TextMatrix(i, 2)
            CargaFamiliar.txtFechaAlta.Text = .TextMatrix(i, 0)
            CargaFamiliar.cmbNivelDeEstudio.Text = .TextMatrix(i, 5)
            If .TextMatrix(i, 7) = "SI" Then
                CargaFamiliar.chkCobraSalario.Value = 1
            Else
                CargaFamiliar.chkCobraSalario.Value = 0
            End If
            If .TextMatrix(i, 6) = "SI" Then
                CargaFamiliar.chkAdherente.Value = 1
            Else
                CargaFamiliar.chkAdherente.Value = 0
            End If
            If .TextMatrix(i, 8) = "SI" Then
                CargaFamiliar.chkDiscapacitado.Value = 1
            Else
                CargaFamiliar.chkDiscapacitado.Value = 0
            End If
            If .TextMatrix(i, 4) = "SI" Then
                CargaFamiliar.chkGanancias.Value = 1
            Else
                CargaFamiliar.chkGanancias.Value = 0
            End If
            Unload ListadoFamiliares
        End If
        i = ""
        X = ""
    End With
    
End Sub

'Public Sub EditarRetencionGananciasViejo()
'
'    Dim i As String
'    Dim x As String
'    Dim SQL As String
'    Dim strCL As String
'    Dim strPL As String
'
'    x = ListadoLiquidacionGanancias.dgCodigosLiquidacionesGanancias.Row
'    i = ListadoLiquidacionGanancias.dgAgentesRetenidos.Row
'
'    With LiquidacionGanancia4ta
'        If i <> 0 Then
'            If IsNumeric(ListadoLiquidacionGanancias.dgAgentesRetenidos.TextMatrix(i, 2)) = True Then
'                .Show
'                bolEditandoRetencionGanancias = True
'                strCL = ListadoLiquidacionGanancias.dgCodigosLiquidacionesGanancias.TextMatrix(x, 0)
'                .txtCodigoLiquidacion.Text = strCL
'                .txtDescripcionPeriodo.Text = ListadoLiquidacionGanancias.dgCodigosLiquidacionesGanancias.TextMatrix(x, 1)
'                SQL = "Select * From CODIGOLIQUIDACIONES Where CODIGO = '" & strCL & "'"
'                Set rstBuscarSlave = New ADODB.Recordset
'                rstBuscarSlave.Open SQL, dbSlave, adOpenDynamic, adLockReadOnly
'                .txtPeriodo = rstBuscarSlave!PERIODO
'                rstBuscarSlave.Close
'                Set rstBuscarSlave = Nothing
'                .txtPuestoLaboral.Text = ListadoLiquidacionGanancias.dgAgentesRetenidos.TextMatrix(i, 0)
'                .txtPuestoLaboral.Enabled = False
'                .txtDescripcionAgente.Text = ListadoLiquidacionGanancias.dgAgentesRetenidos.TextMatrix(i, 1)
'                Call LiquidacionGanancias("PasoUno", .txtPuestoLaboral.Text, .txtCodigoLiquidacion.Text)
'                SQL = "Select * From LIQUIDACIONGANANCIAS4TACATEGORIA Where CODIGOLIQUIDACION = '" & .txtCodigoLiquidacion.Text & "' " _
'                & "And PUESTOLABORAL = '" & .txtPuestoLaboral.Text & "'"
'                Set rstBuscarSlave = New ADODB.Recordset
'                rstBuscarSlave.Open SQL, dbSlave, adOpenDynamic, adLockReadOnly
'                .txtHaberOptimo.Text = De_Num_a_Tx_01(rstBuscarSlave!HaberOptimo)
'                .txtPluriempleo.Text = De_Num_a_Tx_01(rstBuscarSlave!Pluriempleo)
'                .txtAjuste.Text = De_Num_a_Tx_01(rstBuscarSlave!Ajuste)
'                .txtJubilacion.Text = De_Num_a_Tx_01(rstBuscarSlave!Jubilacion)
'                .txtObraSocial.Text = De_Num_a_Tx_01(rstBuscarSlave!ObraSocial)
'                .txtAdherente.Text = De_Num_a_Tx_01(rstBuscarSlave!AdherenteObraSocial)
'                .txtSeguroObligatorio.Text = De_Num_a_Tx_01(rstBuscarSlave!SeguroDeVidaObligatorio)
'                If De_Txt_a_Num_01(.txtSeguroOptativo.Text) <> 0 Then
'                    If De_Txt_a_Num_01(.txtSeguroOptativo.Text) < rstBuscarSlave!SeguroDeVidaOptativo Then
'                        .dgDeduccionesGenerales.TextMatrix(2, 1) = De_Num_a_Tx_01(rstBuscarSlave!SeguroDeVidaOptativo) - De_Txt_a_Num_01(.txtSeguroOptativo.Text)
'                    End If
'                Else
'                    .dgDeduccionesGenerales.TextMatrix(2, 1) = De_Num_a_Tx_01(rstBuscarSlave!SeguroDeVidaOptativo)
'                End If
'                .txtCuotaSindical.Text = De_Num_a_Tx_01(rstBuscarSlave!CuotaSindical)
'                .dgDeduccionesGenerales.TextMatrix(1, 1) = De_Num_a_Tx_01(rstBuscarSlave!ServicioDomestico)
'                .dgDeduccionesGenerales.TextMatrix(3, 1) = De_Num_a_Tx_01(rstBuscarSlave!CuotaMedicoAsistencial)
'                .dgDeduccionesGenerales.TextMatrix(4, 1) = De_Num_a_Tx_01(rstBuscarSlave!Donaciones)
'                'rstBuscarSlave!HonorariosMedicos = 0 'Determinar qué hacer en Liquidación Final / Anual
'                .dgDeduccionesPersonales.TextMatrix(1, 1) = De_Num_a_Tx_01(rstBuscarSlave!MinimoNoImponible)
'                .dgDeduccionesPersonales.TextMatrix(6, 1) = De_Num_a_Tx_01(rstBuscarSlave!DeduccionEspecial)
'                .dgDeduccionesPersonales.TextMatrix(3, 1) = De_Num_a_Tx_01(rstBuscarSlave!Conyuge)
'                .dgDeduccionesPersonales.TextMatrix(4, 1) = De_Num_a_Tx_01(rstBuscarSlave!Hijo)
'                .dgDeduccionesPersonales.TextMatrix(5, 1) = De_Num_a_Tx_01(rstBuscarSlave!OtrasCargasDeFamilia)
'                .txtAjuesteRetencion.Text = De_Num_a_Tx_01(rstBuscarSlave!AjusteRetencion)
'                .txtRetencionPeriodo.Text = De_Num_a_Tx_01(rstBuscarSlave!Retencion)
'                'Calculamos el total de Deducciones Generales
'                .dgDeduccionesGenerales.TextMatrix(6, 1) = De_Num_a_Tx_01(De_Txt_a_Num_01(.dgDeduccionesGenerales.TextMatrix(6, 1)) _
'                - De_Txt_a_Num_01(.dgDeduccionesGenerales.TextMatrix(5, 1)))
'                .dgDeduccionesGenerales.TextMatrix(5, 1) = De_Num_a_Tx_01((De_Txt_a_Num_01(.dgDeduccionesGenerales.TextMatrix(1, 1)) _
'                + De_Txt_a_Num_01(.dgDeduccionesGenerales.TextMatrix(2, 1)) + De_Txt_a_Num_01(.dgDeduccionesGenerales.TextMatrix(3, 1)) _
'                + De_Txt_a_Num_01(.dgDeduccionesGenerales.TextMatrix(4, 1))) * (-1))
'                .dgDeduccionesGenerales.TextMatrix(6, 1) = De_Num_a_Tx_01(De_Txt_a_Num_01(.dgDeduccionesGenerales.TextMatrix(6, 1)) _
'                + De_Txt_a_Num_01(.dgDeduccionesGenerales.TextMatrix(5, 1)))
'                'Calculamos el total de Deducciones Personales
'                .dgDeduccionesPersonales.TextMatrix(8, 1) = De_Num_a_Tx_01(De_Txt_a_Num_01(.dgDeduccionesPersonales.TextMatrix(8, 1)) _
'                - De_Txt_a_Num_01(.dgDeduccionesPersonales.TextMatrix(7, 1)))
'                .dgDeduccionesPersonales.TextMatrix(7, 1) = De_Num_a_Tx_01((De_Txt_a_Num_01(.dgDeduccionesPersonales.TextMatrix(1, 1)) _
'                + De_Txt_a_Num_01(.dgDeduccionesPersonales.TextMatrix(3, 1)) + De_Txt_a_Num_01(.dgDeduccionesPersonales.TextMatrix(4, 1)) _
'                + De_Txt_a_Num_01(.dgDeduccionesPersonales.TextMatrix(5, 1)) + De_Txt_a_Num_01(.dgDeduccionesPersonales.TextMatrix(6, 1))) * (-1))
'                .dgDeduccionesPersonales.TextMatrix(8, 1) = De_Num_a_Tx_01(De_Txt_a_Num_01(.dgDeduccionesPersonales.TextMatrix(8, 1)) _
'                + De_Txt_a_Num_01(.dgDeduccionesPersonales.TextMatrix(7, 1)))
'                'Volvemos a Calcular los Totales
'                rstBuscarSlave.Close
'                Set rstBuscarSlave = Nothing
'                Call LiquidacionGanancias("PasoDos")
'                'Call LiquidacionGanancias("PasoTres")
'                Call LiquidacionGanancias("PasoCuatro")
'                Call LiquidacionGanancias("PasoSiete")
'            Else
'                .Show
'                .txtCodigoLiquidacion.Text = ListadoLiquidacionGanancias.dgCodigosLiquidacionesGanancias.TextMatrix(x, 0)
'                .txtDescripcionPeriodo.Text = ListadoLiquidacionGanancias.dgCodigosLiquidacionesGanancias.TextMatrix(x, 1)
'                SQL = "Select * From CODIGOLIQUIDACIONES Where CODIGO = '" & ListadoLiquidacionGanancias.dgCodigosLiquidacionesGanancias.TextMatrix(x, 0) & "'"
'                Set rstBuscarSlave = New ADODB.Recordset
'                rstBuscarSlave.Open SQL, dbSlave, adOpenDynamic, adLockReadOnly
'                .txtPeriodo = rstBuscarSlave!PERIODO
'                rstBuscarSlave.Close
'                Set rstBuscarSlave = Nothing
'                .txtPuestoLaboral.Enabled = True
'                .txtPuestoLaboral.SetFocus
'                .txtPuestoLaboral.Text = ListadoLiquidacionGanancias.dgAgentesRetenidos.TextMatrix(i, 0)
'                .txtHaberOptimo.SetFocus
'                .txtPuestoLaboral.Enabled = False
'            End If
'        End If
'        Unload ListadoLiquidacionGanancias
'        i = ""
'        x = ""
'        SQL = ""
'    End With
'
'End Sub

Public Sub EditarRetencionGanancias()

    Dim i As String
    Dim X As String
    Dim SQL As String
    Dim strCL As String
    Dim strPL As String

    
    X = ListadoLiquidacionGanancias.dgCodigosLiquidacionesGanancias.Row
    i = ListadoLiquidacionGanancias.dgAgentesRetenidos.Row
    
    With LiquidacionGanancia4ta
        If i <> 0 Then
            strCL = ListadoLiquidacionGanancias.dgCodigosLiquidacionesGanancias.TextMatrix(X, 0)
            strPL = ListadoLiquidacionGanancias.dgAgentesRetenidos.TextMatrix(i, 0)
            If IsNumeric(ListadoLiquidacionGanancias.dgAgentesRetenidos.TextMatrix(i, 2)) = True Then
                'Abrimos formulario y completamos datos básicos
                .Show
                bolEditandoRetencionGanancias = True
                .txtCodigoLiquidacion.Text = strCL & " - (" & BuscarPeriodoLiquidacion(strCL) & ")"
                .txtDescripcionPeriodo.Text = ListadoLiquidacionGanancias.dgCodigosLiquidacionesGanancias.TextMatrix(X, 1)
                '.txtPeriodo.Text = BuscarPeriodoLiquidacion(strCL)
                .txtPuestoLaboral.Text = strPL
                .txtPuestoLaboral.Enabled = False
                .txtDescripcionAgente.Text = ListadoLiquidacionGanancias.dgAgentesRetenidos.TextMatrix(i, 1)
                Call EditandoLiquidacionGanancia4ta(strPL, strCL)
            Else
                .Show
                .txtCodigoLiquidacion.Text = strCL & " - (" & BuscarPeriodoLiquidacion(strCL) & ")"
                .txtDescripcionPeriodo.Text = ListadoLiquidacionGanancias.dgCodigosLiquidacionesGanancias.TextMatrix(X, 1)
                '.txtPeriodo = BuscarPeriodoLiquidacion(strCL)
                .txtPuestoLaboral.Enabled = True
                .txtPuestoLaboral.SetFocus
                .txtPuestoLaboral.Text = strPL
                .txtHaberOptimo.SetFocus
                .txtPuestoLaboral.Enabled = False
            End If
        End If
        Unload ListadoLiquidacionGanancias
        i = ""
        X = ""
        SQL = ""
        strCL = ""

    End With
    
End Sub

Public Sub EditarAutocarga()

    Dim i As String
    Dim intRespuesta As Integer
    Dim strFecha As String
    Dim SQL As String
    
    With Autocarga.dgListadoAutocarga
        i = .Row
        If i <> 0 Then
            SQL = "Select PROVEEDOR From LIQUIDACIONHONORARIOS" _
            & " Where COMPROBANTE = '" & .TextMatrix(i, 0) & "'" _
            & " And PROVEEDOR Not In (Select AGENTES As PROVEEDOR From PRECARIZADOS)"
            If SQLNoMatch(SQL) = False Then
                'Insertamos los agentes que no tienen estructura en la tabla Precarizados
                SQL = "Insert Into PRECARIZADOS (Agentes, Actividad, Partida)" _
                & " Select Proveedor, '" & "00-00-00" & "', '" & "000" & "'" _
                & " From (" & SQL & ")"
                dbSlave.BeginTrans
                dbSlave.Execute SQL
                dbSlave.CommitTrans
                intRespuesta = MsgBox("Existen Agentes SIN Estructura Presupuestaria." & vbCrLf & "Desea clasificarlos en este momento? De no hacerlo, no podrá seguir adelante con la operación", vbQuestion + vbYesNo, "AGENTES SIN ESTRUCTURA")
                If intRespuesta = 6 Then
                    Unload Autocarga
                    ListadoPrecarizados.Show
                    Call ConfigurardgPrecarizados
                    Call CargardgPrecarizados
                End If
            Else
                SQL = "Select * From LIQUIDACIONHONORARIOS Inner Join PRECARIZADOS" _
                & " On LIQUIDACIONHONORARIOS.PROVEEDOR = PRECARIZADOS.AGENTES" _
                & " Where COMPROBANTE = '" & .TextMatrix(i, 0) & "'" _
                & " And (Left(PRECARIZADOS.ACTIVIDAD,2) = '00'" _
                & " Or Right(PRECARIZADOS.ACTIVIDAD,2) = '00'" _
                & " Or PRECARIZADOS.Partida = '000')"
                If SQLNoMatch(SQL) = False Then
                    intRespuesta = MsgBox("Existen Agentes SIN Estructura Presupuestaria." & vbCrLf & "Desea clasificarlos en este momento? De no hacerlo, no podrá seguir adelante con la operación", vbQuestion + vbYesNo, "AGENTES SIN ESTRUCTURA")
                    If intRespuesta = 6 Then
                        Unload Autocarga
                        ListadoPrecarizados.Show
                        Call ConfigurardgPrecarizados
                        Call CargardgPrecarizados
                    End If
                Else
                    strEditandoAutocarga = .TextMatrix(i, 0)
                    CargaComprobanteSIIF.Show
                    strFecha = Format(Day(Now()), "00") & "/" & Format(Month(Now()), "00")
                    strFecha = strFecha & "/" & Year(Now())
                    CargaComprobanteSIIF.txtFecha.Text = strFecha
                    CargaComprobanteSIIF.txtImporte.Text = .TextMatrix(i, 2)
                    CargaComprobanteSIIF.txtRetenciones.Text = .TextMatrix(i, 3)
                    CargaComprobanteSIIF.txtFuente.Text = "11"
                    CargaComprobanteSIIF.txtDescripcionFuente.Text = "RECURSOS PROPIOS DE LAS INSTITUCIONES"
                    CargaComprobanteSIIF.txtCuenta.Text = "130832005"
                    CargaComprobanteSIIF.txtDescripcionCuenta.Text = "FUNCIONAMIENTO"
                    Unload Autocarga
                End If
            End If
        End If
        i = ""
        intRespuesta = 0
    End With
    
End Sub

Public Sub EditarPrecarizadoImputado()

    Dim i As String
    With ListadoHonorariosImputados.dgListadoHonorariosImputados
        i = .Row
        If i <> 0 Then
            strEditandoPrecarizadoImputado = ListadoHonorariosImputados.txtComprobante.Text & "-" _
            & .TextMatrix(i, 1) & "-" & De_Num_a_Tx_01(.TextMatrix(i, 2))
            CargaPrecarizadoImputado.txtComprobante.Text = ListadoHonorariosImputados.txtComprobante.Text
            CargaPrecarizadoImputado.txtComprobante.Enabled = False
            CargaPrecarizadoImputado.txtPeriodo.Text = ListadoHonorariosImputados.txtPeriodo.Text
            CargaPrecarizadoImputado.txtPeriodo.Enabled = False
            CargaPrecarizadoImputado.cmbNombreCompleto.Text = .TextMatrix(i, 1)
            CargaPrecarizadoImputado.cmbNombreCompleto.Enabled = False
            SQL = "Select ACTIVIDAD, PARTIDA From PRECARIZADOS " _
            & "Where AGENTES = '" & .TextMatrix(i, 1) & "'"
            Set rstBuscarSlave = New ADODB.Recordset
            rstBuscarSlave.Open SQL, dbSlave, adOpenDynamic, adLockReadOnly
            CargaPrecarizadoImputado.mskEstructuraPrevista.Text = rstBuscarSlave!ACTIVIDAD _
            & "-" & rstBuscarSlave!PARTIDA
            CargaPrecarizadoImputado.mskEstructuraPrevista.Enabled = False
            rstBuscarSlave.Close
            Set rstBuscarSlave = Nothing
            CargaPrecarizadoImputado.mskEstructuraImputada.Text = .TextMatrix(i, 0)
            CargaPrecarizadoImputado.txtMontoBruto.Text = De_Num_a_Tx_01(.TextMatrix(i, 2))
            Unload ListadoHonorariosImputados
        End If
        i = ""
    End With
    
End Sub

Public Sub EditarCodigoSIRADIG()

    Dim i As String
    With ListadoCodigosSIRADIG.dgCodigosSIRADIG
        i = .Row
        If i <> 0 Then
            strCargaCodigoSIRADIG = strListadoCodigoSIRADIG
            strEditandoCodigoSIRADIG = .TextMatrix(i, 0)
            CargaCodigoSIRADIG.txtCodigo.Text = .TextMatrix(i, 0)
            CargaCodigoSIRADIG.txtDenominacion.Text = .TextMatrix(i, 1)
            strListadoCodigoSIRADIG = ""
            Unload ListadoCodigosSIRADIG
        End If
        i = ""
    End With
    
End Sub

Public Sub EditarRetencionGananciasSIRADIG()

    Dim i As String
    Dim X As String
    Dim SQL As String
    Dim strCL As String
    Dim strPL As String

    
    X = ListadoLiquidacionGanancias.dgCodigosLiquidacionesGanancias.Row
    i = ListadoLiquidacionGanancias.dgAgentesRetenidos.Row
    
    With LiquidacionGanancia4taSIRADIG
        If i <> 0 Then
            strCL = ListadoLiquidacionGanancias.dgCodigosLiquidacionesGanancias.TextMatrix(X, 0)
            strPL = ListadoLiquidacionGanancias.dgAgentesRetenidos.TextMatrix(i, 0)
            If IsNumeric(ListadoLiquidacionGanancias.dgAgentesRetenidos.TextMatrix(i, 2)) = True Then
                'Abrimos formulario y completamos datos básicos
                .Show
                bolEditandoRetencionGanancias = True
                .txtCodigoLiquidacion.Text = strCL & " - (" & BuscarPeriodoLiquidacion(strCL) & ")"
                .txtDescripcionPeriodo.Text = ListadoLiquidacionGanancias.dgCodigosLiquidacionesGanancias.TextMatrix(X, 1)
                Call CargarCmbLiquidacionGanancia4taSIRADIG(strPL, strCL)
                '.txtPeriodo.Text = BuscarPeriodoLiquidacion(strCL)
                .txtPuestoLaboral.Text = strPL
                .txtPuestoLaboral.Enabled = False
                .txtDescripcionAgente.Text = ListadoLiquidacionGanancias.dgAgentesRetenidos.TextMatrix(i, 1)
                'MODIFICAR PROCEDIMIENTO
                Call EditandoLiquidacionGanancia4taSIRADIG(strPL, strCL)
            Else
                .Show
                .txtCodigoLiquidacion.Text = strCL & " - (" & BuscarPeriodoLiquidacion(strCL) & ")"
                .txtDescripcionPeriodo.Text = ListadoLiquidacionGanancias.dgCodigosLiquidacionesGanancias.TextMatrix(X, 1)
                'No es necesario lo siguiente? -- Call CargarCmbLiquidacionGanancia4taSIRADIG(strPL, strCL)
                '.txtPeriodo = BuscarPeriodoLiquidacion(strCL)
                .txtPuestoLaboral.Enabled = True
                .txtPuestoLaboral.SetFocus
                .txtPuestoLaboral.Text = strPL
                .txtHaberOptimo.SetFocus
                .txtPuestoLaboral.Enabled = False
            End If
        End If
        Unload ListadoLiquidacionGanancias
        i = ""
        X = ""
        SQL = ""
        strCL = ""

    End With
    
End Sub
