Attribute VB_Name = "ProcedimientoLiquidacionViejo"
'Public Sub LiquidacionGananciasViejo(PasoLiquidacion As String, Optional PuestoLaboral As String, Optional CodigoLiquidacion As String)
'
'    Dim SQL As String
'    Dim datFecha As Date
'    Dim dblAcumulado As Double
'    Dim dblHaberSISPERAcumuladoEstimado As Double
'
'    Select Case PasoLiquidacion
'    Case "PasoUno"
'        With LiquidacionGanancia4ta
'            Set rstRegistroSlave = New ADODB.Recordset
'            SQL = "Select CODIGOLIQUIDACION From LIQUIDACIONSUELDOS Where CODIGOLIQUIDACION <= '" & CodigoLiquidacion & "' " _
'            & "And CODIGOCONCEPTO = '0115' Order by CODIGOLIQUIDACION Desc"
'            rstRegistroSlave.Open SQL, dbSlave, adOpenForwardOnly, adLockOptimistic
'            rstRegistroSlave.MoveFirst
'            CodigoLiquidacion = rstRegistroSlave!CodigoLiquidacion
'            rstRegistroSlave.Close
'            'Buscamos el Haber Óptimo
'            SQL = "Select * from LIQUIDACIONSUELDOS Where PUESTOLABORAL = '" & PuestoLaboral & "' And CODIGOLIQUIDACION = '" & CodigoLiquidacion & "' And CODIGOCONCEPTO = '0115'"
'            rstRegistroSlave.Open SQL, dbSlave, adOpenForwardOnly, adLockOptimistic
'            dblAcumulado = rstRegistroSlave!Importe
'            rstRegistroSlave.Close
'            'Ajustamos por Ley Nro 3.801
'            '1) Compensación Funcional
'            'SQL = "Select * from LIQUIDACIONSUELDOS Where PUESTOLABORAL = '" & PuestoLaboral & "' And CODIGOLIQUIDACION = '" & CodigoLiquidacion & "' And CODIGOCONCEPTO = '0002'"
'            'If SQLNoMatch(SQL) = False Then
'            '    rstRegistroSlave.Open SQL, dbSlave, adOpenForwardOnly, adLockOptimistic
'            '    dblAcumulado = dblAcumulado - rstRegistroSlave!Importe
'            '    rstRegistroSlave.Close
'            'End If
'            '2) Gastos de Representación
'            'SQL = "Select * from LIQUIDACIONSUELDOS Where PUESTOLABORAL = '" & PuestoLaboral & "' And CODIGOLIQUIDACION = '" & CodigoLiquidacion & "' And CODIGOCONCEPTO = '0028'"
'            'If SQLNoMatch(SQL) = False Then
'            '    rstRegistroSlave.Open SQL, dbSlave, adOpenForwardOnly, adLockOptimistic
'            '    dblAcumulado = dblAcumulado - rstRegistroSlave!Importe
'            '    rstRegistroSlave.Close
'            'End If
'            '3) Dedicación Exclusiva
'            'SQL = "Select * from LIQUIDACIONSUELDOS Where PUESTOLABORAL = '" & PuestoLaboral & "' And CODIGOLIQUIDACION = '" & CodigoLiquidacion & "' And CODIGOCONCEPTO = '0064'"
'            'If SQLNoMatch(SQL) = False Then
'            '    rstRegistroSlave.Open SQL, dbSlave, adOpenForwardOnly, adLockOptimistic
'            '    dblAcumulado = dblAcumulado - rstRegistroSlave!Importe
'            '    rstRegistroSlave.Close
'            'End If
'            .txtHaberOptimo.Text = De_Num_a_Tx_01(dblAcumulado, , 2)
'            If rstRegistroSlave.State = adStateOpen Then
'                rstRegistroSlave.Close
'            End If
'            dblAcumulado = 0
'            'Cargamos Sueldo Pluriempleo
'            SQL = "Select * from LIQUIDACIONGANANCIAS4TACATEGORIA Where PUESTOLABORAL = '" & PuestoLaboral & "' And CODIGOLIQUIDACION < '" & .txtCodigoLiquidacion.Text & "' " _
'            & " And PLURIEMPLEO <> 0 Order By CODIGOLIQUIDACION Desc"
'            If SQLNoMatch(SQL) = True Then
'               .txtPluriempleo.Text = De_Num_a_Tx_01(0)
'            Else
'                rstRegistroSlave.Open SQL, dbSlave, adOpenForwardOnly, adLockOptimistic
'                .txtPluriempleo.Text = De_Num_a_Tx_01(rstRegistroSlave!Pluriempleo, , 2)
'                rstRegistroSlave.Close
'                .txtPluriempleo.BackColor = &HFF&
'            End If
'            'Cargamos el Ajuste
'            .txtAjuste.Text = De_Num_a_Tx_01(0)
'            'Buscamos el Descuento por Jubilacion
'            SQL = "Select * from LIQUIDACIONSUELDOS Where PUESTOLABORAL = '" & PuestoLaboral & "' And CODIGOLIQUIDACION = '" & CodigoLiquidacion & "' And CODIGOCONCEPTO = '0208'"
'            rstRegistroSlave.Open SQL, dbSlave, adOpenForwardOnly, adLockOptimistic
'            .txtJubilacion.Text = De_Num_a_Tx_01(rstRegistroSlave!Importe, , 2)
'            rstRegistroSlave.Close
'            'Buscamos el Descuento por ObraSocial
'            SQL = "Select * from LIQUIDACIONSUELDOS Where PUESTOLABORAL = '" & PuestoLaboral & "' And CODIGOLIQUIDACION = '" & CodigoLiquidacion & "' And CODIGOCONCEPTO = '0212'"
'            rstRegistroSlave.Open SQL, dbSlave, adOpenForwardOnly, adLockOptimistic
'            .txtObraSocial.Text = De_Num_a_Tx_01(rstRegistroSlave!Importe, , 2)
'            rstRegistroSlave.Close
'            'Cargamos el Adherente Obra Social
'            SQL = "Select * from LIQUIDACIONSUELDOS Where PUESTOLABORAL = '" & PuestoLaboral & "' And CODIGOLIQUIDACION = '" & CodigoLiquidacion & "' And CODIGOCONCEPTO = '0234'"
'            If SQLNoMatch(SQL) = True Then
'                .txtAdherente.Text = De_Num_a_Tx_01(0)
'            Else
'                rstRegistroSlave.Open SQL, dbSlave, adOpenForwardOnly, adLockOptimistic
'                .txtAdherente.Text = De_Num_a_Tx_01(rstRegistroSlave!Importe, , 2)
'                rstRegistroSlave.Close
'                SQL = "Select * from LIQUIDACIONGANANCIAS4TACATEGORIA Where PUESTOLABORAL = '" & PuestoLaboral & "' And CODIGOLIQUIDACION = '" & CodigoLiquidacion & "' " _
'                & "Order by CODIGOLIQUIDACION Desc"
'                rstRegistroSlave.Open SQL, dbSlave, adOpenForwardOnly, adLockOptimistic
'                If rstRegistroSlave.BOF = False Then
'                    If rstRegistroSlave!AdherenteObraSocial <> De_Txt_a_Num_01(.txtAdherente.Text) Then
'                        .txtAdherente.Text = De_Num_a_Tx_01(rstRegistroSlave!AdherenteObraSocial)
'                        .txtAdherente.BackColor = &HFF&
'                    End If
'                Else
'                    .txtAdherente.BackColor = &HFF&
'                End If
'                rstRegistroSlave.Close
'            End If
'            'Buscamos el Descuento Seguro Obligatorio
'            SQL = "Select * from LIQUIDACIONSUELDOS Where ((PUESTOLABORAL = '" & PuestoLaboral & "' And CODIGOLIQUIDACION = '" & CodigoLiquidacion & "' And CODIGOCONCEPTO = '0360') Or " _
'            & "(PUESTOLABORAL = '" & PuestoLaboral & "' And CODIGOLIQUIDACION = '" & CodigoLiquidacion & "' And CODIGOCONCEPTO = '0366') Or " _
'            & "(PUESTOLABORAL = '" & PuestoLaboral & "' And CODIGOLIQUIDACION = '" & CodigoLiquidacion & "' And CODIGOCONCEPTO = '0369'))"
'            If SQLNoMatch(SQL) = True Then
'                .txtSeguroObligatorio.Text = De_Num_a_Tx_01(0)
'            Else
'                rstRegistroSlave.Open SQL, dbSlave, adOpenForwardOnly, adLockOptimistic
'                .txtSeguroObligatorio.Text = De_Num_a_Tx_01(rstRegistroSlave!Importe, , 2)
'                rstRegistroSlave.Close
'            End If
'            'Buscamos el Descuento Cuota Sindical
'            SQL = "Select * from LIQUIDACIONSUELDOS Where ((PUESTOLABORAL = '" & PuestoLaboral & "' And CODIGOLIQUIDACION = '" & CodigoLiquidacion & "' And CODIGOCONCEPTO = '0219') Or " _
'            & "(PUESTOLABORAL = '" & PuestoLaboral & "' And CODIGOLIQUIDACION = '" & CodigoLiquidacion & "' And CODIGOCONCEPTO = '0227'))"
'            If SQLNoMatch(SQL) = True Then
'                .txtCuotaSindical.Text = De_Num_a_Tx_01(0)
'            Else
'                rstRegistroSlave.Open SQL, dbSlave, adOpenForwardOnly, adLockOptimistic
'                .txtCuotaSindical.Text = De_Num_a_Tx_01(rstRegistroSlave!Importe, , 2)
'                rstRegistroSlave.Close
'            End If
'            'Buscamos el Descuento Seguro Optativo
'            SQL = "Select * from LIQUIDACIONSUELDOS Where ((PUESTOLABORAL = '" & PuestoLaboral & "' And CODIGOLIQUIDACION = '" & CodigoLiquidacion & "' And CODIGOCONCEPTO = '0317') Or " _
'            & "(PUESTOLABORAL = '" & PuestoLaboral & "' And CODIGOLIQUIDACION = '" & CodigoLiquidacion & "' And CODIGOCONCEPTO = '0361') Or " _
'            & "(PUESTOLABORAL = '" & PuestoLaboral & "' And CODIGOLIQUIDACION = '" & CodigoLiquidacion & "' And CODIGOCONCEPTO = '0367') Or " _
'            & "(PUESTOLABORAL = '" & PuestoLaboral & "' And CODIGOLIQUIDACION = '" & CodigoLiquidacion & "' And CODIGOCONCEPTO = '0370') Or " _
'            & "(PUESTOLABORAL = '" & PuestoLaboral & "' And CODIGOLIQUIDACION = '" & CodigoLiquidacion & "' And CODIGOCONCEPTO = '0373') Or " _
'            & "(PUESTOLABORAL = '" & PuestoLaboral & "' And CODIGOLIQUIDACION = '" & CodigoLiquidacion & "' And CODIGOCONCEPTO = '0374'))"
'            If SQLNoMatch(SQL) = True Then
'               .txtSeguroOptativo.Text = De_Num_a_Tx_01(0)
'            Else
'                'Buscamos el límite de Seguro de Vida
'                datFecha = DateTime.DateSerial(Right(.txtPeriodo.Text, 4), Left(.txtPeriodo.Text, 2), 1)
'                rstRegistroSlave.Open SQL, dbSlave, adOpenForwardOnly, adLockOptimistic
'                Set rstBuscarSlave = New ADODB.Recordset
'                rstBuscarSlave.Open "Select SEGURODEVIDA From DEDUCCIONES4TACATEGORIA Where FECHA <= #" & Format(datFecha, "MM/DD/YYYY") & "# Order by FECHA Desc", dbSlave, adOpenForwardOnly, adLockOptimistic
'                If rstRegistroSlave!Importe < (rstBuscarSlave!SeguroDeVida / 12) Then
'                    .txtSeguroOptativo.Text = De_Num_a_Tx_01(rstRegistroSlave!Importe, , 2)
'                Else
'                    .txtSeguroOptativo.Text = De_Num_a_Tx_01((rstBuscarSlave!SeguroDeVida / 12), , 2)
'                End If
'                rstRegistroSlave.Close
'                rstBuscarSlave.Close
'                Set rstBuscarSlave = Nothing
'            End If
'            'Cargamos Otros Descuentos (VARIABLE SIN USO)
'            .txtOtrosDescuentos.Text = De_Num_a_Tx_01(0)
'            'Cargamos Ajuste Retenciones
'            SQL = "Select * from LIQUIDACIONGANANCIAS4TACATEGORIA Where PUESTOLABORAL = '" & PuestoLaboral & "' And CODIGOLIQUIDACION < '" & .txtCodigoLiquidacion.Text & "' " _
'            & "Order By CODIGOLIQUIDACION Desc"
'            rstRegistroSlave.Open SQL, dbSlave, adOpenForwardOnly, adLockOptimistic
'            If SQLNoMatch(SQL) = True Then
'               .txtAjuesteRetencion.Text = De_Num_a_Tx_01(0)
'            ElseIf rstRegistroSlave!AjusteRetencion = 0 Then
'               .txtAjuesteRetencion.Text = De_Num_a_Tx_01(0)
'            Else
'                .txtAjuesteRetencion.Text = De_Num_a_Tx_01(rstRegistroSlave!AjusteRetencion, , 2)
'                .txtAjuesteRetencion.BackColor = &HFF&
'            End If
'            rstRegistroSlave.Close
'            Set rstRegistroSlave = Nothing
'            Call ControlSueldoLiquidado("Inicial")
'            Call LiquidacionGanancias("PasoDos")
'            Call LiquidacionGanancias("PasoTres")
'            Call LiquidacionGanancias("PasoCuatro")
'            Call LiquidacionGanancias("PasoCinco")
'            Call LiquidacionGanancias("PasoSeis")
'            Call LiquidacionGanancias("PasoSiete")
'        End With
'    Case "PasoDos"
'        With LiquidacionGanancia4ta
'            'Calculamos Subtotal Sueldo
'            .txtSubtotalSueldo.Text = De_Num_a_Tx_01(De_Txt_a_Num_01(.txtHaberOptimo.Text) _
'            + De_Txt_a_Num_01(.txtPluriempleo.Text) + De_Txt_a_Num_01(.txtAjuste.Text), , 2)
'            'Calculamos Subtotal Descuentos
'            dblAcumulado = De_Txt_a_Num_01(.txtJubilacion.Text) + De_Txt_a_Num_01(.txtObraSocial.Text) _
'            + De_Txt_a_Num_01(.txtAdherente.Text) + De_Txt_a_Num_01(.txtSeguroObligatorio.Text) _
'            + De_Txt_a_Num_01(.txtCuotaSindical.Text) + De_Txt_a_Num_01(.txtSeguroOptativo.Text) _
'            + De_Txt_a_Num_01(.txtOtrosDescuentos.Text)
'            dblAcumulado = dblAcumulado * (-1)
'            .txtSubtotalDescuento.Text = De_Num_a_Tx_01(dblAcumulado, , 2)
'            dblAcumulado = 0
'        End With
'    Case "PasoTres"
'        Dim dblAportesPersonales As Double
'        With LiquidacionGanancia4ta
'            'Verificamos si los Aportes Personales guardan relación con el Haber Óptimo
'            dblAportesPersonales = De_Txt_a_Num_01(.txtHaberOptimo.Text) * 0.185
'            dblAportesPersonales = Round(dblAportesPersonales, 2)
'            If dblAportesPersonales <> De_Txt_a_Num_01(.txtJubilacion.Text) Then
'                .txtJubilacion.Text = De_Num_a_Tx_01(dblAportesPersonales, , 2)
'            End If
'            dblAportesPersonales = De_Txt_a_Num_01(.txtHaberOptimo.Text) * 0.05
'            dblAportesPersonales = Round(dblAportesPersonales, 2)
'            If dblAportesPersonales <> De_Txt_a_Num_01(.txtObraSocial.Text) Then
'                .txtObraSocial.Text = De_Num_a_Tx_01(dblAportesPersonales, , 2)
'            End If
'            'Calculamos Subtotal Descuentos
'            dblAcumulado = De_Txt_a_Num_01(.txtJubilacion.Text) + De_Txt_a_Num_01(.txtObraSocial.Text) _
'            + De_Txt_a_Num_01(.txtAdherente.Text) + De_Txt_a_Num_01(.txtSeguroObligatorio.Text) _
'            + De_Txt_a_Num_01(.txtCuotaSindical.Text) + De_Txt_a_Num_01(.txtSeguroOptativo.Text) _
'            + De_Txt_a_Num_01(.txtOtrosDescuentos.Text)
'            dblAcumulado = dblAcumulado * (-1)
'            .txtSubtotalDescuento.Text = De_Num_a_Tx_01(dblAcumulado, , 2)
'            dblAcumulado = 0
'        End With
'        dblAportesPersonales = 0
'    Case "PasoCuatro"
'        Set rstRegistroSlave = New ADODB.Recordset
'        With LiquidacionGanancia4ta
'            'Cargamos la Renta Acumulada
'            SQL = "Select * From CODIGOLIQUIDACIONES Inner Join LIQUIDACIONGANANCIAS4TACATEGORIA On CODIGOLIQUIDACIONES.Codigo = LIQUIDACIONGANANCIAS4TACATEGORIA.CodigoLiquidacion " _
'            & "Where PUESTOLABORAL = '" & .txtPuestoLaboral.Text & "' " _
'            & "And Right(PERIODO,4) = '" & Right(.txtPeriodo.Text, 4) & "' And CODIGO < '" & .txtCodigoLiquidacion.Text & "'"
'            If SQLNoMatch(SQL) = True Then
'                .txtRentaAcumulada.Text = De_Num_a_Tx_01(0)
'            Else
'                rstRegistroSlave.Open SQL, dbSlave, adOpenDynamic, adLockOptimistic
'                If rstRegistroSlave.BOF = False Then
'                    rstRegistroSlave.MoveFirst
'                    While rstRegistroSlave.EOF = False
'                        dblAcumulado = dblAcumulado + rstRegistroSlave!HaberOptimo + rstRegistroSlave!Ajuste + rstRegistroSlave!Pluriempleo
'                        rstRegistroSlave.MoveNext
'                    Wend
'                    .txtRentaAcumulada.Text = De_Num_a_Tx_01(dblAcumulado, , 2)
'                    dblAcumulado = 0
'                End If
'                rstRegistroSlave.Close
'                dblAcumulado = 0
'            End If
'            'Cargamos el Descuento Acumulado (No incluimos seguro de vida optativo)
'            SQL = "Select * From CODIGOLIQUIDACIONES Inner Join LIQUIDACIONGANANCIAS4TACATEGORIA On CODIGOLIQUIDACIONES.Codigo = LIQUIDACIONGANANCIAS4TACATEGORIA.CodigoLiquidacion " _
'            & "Where PUESTOLABORAL = '" & .txtPuestoLaboral.Text & "' " _
'            & "And Right(PERIODO,4) = '" & Right(.txtPeriodo.Text, 4) & "' And CODIGO < '" & .txtCodigoLiquidacion.Text & "'"
'            If SQLNoMatch(SQL) = True Then
'                .txtDescuentoAcumulado.Text = De_Num_a_Tx_01(0)
'            Else
'                rstRegistroSlave.Open SQL, dbSlave, adOpenForwardOnly, adLockOptimistic
'                If rstRegistroSlave.BOF = False Then
'                    rstRegistroSlave.MoveFirst
'                    While rstRegistroSlave.EOF = False
'                        dblAcumulado = dblAcumulado + rstRegistroSlave!Jubilacion + rstRegistroSlave!ObraSocial + rstRegistroSlave!AdherenteObraSocial + _
'                        rstRegistroSlave!SeguroDeVidaObligatorio + rstRegistroSlave!CuotaSindical
'                        rstRegistroSlave.MoveNext
'                    Wend
'                End If
'                dblAcumulado = dblAcumulado * (-1)
'                .txtDescuentoAcumulado.Text = De_Num_a_Tx_01(dblAcumulado, , 2)
'                rstRegistroSlave.Close
'                dblAcumulado = 0
'            End If
'            'Calculamos Ganancia Del Período
'            .txtGananciaPeriodo.Text = De_Num_a_Tx_01(De_Txt_a_Num_01(.txtSubtotalSueldo.Text) _
'            + De_Txt_a_Num_01(.txtSubtotalDescuento.Text) + De_Txt_a_Num_01(.txtSeguroOptativo.Text), , 2)
'            'Calculamos Ganancia Neta antes de Deducciones
'            .txtGananciaNeta.Text = De_Num_a_Tx_01(De_Txt_a_Num_01(.txtRentaAcumulada.Text) _
'            + De_Txt_a_Num_01(.txtDescuentoAcumulado.Text) + De_Txt_a_Num_01(.txtGananciaPeriodo.Text), , 2)
'        End With
'        Set rstRegistroSlave = Nothing
'    Case "PasoCinco"
'        ConfigurardgDeduccionesGeneralesLG4ta
'        CargardgDeduccionesGeneralesLG4ta
'    Case "PasoSeis"
'        ConfigurardgDeduccionesPersonalesLG4ta
'        CargardgDeduccionesPersonalesLG4ta
'    Case "PasoSiete"
'        Set rstRegistroSlave = New ADODB.Recordset
'        With LiquidacionGanancia4ta
'            'Calculamos la Base Imponible
'            dblAcumulado = De_Txt_a_Num_01(.txtGananciaNeta.Text) + De_Txt_a_Num_01(.dgDeduccionesGenerales.TextMatrix(6, 1)) _
'            + De_Txt_a_Num_01(.dgDeduccionesPersonales.TextMatrix(8, 1)) - De_Txt_a_Num_01(.txtSeguroOptativo.Text)
'            .txtBaseImponible.Text = De_Num_a_Tx_01(dblAcumulado, , 2)
'            dblAcumulado = 0
'            'Determinamos el Porcentaje Aplicable junto con el importe Acumulado a Retener
'            If De_Txt_a_Num_01(.txtBaseImponible.Text) > 0 Then
'                SQL = "Select * From ESCALAAPLICABLEGANANCIAS Order by IMPORTEMAXIMO Asc"
'                rstRegistroSlave.Open SQL, dbSlave, adOpenDynamic, adLockOptimistic
'                If rstRegistroSlave.BOF = False Then
'                    rstRegistroSlave.MoveFirst
'                    Do While rstRegistroSlave.EOF = False
'                        If (rstRegistroSlave!ImporteMaximo / 12 * Left(.txtPeriodo.Text, 2)) > De_Txt_a_Num_01(.txtBaseImponible.Text) Then
'                            .txtPorcentajeAplicable.Text = rstRegistroSlave!ImporteVariable * 100 & " %"
'                            Exit Do
'                        End If
'                        rstRegistroSlave.MoveNext
'                    Loop
'                    If rstRegistroSlave!ImporteFijo = 0 Then
'                        .txtSumaVariable.Text = De_Num_a_Tx_01(De_Txt_a_Num_01(.txtBaseImponible.Text) * rstRegistroSlave!ImporteVariable, , 2)
'                        .txtSumaFija.Text = De_Num_a_Tx_01(0)
'                    Else
'                        .txtSumaFija.Text = De_Num_a_Tx_01(rstRegistroSlave!ImporteFijo / 12 * Left(.txtPeriodo.Text, 2))
'                        .txtSumaVariable.Text = De_Num_a_Tx_01(rstRegistroSlave!ImporteVariable, , 2)
'                        rstRegistroSlave.MovePrevious
'                        .txtSumaVariable.Text = De_Num_a_Tx_01((De_Txt_a_Num_01(.txtBaseImponible.Text) - rstRegistroSlave!ImporteMaximo / 12 * Left(.txtPeriodo.Text, 2)) * De_Txt_a_Num_01(.txtSumaVariable.Text), , 2)
'                    End If
'                End If
'                .txtSubtotalRentencion.Text = De_Num_a_Tx_01(De_Txt_a_Num_01(.txtSumaVariable.Text) + De_Txt_a_Num_01(.txtSumaFija.Text), , 2)
'                rstRegistroSlave.Close
'            Else
'                .txtPorcentajeAplicable.Text = De_Num_a_Tx_01(0)
'                .txtSumaVariable.Text = De_Num_a_Tx_01(0)
'                .txtSumaFija.Text = De_Num_a_Tx_01(0)
'                .txtSubtotalRentencion.Text = De_Num_a_Tx_01(0)
'            End If
'            'Determinamos la Retención Acumulada
'            SQL = "Select Sum(LIQUIDACIONGANANCIAS4TACATEGORIA.Retencion) AS SumaDeImporte " _
'            & "From CODIGOLIQUIDACIONES Inner Join LIQUIDACIONGANANCIAS4TACATEGORIA On CODIGOLIQUIDACIONES.Codigo = LIQUIDACIONGANANCIAS4TACATEGORIA.CodigoLiquidacion " _
'            & "Where PUESTOLABORAL = '" & .txtPuestoLaboral.Text & "' " _
'            & "And Right(PERIODO,4) = '" & Right(.txtPeriodo.Text, 4) & "' And CODIGO < '" & .txtCodigoLiquidacion.Text & "'"
'            rstRegistroSlave.Open SQL, dbSlave, adOpenForwardOnly, adLockOptimistic
'            If rstRegistroSlave.BOF = False And IsNull(rstRegistroSlave!SumaDeImporte) = False Then
'                .txtRetencionAcumulada.Text = De_Num_a_Tx_01(rstRegistroSlave!SumaDeImporte, , 2)
'            Else
'                .txtRetencionAcumulada.Text = De_Num_a_Tx_01(0)
'            End If
'            rstRegistroSlave.Close
'            'Determinamos la Retención del Período
'            .txtRetencionPeriodo.Text = De_Num_a_Tx_01(De_Txt_a_Num_01(.txtSubtotalRentencion.Text) - De_Txt_a_Num_01(.txtRetencionAcumulada.Text) _
'            + De_Txt_a_Num_01(.txtAjuesteRetencion.Text))
'        End With
'        Set rstRegistroSlave = Nothing
'    End Select
'
'    dblAcumulado = 0
'    datFecha = 0
'    SQL = ""
'
'End Sub









'Public Sub ControlSueldoLiquidado(MomentoControl As String)
'
'    Dim SQL As String
'    Dim dblAcumuladoSLAVE As Double
'    Dim dblAcumuladoSISPER As Double
'
'    With LiquidacionGanancia4ta
'        'Calculamos la Renta Acumulada
'        SQL = "Select * From CODIGOLIQUIDACIONES Inner Join LIQUIDACIONGANANCIAS4TACATEGORIA On CODIGOLIQUIDACIONES.Codigo = LIQUIDACIONGANANCIAS4TACATEGORIA.CodigoLiquidacion " _
'        & "Where PUESTOLABORAL = '" & .txtPuestoLaboral.Text & "' " _
'        & "And Right(PERIODO,4) = '" & Right(.txtPeriodo.Text, 4) & "' And CODIGO < '" & .txtCodigoLiquidacion.Text & "'"
'        If SQLNoMatch(SQL) = False Then
'            Set rstRegistroSlave = New ADODB.Recordset
'            rstRegistroSlave.Open SQL, dbSlave, adOpenDynamic, adLockOptimistic
'            If rstRegistroSlave.BOF = False Then
'                rstRegistroSlave.MoveFirst
'                While rstRegistroSlave.EOF = False
'                    dblAcumuladoSLAVE = dblAcumuladoSLAVE + Round(rstRegistroSlave!HaberOptimo + rstRegistroSlave!Ajuste, 2)
'                    rstRegistroSlave.MoveNext
'                Wend
'            End If
'            rstRegistroSlave.Close
'            'SQL = "Select Sum(LIQUIDACIONSUELDOS.Importe) AS SumaDeImporte " _
'            '& "From CODIGOLIQUIDACIONES Inner Join LIQUIDACIONSUELDOS On CODIGOLIQUIDACIONES.Codigo = LIQUIDACIONSUELDOS.CodigoLiquidacion " _
'            '& "Where (PUESTOLABORAL = '" & .txtPuestoLaboral.Text & "' And CODIGOCONCEPTO = '0115' " _
'            '& "And Right(PERIODO,4) = '" & Right(.txtPeriodo.Text, 4) & "' And CODIGO < '" & .txtCodigoLiquidacion.Text & "') " _
'            '& "Or (PUESTOLABORAL = '" & .txtPuestoLaboral.Text & "' And CODIGOCONCEPTO = '0150' " _
'            '& "And Right(PERIODO,4) = '" & Right(.txtPeriodo.Text, 4) & "' And CODIGO < '" & .txtCodigoLiquidacion.Text & "')"
'            SQL = "Select Sum(LIQUIDACIONSUELDOS.Importe) AS SumaDeImporte " _
'            & "From CODIGOLIQUIDACIONES Inner Join LIQUIDACIONSUELDOS On CODIGOLIQUIDACIONES.Codigo = LIQUIDACIONSUELDOS.CodigoLiquidacion " _
'            & "Where PUESTOLABORAL = '" & .txtPuestoLaboral.Text & "' And CODIGOCONCEPTO = '9998' " _
'            & "And Right(PERIODO,4) = '" & Right(.txtPeriodo.Text, 4) & "' And CODIGO < '" & .txtCodigoLiquidacion.Text & "' "
'            rstRegistroSlave.Open SQL, dbSlave, adOpenDynamic, adLockOptimistic
'            If rstRegistroSlave.BOF = False Then
'                dblAcumuladoSISPER = Round(rstRegistroSlave!SumaDeImporte, 2)
'                SQL = "Select Sum(LIQUIDACIONSUELDOS.Importe) AS SumaDeImporte " _
'                & "From CODIGOLIQUIDACIONES Inner Join LIQUIDACIONSUELDOS On CODIGOLIQUIDACIONES.Codigo = LIQUIDACIONSUELDOS.CodigoLiquidacion " _
'                & "Where (PUESTOLABORAL = '" & .txtPuestoLaboral.Text & "' And CODIGOCONCEPTO = '0003' " _
'                & "And Right(PERIODO,4) = '" & Right(.txtPeriodo.Text, 4) & "' And CODIGO < '" & .txtCodigoLiquidacion.Text & "') "
'                If SQLNoMatch(SQL) = False Then
'                    rstRegistroSlave.Close
'                    rstRegistroSlave.Open SQL, dbSlave, adOpenDynamic, adLockOptimistic
'                    dblAcumuladoSISPER = dblAcumuladoSISPER - Round(rstRegistroSlave!SumaDeImporte, 2)
'                End If
'                SQL = "Select Sum(LIQUIDACIONSUELDOS.Importe) AS SumaDeImporte " _
'                & "From CODIGOLIQUIDACIONES Inner Join LIQUIDACIONSUELDOS On CODIGOLIQUIDACIONES.Codigo = LIQUIDACIONSUELDOS.CodigoLiquidacion " _
'                & "Where (PUESTOLABORAL = '" & .txtPuestoLaboral.Text & "' And CODIGOCONCEPTO = '0104' " _
'                & "And Right(PERIODO,4) = '" & Right(.txtPeriodo.Text, 4) & "' And CODIGO < '" & .txtCodigoLiquidacion.Text & "') "
'                If SQLNoMatch(SQL) = False Then
'                    rstRegistroSlave.Close
'                    rstRegistroSlave.Open SQL, dbSlave, adOpenDynamic, adLockOptimistic
'                    dblAcumuladoSISPER = dblAcumuladoSISPER - Round(rstRegistroSlave!SumaDeImporte, 2)
'                End If
'                Select Case MomentoControl
'                Case Is = "Inicial"
'                    If dblAcumuladoSLAVE <> dblAcumuladoSISPER Then
'                        MsgBox "El importe acumulado de GCIA. BRUTA (" & dblAcumuladoSLAVE & ") no coincide con lo liquidado en SISPER(" & dblAcumuladoSISPER & ")" _
'                        & vbCrLf & "Verificar que se hayan importado todas las liquidaciones del SISPER y que las Retenciones efectuadas anteriormente sean correctas", vbCritical + vbOKOnly, "VERIFICAR LIQUIDACIONES PREVIAS"
'                    End If
'                Case Is = "Final"
'                    dblAcumuladoSLAVE = dblAcumuladoSLAVE + De_Txt_a_Num_01(.txtSubtotalSueldo.Text)
'                    'Buscamos el Haber Óptimo
'                    rstRegistroSlave.Close
'                    SQL = "Select * from LIQUIDACIONSUELDOS Where PUESTOLABORAL = '" & .txtPuestoLaboral.Text & "' " _
'                    & "And CODIGOLIQUIDACION <= '" & .txtCodigoLiquidacion.Text & "' " _
'                    & "And CODIGOCONCEPTO = '9998' Order by CODIGOLIQUIDACION Desc"
'                    rstRegistroSlave.Open SQL, dbSlave, adOpenForwardOnly, adLockOptimistic
'                    rstRegistroSlave.MoveFirst
'                    dblAcumuladoSISPER = dblAcumuladoSISPER + rstRegistroSlave!Importe
'                    SQL = "Select * from LIQUIDACIONSUELDOS Where PUESTOLABORAL = '" & .txtPuestoLaboral.Text & "' " _
'                    & "And CODIGOLIQUIDACION <= '" & .txtCodigoLiquidacion.Text & "' " _
'                    & "And CODIGOCONCEPTO = '0003' Order by CODIGOLIQUIDACION Desc"
'                    If SQLNoMatch(SQL) = False Then
'                        rstRegistroSlave.Close
'                        rstRegistroSlave.Open SQL, dbSlave, adOpenForwardOnly, adLockOptimistic
'                        rstRegistroSlave.MoveFirst
'                        dblAcumuladoSISPER = dblAcumuladoSISPER - rstRegistroSlave!Importe
'                    End If
'                    If dblAcumuladoSLAVE <> dblAcumuladoSISPER Then
'                        MsgBox "El importe acumulado de GCIA. BRUTA(" & dblAcumuladoSLAVE & ") no coincide con lo liquidado en SISPER(" & dblAcumuladoSISPER & ")" _
'                        & vbCrLf & "Verificar que se hayan importado todas las liquidaciones del SISPER y que las Retenciones efectuadas anteriormente sean correctas", vbCritical + vbOKOnly, "VERIFICAR LIQUIDACIONES PREVIAS"
'                    End If
'                End Select
'            End If
'            rstRegistroSlave.Close
'            Set rstRegistroSlave = Nothing
'            dblAcumuladoSISPER = 0
'        End If
'    End With
'
'End Sub

'Public Sub LiquidacionGanancias(PasoLiquidacion As String, Optional PuestoLaboral As String, Optional CodigoLiquidacion As String)
'
'    Dim SQL As String
'    Dim datFecha As Date
'    Dim dblAcumulado As Double
'    Dim dblHaberSISPERAcumuladoEstimado As Double
'
'    Select Case PasoLiquidacion
'    Case "PasoUno" 'Carga 1)Renta Imponible y 2)Descuentos recibo INVICO (no inlcuye subtotales y llama al resto de los pasos)
'        With LiquidacionGanancia4ta
'            Set rstRegistroSlave = New ADODB.Recordset
'            'En la versión anterior se buscaba la última liquidación disponilble para trabajar sobre esa
'                'SQL = "Select CODIGOLIQUIDACION From LIQUIDACIONSUELDOS Where CODIGOLIQUIDACION <= '" & CodigoLiquidacion & "' " _
'                & "And CODIGOCONCEPTO = '0115' Order by CODIGOLIQUIDACION Desc"
'                'rstRegistroSlave.Open SQL, dbSlave, adOpenForwardOnly, adLockOptimistic
'                'rstRegistroSlave.MoveFirst
'                'CodigoLiquidacion = rstRegistroSlave!CodigoLiquidacion
'                'rstRegistroSlave.Close
'            'En la nueva versión se trabaja en la liquidación actual con el Haber Bruto al cual se le resta asignaciones familiares y discapacitados
'            SQL = "Select * from LIQUIDACIONSUELDOS Where PUESTOLABORAL = '" & PuestoLaboral & "' And CODIGOLIQUIDACION = '" & CodigoLiquidacion & "' And CODIGOCONCEPTO = '9998'"
'            rstRegistroSlave.Open SQL, dbSlave, adOpenForwardOnly, adLockOptimistic
'            dblAcumulado = rstRegistroSlave!Importe
'            rstRegistroSlave.Close
'            'Verificamos si la liquidación continene asignaciones familiares y las deducimos
'            SQL = "Select * from LIQUIDACIONSUELDOS Where PUESTOLABORAL = '" & PuestoLaboral & "' And CODIGOLIQUIDACION = '" & CodigoLiquidacion & "' And CODIGOCONCEPTO = '0003'"
'            If SQLNoMatch(SQL) = False Then
'                'Si existe, buscamos el importe y lo restamos al Total Bruto
'                rstRegistroSlave.Open SQL, dbSlave, adOpenForwardOnly, adLockOptimistic
'                dblAcumulado = dblAcumulado - rstRegistroSlave!Importe
'                rstRegistroSlave.Close
'            End If
'            'Verificamos si la liquidación continene discapacitado y las deducimos
'            SQL = "Select * from LIQUIDACIONSUELDOS Where PUESTOLABORAL = '" & PuestoLaboral & "' And CODIGOLIQUIDACION = '" & CodigoLiquidacion & "' And CODIGOCONCEPTO = '0058'"
'            If SQLNoMatch(SQL) = False Then
'                'Si existe, buscamos el importe y lo restamos al Total Bruto
'                rstRegistroSlave.Open SQL, dbSlave, adOpenForwardOnly, adLockOptimistic
'                dblAcumulado = dblAcumulado - rstRegistroSlave!Importe
'                rstRegistroSlave.Close
'            End If
'            'Ajustamos por Ley Nro 3.801 (Versión Anterior)
'                '1) Compensación Funcional
'                'SQL = "Select * from LIQUIDACIONSUELDOS Where PUESTOLABORAL = '" & PuestoLaboral & "' And CODIGOLIQUIDACION = '" & CodigoLiquidacion & "' And CODIGOCONCEPTO = '0002'"
'                'If SQLNoMatch(SQL) = False Then
'                '    rstRegistroSlave.Open SQL, dbSlave, adOpenForwardOnly, adLockOptimistic
'                '    dblAcumulado = dblAcumulado - rstRegistroSlave!Importe
'                '    rstRegistroSlave.Close
'                'End If
'                '2) Gastos de Representación
'                'SQL = "Select * from LIQUIDACIONSUELDOS Where PUESTOLABORAL = '" & PuestoLaboral & "' And CODIGOLIQUIDACION = '" & CodigoLiquidacion & "' And CODIGOCONCEPTO = '0028'"
'                'If SQLNoMatch(SQL) = False Then
'                '    rstRegistroSlave.Open SQL, dbSlave, adOpenForwardOnly, adLockOptimistic
'                '    dblAcumulado = dblAcumulado - rstRegistroSlave!Importe
'                '    rstRegistroSlave.Close
'                'End If
'                '3) Dedicación Exclusiva
'                'SQL = "Select * from LIQUIDACIONSUELDOS Where PUESTOLABORAL = '" & PuestoLaboral & "' And CODIGOLIQUIDACION = '" & CodigoLiquidacion & "' And CODIGOCONCEPTO = '0064'"
'                'If SQLNoMatch(SQL) = False Then
'                '    rstRegistroSlave.Open SQL, dbSlave, adOpenForwardOnly, adLockOptimistic
'                '    dblAcumulado = dblAcumulado - rstRegistroSlave!Importe
'                '    rstRegistroSlave.Close
'                'End If
'            .txtHaberOptimo.Text = De_Num_a_Tx_01(dblAcumulado, , 2)
'            If rstRegistroSlave.State = adStateOpen Then
'                rstRegistroSlave.Close
'            End If
'            dblAcumulado = 0
'            'Cargamos Sueldo Pluriempleo
'            SQL = "Select * from LIQUIDACIONGANANCIAS4TACATEGORIA" _
'            & " Where PUESTOLABORAL = '" & PuestoLaboral _
'            & "' And CODIGOLIQUIDACION <= '" & CodigoLiquidacion _
'            & "' And PLURIEMPLEO  <> 0" _
'            & " Order By CODIGOLIQUIDACION Desc"
'            If SQLNoMatch(SQL) = True Then
'               .txtPluriempleo.Text = De_Num_a_Tx_01(0)
'            Else
'                rstRegistroSlave.Open SQL, dbSlave, adOpenForwardOnly, adLockOptimistic
'                .txtPluriempleo.Text = De_Num_a_Tx_01(rstRegistroSlave!Pluriempleo, , 2)
'                .txtPluriempleo.BackColor = &HFF&
'                rstRegistroSlave.Close
'            End If
'            'Cargamos el Ajuste
'            .txtAjuste.Text = De_Num_a_Tx_01(0)
'            'Buscamos el Descuento por Jubilacion
'            SQL = "Select * from LIQUIDACIONSUELDOS Where PUESTOLABORAL = '" & PuestoLaboral & "' And CODIGOLIQUIDACION = '" & CodigoLiquidacion & "' And CODIGOCONCEPTO = '0208'"
'            rstRegistroSlave.Open SQL, dbSlave, adOpenForwardOnly, adLockOptimistic
'            .txtJubilacion.Text = De_Num_a_Tx_01(rstRegistroSlave!Importe, , 2)
'            rstRegistroSlave.Close
'            'Buscamos el Descuento por ObraSocial
'            SQL = "Select * from LIQUIDACIONSUELDOS Where PUESTOLABORAL = '" & PuestoLaboral & "' And CODIGOLIQUIDACION = '" & CodigoLiquidacion & "' And CODIGOCONCEPTO = '0212'"
'            rstRegistroSlave.Open SQL, dbSlave, adOpenForwardOnly, adLockOptimistic
'            .txtObraSocial.Text = De_Num_a_Tx_01(rstRegistroSlave!Importe, , 2)
'            rstRegistroSlave.Close
'            'Cargamos el Adherente Obra Social
'            SQL = "Select * from LIQUIDACIONSUELDOS Where PUESTOLABORAL = '" & PuestoLaboral & "' And CODIGOLIQUIDACION = '" & CodigoLiquidacion & "' And CODIGOCONCEPTO = '0234'"
'            If SQLNoMatch(SQL) = True Then
'                .txtAdherente.Text = De_Num_a_Tx_01(0)
'            Else
'                rstRegistroSlave.Open SQL, dbSlave, adOpenForwardOnly, adLockOptimistic
'                .txtAdherente.Text = De_Num_a_Tx_01(rstRegistroSlave!Importe, , 2)
'                rstRegistroSlave.Close
'                SQL = "Select * from LIQUIDACIONGANANCIAS4TACATEGORIA Where PUESTOLABORAL = '" & PuestoLaboral & "' And CODIGOLIQUIDACION <= '" & CodigoLiquidacion & "' " _
'                & "Order by CODIGOLIQUIDACION Desc"
'                If SQLNoMatch(SQL) = False Then
'                    rstRegistroSlave.Open SQL, dbSlave, adOpenForwardOnly, adLockOptimistic
'                    rstRegistroSlave.MoveFirst
'                    If rstRegistroSlave!AdherenteObraSocial <> De_Txt_a_Num_01(.txtAdherente.Text) Then
'                        .txtAdherente.Text = De_Num_a_Tx_01(rstRegistroSlave!AdherenteObraSocial)
'                        .txtAdherente.BackColor = &HFF&
'                    End If
'                    rstRegistroSlave.Close
'                End If
'            End If
'            'Buscamos el Descuento Seguro Obligatorio
'            SQL = "Select * from LIQUIDACIONSUELDOS Where ((PUESTOLABORAL = '" & PuestoLaboral & "' And CODIGOLIQUIDACION = '" & CodigoLiquidacion & "' And CODIGOCONCEPTO = '0360') Or " _
'            & "(PUESTOLABORAL = '" & PuestoLaboral & "' And CODIGOLIQUIDACION = '" & CodigoLiquidacion & "' And CODIGOCONCEPTO = '0366') Or " _
'            & "(PUESTOLABORAL = '" & PuestoLaboral & "' And CODIGOLIQUIDACION = '" & CodigoLiquidacion & "' And CODIGOCONCEPTO = '0369'))"
'            If SQLNoMatch(SQL) = True Then
'                .txtSeguroObligatorio.Text = De_Num_a_Tx_01(0)
'            Else
'                rstRegistroSlave.Open SQL, dbSlave, adOpenForwardOnly, adLockOptimistic
'                .txtSeguroObligatorio.Text = De_Num_a_Tx_01(rstRegistroSlave!Importe, , 2)
'                rstRegistroSlave.Close
'            End If
'            'Buscamos el Descuento Cuota Sindical
'            SQL = "Select * from LIQUIDACIONSUELDOS Where ((PUESTOLABORAL = '" & PuestoLaboral & "' And CODIGOLIQUIDACION = '" & CodigoLiquidacion & "' And CODIGOCONCEPTO = '0219') Or " _
'            & "(PUESTOLABORAL = '" & PuestoLaboral & "' And CODIGOLIQUIDACION = '" & CodigoLiquidacion & "' And CODIGOCONCEPTO = '0227'))"
'            If SQLNoMatch(SQL) = True Then
'                .txtCuotaSindical.Text = De_Num_a_Tx_01(0)
'            Else
'                rstRegistroSlave.Open SQL, dbSlave, adOpenForwardOnly, adLockOptimistic
'                .txtCuotaSindical.Text = De_Num_a_Tx_01(rstRegistroSlave!Importe, , 2)
'                rstRegistroSlave.Close
'            End If
'            'Buscamos el Descuento Seguro Optativo
'            SQL = "Select * from LIQUIDACIONSUELDOS Where ((PUESTOLABORAL = '" & PuestoLaboral & "' And CODIGOLIQUIDACION = '" & CodigoLiquidacion & "' And CODIGOCONCEPTO = '0317') Or " _
'            & "(PUESTOLABORAL = '" & PuestoLaboral & "' And CODIGOLIQUIDACION = '" & CodigoLiquidacion & "' And CODIGOCONCEPTO = '0361') Or " _
'            & "(PUESTOLABORAL = '" & PuestoLaboral & "' And CODIGOLIQUIDACION = '" & CodigoLiquidacion & "' And CODIGOCONCEPTO = '0367') Or " _
'            & "(PUESTOLABORAL = '" & PuestoLaboral & "' And CODIGOLIQUIDACION = '" & CodigoLiquidacion & "' And CODIGOCONCEPTO = '0370') Or " _
'            & "(PUESTOLABORAL = '" & PuestoLaboral & "' And CODIGOLIQUIDACION = '" & CodigoLiquidacion & "' And CODIGOCONCEPTO = '0373') Or " _
'            & "(PUESTOLABORAL = '" & PuestoLaboral & "' And CODIGOLIQUIDACION = '" & CodigoLiquidacion & "' And CODIGOCONCEPTO = '0374'))"
'            If SQLNoMatch(SQL) = True Then
'               .txtSeguroOptativo.Text = De_Num_a_Tx_01(0)
'            Else
'                'Buscamos el límite de Seguro de Vida
'                datFecha = DateTime.DateSerial(Right(.txtPeriodo.Text, 4), Left(.txtPeriodo.Text, 2), 1)
'                rstRegistroSlave.Open SQL, dbSlave, adOpenForwardOnly, adLockOptimistic
'                Set rstBuscarSlave = New ADODB.Recordset
'                rstBuscarSlave.Open "Select SEGURODEVIDA From DEDUCCIONES4TACATEGORIA Where FECHA <= #" & Format(datFecha, "MM/DD/YYYY") & "# Order by FECHA Desc", dbSlave, adOpenForwardOnly, adLockOptimistic
'                If rstRegistroSlave!Importe < (rstBuscarSlave!SeguroDeVida / 12) Then
'                    .txtSeguroOptativo.Text = De_Num_a_Tx_01(rstRegistroSlave!Importe, , 2)
'                Else
'                    .txtSeguroOptativo.Text = De_Num_a_Tx_01((rstBuscarSlave!SeguroDeVida / 12), , 2)
'                End If
'                rstRegistroSlave.Close
'                rstBuscarSlave.Close
'                Set rstBuscarSlave = Nothing
'            End If
'            'Cargamos Otros Descuentos (VARIABLE SIN USO)
'            .txtOtrosDescuentos.Text = De_Num_a_Tx_01(0)
'            'Cargamos Ajuste Retenciones
'            SQL = "Select * from LIQUIDACIONGANANCIAS4TACATEGORIA Where PUESTOLABORAL = '" & PuestoLaboral & "' And CODIGOLIQUIDACION < '" & .txtCodigoLiquidacion.Text & "' " _
'            & "Order By CODIGOLIQUIDACION Desc"
'            rstRegistroSlave.Open SQL, dbSlave, adOpenForwardOnly, adLockOptimistic
'            If SQLNoMatch(SQL) = True Then
'               .txtAjuesteRetencion.Text = De_Num_a_Tx_01(0)
'            ElseIf rstRegistroSlave!AjusteRetencion = 0 Then
'               .txtAjuesteRetencion.Text = De_Num_a_Tx_01(0)
'            Else
'                .txtAjuesteRetencion.Text = De_Num_a_Tx_01(rstRegistroSlave!AjusteRetencion, , 2)
'                .txtAjuesteRetencion.BackColor = &HFF&
'            End If
'            rstRegistroSlave.Close
'            Set rstRegistroSlave = Nothing
'            Call ControlSueldoLiquidado("Inicial") 'Verificar
'            Call LiquidacionGanancias("PasoDos")
'            Call LiquidacionGanancias("PasoTres")
'            Call LiquidacionGanancias("PasoCuatro")
'            Call LiquidacionGanancias("PasoCinco")
'            Call LiquidacionGanancias("PasoSeis")
'            Call LiquidacionGanancias("PasoSiete")
'        End With
'    Case "PasoDos" 'Calculo Subtotales 1)Renta Imponible y 2)Descuentos Recibo INVICO
'        With LiquidacionGanancia4ta
'            'Calculamos Subtotal Sueldo
'            .txtSubtotalSueldo.Text = De_Num_a_Tx_01(De_Txt_a_Num_01(.txtHaberOptimo.Text) _
'            + De_Txt_a_Num_01(.txtPluriempleo.Text) + De_Txt_a_Num_01(.txtAjuste.Text), , 2)
'            'Calculamos Subtotal Descuentos
'            dblAcumulado = De_Txt_a_Num_01(.txtJubilacion.Text) + De_Txt_a_Num_01(.txtObraSocial.Text) _
'            + De_Txt_a_Num_01(.txtAdherente.Text) + De_Txt_a_Num_01(.txtSeguroObligatorio.Text) _
'            + De_Txt_a_Num_01(.txtCuotaSindical.Text) + De_Txt_a_Num_01(.txtSeguroOptativo.Text) _
'            + De_Txt_a_Num_01(.txtOtrosDescuentos.Text)
'            dblAcumulado = dblAcumulado * (-1)
'            .txtSubtotalDescuento.Text = De_Num_a_Tx_01(dblAcumulado, , 2)
'            dblAcumulado = 0
'        End With
'    Case "PasoTres" 'Recalculo Descuento Jubilación y O.Social en función del Haber Óptimo
'        Dim dblAportesPersonales As Double
'        With LiquidacionGanancia4ta
'            'Verificamos si los Aportes Personales guardan relación con el Haber Óptimo
'            dblAportesPersonales = De_Txt_a_Num_01(.txtHaberOptimo.Text) * 0.185
'            dblAportesPersonales = Round(dblAportesPersonales, 2)
'            If dblAportesPersonales <> De_Txt_a_Num_01(.txtJubilacion.Text) Then
'                .txtJubilacion.Text = De_Num_a_Tx_01(dblAportesPersonales, , 2)
'            End If
'            dblAportesPersonales = De_Txt_a_Num_01(.txtHaberOptimo.Text) * 0.05
'            dblAportesPersonales = Round(dblAportesPersonales, 2)
'            If dblAportesPersonales <> De_Txt_a_Num_01(.txtObraSocial.Text) Then
'                .txtObraSocial.Text = De_Num_a_Tx_01(dblAportesPersonales, , 2)
'            End If
'            'Calculamos Subtotal Descuentos
'            dblAcumulado = De_Txt_a_Num_01(.txtJubilacion.Text) + De_Txt_a_Num_01(.txtObraSocial.Text) _
'            + De_Txt_a_Num_01(.txtAdherente.Text) + De_Txt_a_Num_01(.txtSeguroObligatorio.Text) _
'            + De_Txt_a_Num_01(.txtCuotaSindical.Text) + De_Txt_a_Num_01(.txtSeguroOptativo.Text) _
'            + De_Txt_a_Num_01(.txtOtrosDescuentos.Text)
'            dblAcumulado = dblAcumulado * (-1)
'            .txtSubtotalDescuento.Text = De_Num_a_Tx_01(dblAcumulado, , 2)
'            dblAcumulado = 0
'        End With
'        dblAportesPersonales = 0
'    Case "PasoCuatro" '3)Gcia. Neta Acumulada
'        Set rstRegistroSlave = New ADODB.Recordset
'        With LiquidacionGanancia4ta
'            'Cargamos la Renta Acumulada
'            SQL = "Select * From CODIGOLIQUIDACIONES Inner Join LIQUIDACIONGANANCIAS4TACATEGORIA On CODIGOLIQUIDACIONES.Codigo = LIQUIDACIONGANANCIAS4TACATEGORIA.CodigoLiquidacion " _
'            & "Where PUESTOLABORAL = '" & .txtPuestoLaboral.Text & "' " _
'            & "And Right(PERIODO,4) = '" & Right(.txtPeriodo.Text, 4) & "' And CODIGO < '" & .txtCodigoLiquidacion.Text & "'"
'            If SQLNoMatch(SQL) = True Then
'                .txtRentaAcumulada.Text = De_Num_a_Tx_01(0)
'            Else
'                rstRegistroSlave.Open SQL, dbSlave, adOpenDynamic, adLockOptimistic
'                If rstRegistroSlave.BOF = False Then
'                    rstRegistroSlave.MoveFirst
'                    While rstRegistroSlave.EOF = False
'                        dblAcumulado = dblAcumulado + rstRegistroSlave!HaberOptimo + rstRegistroSlave!Ajuste + rstRegistroSlave!Pluriempleo
'                        rstRegistroSlave.MoveNext
'                    Wend
'                    .txtRentaAcumulada.Text = De_Num_a_Tx_01(dblAcumulado, , 2)
'                    dblAcumulado = 0
'                End If
'                rstRegistroSlave.Close
'                dblAcumulado = 0
'            End If
'            'Cargamos el Descuento Acumulado (No incluimos seguro de vida optativo)
'            SQL = "Select * From CODIGOLIQUIDACIONES Inner Join LIQUIDACIONGANANCIAS4TACATEGORIA On CODIGOLIQUIDACIONES.Codigo = LIQUIDACIONGANANCIAS4TACATEGORIA.CodigoLiquidacion " _
'            & "Where PUESTOLABORAL = '" & .txtPuestoLaboral.Text & "' " _
'            & "And Right(PERIODO,4) = '" & Right(.txtPeriodo.Text, 4) & "' And CODIGO < '" & .txtCodigoLiquidacion.Text & "'"
'            If SQLNoMatch(SQL) = True Then
'                .txtDescuentoAcumulado.Text = De_Num_a_Tx_01(0)
'            Else
'                rstRegistroSlave.Open SQL, dbSlave, adOpenForwardOnly, adLockOptimistic
'                If rstRegistroSlave.BOF = False Then
'                    rstRegistroSlave.MoveFirst
'                    While rstRegistroSlave.EOF = False
'                        dblAcumulado = dblAcumulado + rstRegistroSlave!Jubilacion + rstRegistroSlave!ObraSocial + rstRegistroSlave!AdherenteObraSocial + _
'                        rstRegistroSlave!SeguroDeVidaObligatorio + rstRegistroSlave!CuotaSindical
'                        rstRegistroSlave.MoveNext
'                    Wend
'                End If
'                dblAcumulado = dblAcumulado * (-1)
'                .txtDescuentoAcumulado.Text = De_Num_a_Tx_01(dblAcumulado, , 2)
'                rstRegistroSlave.Close
'                dblAcumulado = 0
'            End If
'            'Calculamos Ganancia Del Período
'            .txtGananciaPeriodo.Text = De_Num_a_Tx_01(De_Txt_a_Num_01(.txtSubtotalSueldo.Text) _
'            + De_Txt_a_Num_01(.txtSubtotalDescuento.Text) + De_Txt_a_Num_01(.txtSeguroOptativo.Text), , 2)
'            'Calculamos Ganancia Neta antes de Deducciones
'            .txtGananciaNeta.Text = De_Num_a_Tx_01(De_Txt_a_Num_01(.txtRentaAcumulada.Text) _
'            + De_Txt_a_Num_01(.txtDescuentoAcumulado.Text) + De_Txt_a_Num_01(.txtGananciaPeriodo.Text), , 2)
'        End With
'        Set rstRegistroSlave = Nothing
'    Case "PasoCinco" '5)Deducciones Generales Acum.
'        ConfigurardgDeduccionesGeneralesLG4ta
'        CargardgDeduccionesGeneralesLG4ta
'    Case "PasoSeis" '4)Deducciones Personales Acum.
'        ConfigurardgDeduccionesPersonalesLG4ta
'        CargardgDeduccionesPersonalesLG4ta
'    Case "PasoSiete" '6)Importe a retener Gcias.
'        Set rstRegistroSlave = New ADODB.Recordset
'        With LiquidacionGanancia4ta
'            'Calculamos la Base Imponible
'            dblAcumulado = De_Txt_a_Num_01(.txtGananciaNeta.Text) + De_Txt_a_Num_01(.dgDeduccionesGenerales.TextMatrix(6, 1)) _
'            + De_Txt_a_Num_01(.dgDeduccionesPersonales.TextMatrix(8, 1)) - De_Txt_a_Num_01(.txtSeguroOptativo.Text)
'            .txtBaseImponible.Text = De_Num_a_Tx_01(dblAcumulado, , 2)
'            dblAcumulado = 0
'            'Determinamos el Porcentaje Aplicable junto con el importe Acumulado a Retener
'            If De_Txt_a_Num_01(.txtBaseImponible.Text) > 0 Then
'                datFecha = BuscarUltimoDiaDelPeriodo(.txtPeriodo.Text)
'                SQL = "Select * From ESCALAAPLICABLEGANANCIAS" _
'                & " Where FECHA <= #" & Format(datFecha, "MM/DD/YYYY") _
'                & "# Order by FECHA Desc, IMPORTEMAXIMO Asc"
'                rstRegistroSlave.Open SQL, dbSlave, adOpenDynamic, adLockOptimistic
'                If rstRegistroSlave.BOF = False Then
'                    rstRegistroSlave.MoveFirst
'                    Do While rstRegistroSlave.EOF = False
'                        If (rstRegistroSlave!ImporteMaximo / 12 * Left(.txtPeriodo.Text, 2)) > De_Txt_a_Num_01(.txtBaseImponible.Text) Then
'                            .txtPorcentajeAplicable.Text = rstRegistroSlave!ImporteVariable * 100 & " %"
'                            Exit Do
'                        End If
'                        rstRegistroSlave.MoveNext
'                    Loop
'                    If rstRegistroSlave!ImporteFijo = 0 Then
'                        .txtSumaVariable.Text = De_Num_a_Tx_01(De_Txt_a_Num_01(.txtBaseImponible.Text) * rstRegistroSlave!ImporteVariable, , 2)
'                        .txtSumaFija.Text = De_Num_a_Tx_01(0)
'                    Else
'                        .txtSumaFija.Text = De_Num_a_Tx_01(rstRegistroSlave!ImporteFijo / 12 * Left(.txtPeriodo.Text, 2))
'                        .txtSumaVariable.Text = De_Num_a_Tx_01(rstRegistroSlave!ImporteVariable, , 2)
'                        rstRegistroSlave.MovePrevious
'                        .txtSumaVariable.Text = De_Num_a_Tx_01((De_Txt_a_Num_01(.txtBaseImponible.Text) - rstRegistroSlave!ImporteMaximo / 12 * Left(.txtPeriodo.Text, 2)) * De_Txt_a_Num_01(.txtSumaVariable.Text), , 2)
'                    End If
'                End If
'                .txtSubtotalRentencion.Text = De_Num_a_Tx_01(De_Txt_a_Num_01(.txtSumaVariable.Text) + De_Txt_a_Num_01(.txtSumaFija.Text), , 2)
'                rstRegistroSlave.Close
'            Else
'                .txtPorcentajeAplicable.Text = De_Num_a_Tx_01(0)
'                .txtSumaVariable.Text = De_Num_a_Tx_01(0)
'                .txtSumaFija.Text = De_Num_a_Tx_01(0)
'                .txtSubtotalRentencion.Text = De_Num_a_Tx_01(0)
'            End If
'            'Determinamos la Retención Acumulada
'            SQL = "Select Sum(LIQUIDACIONGANANCIAS4TACATEGORIA.Retencion) AS SumaDeImporte " _
'            & "From CODIGOLIQUIDACIONES Inner Join LIQUIDACIONGANANCIAS4TACATEGORIA On CODIGOLIQUIDACIONES.Codigo = LIQUIDACIONGANANCIAS4TACATEGORIA.CodigoLiquidacion " _
'            & "Where PUESTOLABORAL = '" & .txtPuestoLaboral.Text & "' " _
'            & "And Right(PERIODO,4) = '" & Right(.txtPeriodo.Text, 4) & "' And CODIGO < '" & .txtCodigoLiquidacion.Text & "'"
'            rstRegistroSlave.Open SQL, dbSlave, adOpenForwardOnly, adLockOptimistic
'            If rstRegistroSlave.BOF = False And IsNull(rstRegistroSlave!SumaDeImporte) = False Then
'                .txtRetencionAcumulada.Text = De_Num_a_Tx_01(rstRegistroSlave!SumaDeImporte, , 2)
'            Else
'                .txtRetencionAcumulada.Text = De_Num_a_Tx_01(0)
'            End If
'            rstRegistroSlave.Close
'            'Determinamos la Retención del Período
'            .txtRetencionPeriodo.Text = De_Num_a_Tx_01(De_Txt_a_Num_01(.txtSubtotalRentencion.Text) - De_Txt_a_Num_01(.txtRetencionAcumulada.Text) _
'            + De_Txt_a_Num_01(.txtAjuesteRetencion.Text))
'        End With
'        Set rstRegistroSlave = Nothing
'    End Select
'
'    dblAcumulado = 0
'    datFecha = 0
'    SQL = ""
'
'End Sub
'
Public Function TieneDeduccionGeneral(Deduccion As String, PuestoLaboral As String, FechaTope As Date) As Boolean

    Dim SQL As String

    SQL = "Select " & Deduccion & " From IMPORTEDEDUCCIONESGENERALES Where PUESTOLABORAL = '" & PuestoLaboral & "' " _
    & "And FECHA <= #" & Format(FechaTope, "MM/DD/YYYY") & "# Order by FECHA Desc"
    If SQLNoMatch(SQL) = True Then
        TieneDeduccionGeneral = False
    Else
        Set rstBuscarSlave = New ADODB.Recordset
        rstBuscarSlave.Open SQL, dbSlave, adOpenDynamic, adLockOptimistic
        If rstBuscarSlave.Fields(0) <> 0 Then
            TieneDeduccionGeneral = True
        Else
            TieneDeduccionGeneral = False
        End If
        rstBuscarSlave.Close
        Set rstBuscarSlave = Nothing
    End If

    SQL = ""

End Function

