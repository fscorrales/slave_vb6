Attribute VB_Name = "ConfigurarFomulario"
Public strCargaCodigoSIRADIG As String
Public strListadoCodigoSIRADIG As String
Public dblDiferenciaLiquidacionPublica As Double

Public Sub CenterMe(frmForm As Form, Ancho As Integer, Alto As Integer)
    With frmForm
        .Width = Ancho
        .Height = Alto
    End With
    frmForm.Left = (Screen.Width - frmForm.Width) / 2
    frmForm.Top = (Screen.Height - frmForm.Height) / 2
End Sub

Public Sub CargaCompletaLiquidacionGanancia4ta(PuestoLaboral As String, CodigoLiquidacion As String)

    Call CargarLiquidacionGanancia4ta("RentaImponible", PuestoLaboral, CodigoLiquidacion)
    Call CargarLiquidacionGanancia4ta("DescuentosRecibo", PuestoLaboral, CodigoLiquidacion)
    Call CargarLiquidacionGanancia4ta("RentayDescuentosAcumulados", PuestoLaboral, CodigoLiquidacion)
    Call CargarLiquidacionGanancia4ta("SueldoNeto", PuestoLaboral, CodigoLiquidacion)
    Call CargarLiquidacionGanancia4ta("GananciaNetaAcumulada", PuestoLaboral, CodigoLiquidacion)
    Call CargarLiquidacionGanancia4ta("DeduccionesGenerales", PuestoLaboral, CodigoLiquidacion)
    Call CargarLiquidacionGanancia4ta("DeduccionesPersonales", PuestoLaboral, CodigoLiquidacion)
    Call CargarLiquidacionGanancia4ta("RetencionTotal", PuestoLaboral, CodigoLiquidacion)
    Call CargarLiquidacionGanancia4ta("RetencionAcumulada", PuestoLaboral, CodigoLiquidacion)
    Call CargarLiquidacionGanancia4ta("AjusteRetencion", PuestoLaboral, CodigoLiquidacion)
    Call CargarLiquidacionGanancia4ta("RetencionPeriodo", PuestoLaboral, CodigoLiquidacion)
    
End Sub

Public Sub CargaCompletaLiquidacionGanancia4taSIRADIG(PuestoLaboral As String, CodigoLiquidacion As String)

    Call CargarLiquidacionGanancia4taSIRADIG("RentaImponible", PuestoLaboral, CodigoLiquidacion)
    Call CargarLiquidacionGanancia4taSIRADIG("DescuentosRecibo", PuestoLaboral, CodigoLiquidacion)
    Call CargarLiquidacionGanancia4taSIRADIG("GananciaBrutaAcumulada", PuestoLaboral, CodigoLiquidacion)
    Call CargarLiquidacionGanancia4taSIRADIG("DeduccionesGenerales", PuestoLaboral, CodigoLiquidacion)
    Call CargarLiquidacionGanancia4taSIRADIG("DeduccionesPersonales", PuestoLaboral, CodigoLiquidacion)
    Call CargarLiquidacionGanancia4taSIRADIG("RetencionTotal", PuestoLaboral, CodigoLiquidacion)
    Call CargarLiquidacionGanancia4taSIRADIG("RetencionAcumulada", PuestoLaboral, CodigoLiquidacion)
    Call CargarLiquidacionGanancia4taSIRADIG("AjusteRetencion", PuestoLaboral, CodigoLiquidacion)
    Call CargarLiquidacionGanancia4taSIRADIG("RetencionPeriodo", PuestoLaboral, CodigoLiquidacion)
    
End Sub


Public Sub RecalcularLiquidacionGanancia4ta(PuestoLaboral As String, _
CodigoLiquidacion As String, Optional RecalcularDescuentosObligatorios As Boolean = True, _
Optional LiquidacionFinal As Boolean = False)

    Dim dblImporteaIncorporar As Double
    Dim dblDiferenciaLiquidacion As Double
    Dim Ajustar As Integer
    
    dblImporteaIncorporar = De_Txt_a_Num_01(LiquidacionGanancia4ta.txtHaberOptimo.Text)
    dblDiferenciaLiquidacion = DiferenciaHaberBrutoAcumuladoSISPERvsSLAVE(PuestoLaboral, CodigoLiquidacion, dblImporteaIncorporar, LiquidacionFinal)
    If dblDiferenciaLiquidacion <> 0 Then
        Ajustar = MsgBox("Se ha detectado una diferencia de $ " & dblDiferenciaLiquidacion & " en el importe acumulado de Haber Bruto entre SISPER y SLAVE" & vbCrLf & "Desea ajustar el Haber Óptimo del período en base a la diferencia detectada?", vbQuestion + vbYesNo, "DIFERENCIA ACUMULADA SISPER VS SLAVE")
        If Ajustar = 6 Then
            dblImporteaIncorporar = dblImporteaIncorporar + dblDiferenciaLiquidacion
            LiquidacionGanancia4ta.txtHaberOptimo.Text = De_Num_a_Tx_01(dblImporteaIncorporar)
            RecalcularDescuentosObligatorios = True
        End If
    End If
    dblImporteaIncorporar = 0
    dblDiferenciaLiquidacion = 0
    
    Call CargarLiquidacionGanancia4ta("SubtotalSueldo", PuestoLaboral, CodigoLiquidacion)
    If RecalcularDescuentosObligatorios = True Then
        Call CargarLiquidacionGanancia4ta("RecalculoDescuentosObligatorios", PuestoLaboral, CodigoLiquidacion)
    End If
    
    dblImporteaIncorporar = De_Txt_a_Num_01(LiquidacionGanancia4ta.txtCuotaSindical.Text)
    If dblImporteaIncorporar > 0 Then
        dblDiferenciaLiquidacion = DiferenciaCuotaSindicalAcumuladaSISPERvsSLAVE(PuestoLaboral, CodigoLiquidacion, dblImporteaIncorporar, LiquidacionFinal)
        If dblDiferenciaLiquidacion <> 0 Then
            Ajustar = MsgBox("Se ha detectado una diferencia de $ " & dblDiferenciaLiquidacion & " en el importe acumulado de CuotaSindical entre SISPER y SLAVE" & vbCrLf & "Desea ajustar la Cuota Sindical del período en base a la diferencia detectada?", vbQuestion + vbYesNo, "DIFERENCIA ACUMULADA SISPER VS SLAVE")
            If Ajustar = 6 Then
                dblImporteaIncorporar = dblImporteaIncorporar + dblDiferenciaLiquidacion
                LiquidacionGanancia4ta.txtCuotaSindical.Text = De_Num_a_Tx_01(dblImporteaIncorporar)
            End If
        End If
    End If
    dblImporteaIncorporar = 0
    dblDiferenciaLiquidacion = 0

    Call CargarLiquidacionGanancia4ta("SubtotalDescuentos", PuestoLaboral, CodigoLiquidacion)
    Call CargarLiquidacionGanancia4ta("SueldoNeto", PuestoLaboral, CodigoLiquidacion)
    Call CargarLiquidacionGanancia4ta("GananciaNetaAcumulada", PuestoLaboral, CodigoLiquidacion)
    Call CargarLiquidacionGanancia4ta("DeduccionesGenerales", PuestoLaboral, CodigoLiquidacion, LiquidacionFinal)
    Call CargarLiquidacionGanancia4ta("DeduccionesPersonales", PuestoLaboral, CodigoLiquidacion, LiquidacionFinal)
    Call CargarLiquidacionGanancia4ta("RetencionTotal", PuestoLaboral, CodigoLiquidacion)
    Call CargarLiquidacionGanancia4ta("RetencionPeriodo", PuestoLaboral, CodigoLiquidacion)
    
End Sub

Public Sub RecalcularLiquidacionGanancia4taSIRADIG(PuestoLaboral As String, _
CodigoLiquidacion As String, Optional RecalcularDescuentosObligatorios As Boolean = True, _
Optional LiquidacionFinal As Boolean = False)

    Dim dblImporteaIncorporar As Double
    Dim dblDiferenciaLiquidacion As Double
    Dim Ajustar As Integer
    
    dblImporteaIncorporar = De_Txt_a_Num_01(LiquidacionGanancia4taSIRADIG.txtHaberOptimo.Text)
    dblDiferenciaLiquidacion = DiferenciaHaberBrutoAcumuladoSISPERvsSLAVE(PuestoLaboral, CodigoLiquidacion, dblImporteaIncorporar, LiquidacionFinal)
    If dblDiferenciaLiquidacion <> 0 Then
        Ajustar = MsgBox("Se ha detectado una diferencia de $ " & dblDiferenciaLiquidacion & " en el importe acumulado de Haber Bruto entre SISPER y SLAVE" & vbCrLf & "Desea ajustar el Haber Óptimo del período en base a la diferencia detectada?", vbQuestion + vbYesNo, "DIFERENCIA ACUMULADA SISPER VS SLAVE")
        If Ajustar = 6 Then
            dblImporteaIncorporar = dblImporteaIncorporar + dblDiferenciaLiquidacion
            LiquidacionGanancia4taSIRADIG.txtHaberOptimo.Text = De_Num_a_Tx_01(dblImporteaIncorporar)
            RecalcularDescuentosObligatorios = True
        End If
    End If
    dblImporteaIncorporar = 0
    dblDiferenciaLiquidacion = 0
    
    Call CargarLiquidacionGanancia4taSIRADIG("SubtotalSueldo", PuestoLaboral, CodigoLiquidacion)
    If RecalcularDescuentosObligatorios = True Then
        Call CargarLiquidacionGanancia4taSIRADIG("RecalculoDescuentosObligatorios", PuestoLaboral, CodigoLiquidacion)
    End If
    
    dblImporteaIncorporar = De_Txt_a_Num_01(LiquidacionGanancia4taSIRADIG.txtCuotaSindical.Text)
    If dblImporteaIncorporar > 0 Then
        dblDiferenciaLiquidacion = DiferenciaCuotaSindicalAcumuladaSISPERvsSLAVE(PuestoLaboral, CodigoLiquidacion, dblImporteaIncorporar, LiquidacionFinal)
        If dblDiferenciaLiquidacion <> 0 Then
            Ajustar = MsgBox("Se ha detectado una diferencia de $ " & dblDiferenciaLiquidacion & " en el importe acumulado de CuotaSindical entre SISPER y SLAVE" & vbCrLf & "Desea ajustar la Cuota Sindical del período en base a la diferencia detectada?", vbQuestion + vbYesNo, "DIFERENCIA ACUMULADA SISPER VS SLAVE")
            If Ajustar = 6 Then
                dblImporteaIncorporar = dblImporteaIncorporar + dblDiferenciaLiquidacion
                LiquidacionGanancia4taSIRADIG.txtCuotaSindical.Text = De_Num_a_Tx_01(dblImporteaIncorporar)
            End If
        End If
    End If
    dblImporteaIncorporar = 0
    dblDiferenciaLiquidacion = 0

    Call CargarLiquidacionGanancia4taSIRADIG("SubtotalDescuentos", PuestoLaboral, CodigoLiquidacion)
'    Call CargarLiquidacionGanancia4taSIRADIG("SueldoNeto", PuestoLaboral, CodigoLiquidacion)
    Call CargarLiquidacionGanancia4taSIRADIG("GananciaBrutaAcumulada", PuestoLaboral, CodigoLiquidacion)
    Call CargarLiquidacionGanancia4taSIRADIG("DeduccionesGenerales", PuestoLaboral, CodigoLiquidacion, LiquidacionFinal)
    Call CargarLiquidacionGanancia4taSIRADIG("DeduccionesPersonales", PuestoLaboral, CodigoLiquidacion, LiquidacionFinal)
    Call CargarLiquidacionGanancia4taSIRADIG("RetencionTotal", PuestoLaboral, CodigoLiquidacion)
    Call CargarLiquidacionGanancia4taSIRADIG("RetencionPeriodo", PuestoLaboral, CodigoLiquidacion)
    
End Sub

Public Sub EditandoLiquidacionGanancia4ta(PuestoLaboral As String, CodigoLiquidacion As String)

    Dim SQL As String
    Dim dblImporteCalculado As Double
    Dim dblImporteAcumulado As Double
    
    'Creamos la consulta base que va a cargar el formulario
    SQL = "Select * From LIQUIDACIONGANANCIAS4TACATEGORIA" _
    & " Where CODIGOLIQUIDACION = '" & CodigoLiquidacion _
    & "' And PUESTOLABORAL = '" & PuestoLaboral & "'"
    
    dblImporteAcumulado = 0
    
    With LiquidacionGanancia4ta
        'Seguro de Vida Optativo del recibo requiere tratamiento especial, comenzamos cargando eso
        dblImporteCalculado = CalcularDescuentoSeguroDeVidaOptativo(PuestoLaboral, CodigoLiquidacion)
        .txtSeguroOptativo.Text = De_Num_a_Tx_01(dblImporteCalculado)
        'Abrimos el recordset que va a cargar el resto del formulario
        Set rstBuscarSlave = New ADODB.Recordset
        rstBuscarSlave.Open SQL, dbSlave, adOpenDynamic, adLockReadOnly
        'Renta Imponible
        .txtHaberOptimo.Text = De_Num_a_Tx_01(rstBuscarSlave!HaberOptimo)
        .txtPluriempleo.Text = De_Num_a_Tx_01(rstBuscarSlave!Pluriempleo)
        .txtAjuste.Text = De_Num_a_Tx_01(rstBuscarSlave!Ajuste)
        'Descuentos Recibo
        .txtJubilacion.Text = De_Num_a_Tx_01(rstBuscarSlave!Jubilacion)
        .txtObraSocial.Text = De_Num_a_Tx_01(rstBuscarSlave!ObraSocial)
        .txtAdherente.Text = De_Num_a_Tx_01(rstBuscarSlave!AdherenteObraSocial)
        .txtSeguroObligatorio.Text = De_Num_a_Tx_01(rstBuscarSlave!SeguroDeVidaObligatorio)
        .txtCuotaSindical.Text = De_Num_a_Tx_01(rstBuscarSlave!CuotaSindical)
        .txtOtrosDescuentos = De_Num_a_Tx_01(Round(0, 2)) 'No prevista en BD
        'Cargamos las Deducciones Generales
        ConfigurardgDeduccionesGeneralesLG4ta
        dblImporteCalculado = rstBuscarSlave!SeguroDeVidaOptativo
        dblImporteAcumulado = dblImporteCalculado
        Call CargardgDeduccionesGeneralesLG4taIndividual("Seguro de Vida", dblImporteCalculado)
        dblImporteCalculado = rstBuscarSlave!ServicioDomestico
        dblImporteAcumulado = dblImporteAcumulado + dblImporteCalculado
        Call CargardgDeduccionesGeneralesLG4taIndividual("Servicio Doméstido", dblImporteCalculado)
        dblImporteCalculado = rstBuscarSlave!Alquileres
        dblImporteAcumulado = dblImporteAcumulado + dblImporteCalculado
        Call CargardgDeduccionesGeneralesLG4taIndividual("Alquiler", dblImporteCalculado)
        dblImporteCalculado = rstBuscarSlave!CuotaSindical
        dblImporteAcumulado = dblImporteAcumulado + dblImporteCalculado
        Call CargardgDeduccionesGeneralesLG4taIndividual("Cuota Sindical", dblImporteCalculado)
        dblImporteCalculado = rstBuscarSlave!CuotaMedicoAsistencial
        dblImporteAcumulado = dblImporteAcumulado + dblImporteCalculado
        Call CargardgDeduccionesGeneralesLG4taIndividual("Cuota Médica Asist.", dblImporteCalculado)
        dblImporteCalculado = rstBuscarSlave!Donaciones
        dblImporteAcumulado = dblImporteAcumulado + dblImporteCalculado
        Call CargardgDeduccionesGeneralesLG4taIndividual("Donaciones", dblImporteCalculado)
        dblImporteCalculado = rstBuscarSlave!HonorariosMedicos
        If dblImporteCalculado <> 0 Then
            .chkLiquidacionFinal.Value = 1
        End If
        dblImporteAcumulado = dblImporteAcumulado + dblImporteCalculado
        Call CargardgDeduccionesGeneralesLG4taIndividual("Honorarios Médicos", dblImporteCalculado)
        Call CargardgDeduccionesGeneralesLG4taIndividual("Total Mensual", dblImporteAcumulado)
        'Cargamos las Deducciones Personales
        ConfigurardgDeduccionesPersonalesLG4ta
        dblImporteCalculado = rstBuscarSlave!MinimoNoImponible
        dblImporteAcumulado = dblImporteCalculado
        Call CargardgDeduccionesPersonalesLG4taIndividual("Mín. no Imponible", dblImporteCalculado)
        dblImporteCalculado = rstBuscarSlave!Conyuge
        dblImporteAcumulado = dblImporteAcumulado + dblImporteCalculado
        Call CargardgDeduccionesPersonalesLG4taIndividual("Conyuge", dblImporteCalculado)
        dblImporteCalculado = rstBuscarSlave!Hijo
        dblImporteAcumulado = dblImporteAcumulado + dblImporteCalculado
        Call CargardgDeduccionesPersonalesLG4taIndividual("Hijo/s", dblImporteCalculado)
        dblImporteCalculado = rstBuscarSlave!OtrasCargasDeFamilia
        dblImporteAcumulado = dblImporteAcumulado + dblImporteCalculado
        Call CargardgDeduccionesPersonalesLG4taIndividual("Otro/s Flia.", dblImporteCalculado)
        dblImporteCalculado = rstBuscarSlave!DeduccionEspecial
        dblImporteAcumulado = dblImporteAcumulado + dblImporteCalculado
        Call CargardgDeduccionesPersonalesLG4taIndividual("Deducción Especial", dblImporteCalculado)
        Call CargardgDeduccionesPersonalesLG4taIndividual("Total Mensual", dblImporteAcumulado)
        'Ajuste Retencion
        .txtAjuesteRetencion.Text = De_Num_a_Tx_01(rstBuscarSlave!AjusteRetencion)
        'Retencion Periodo
        .txtRetencionPeriodo.Text = De_Num_a_Tx_01(rstBuscarSlave!Retencion)
        'Cerramos del recordset porque no lo necesitamos más
        rstBuscarSlave.Close
        Set rstBuscarSlave = Nothing
        'Calculamos el total Acumulado de las deduciones
        dblImporteCalculado = CalcularDeduccionesGeneralesAcumuladas(PuestoLaboral, CodigoLiquidacion, True)
        Call CargardgDeduccionesGeneralesLG4taIndividual("Total Acumulado", dblImporteCalculado)
        dblImporteCalculado = CalcularDeduccionesPersonalesAcumuladas(PuestoLaboral, CodigoLiquidacion, True)
        Call CargardgDeduccionesPersonalesLG4taIndividual("Total Acumulado", dblImporteCalculado)
        'Procedemos a completar el resto de los datos del formulario
        Call CargarLiquidacionGanancia4ta("SubtotalSueldo", PuestoLaboral, CodigoLiquidacion)
        Call CargarLiquidacionGanancia4ta("SubtotalDescuentos", PuestoLaboral, CodigoLiquidacion)
        Call CargarLiquidacionGanancia4ta("RentayDescuentosAcumulados", PuestoLaboral, CodigoLiquidacion)
        Call CargarLiquidacionGanancia4ta("SueldoNeto", PuestoLaboral, CodigoLiquidacion)
        Call CargarLiquidacionGanancia4ta("GananciaNetaAcumulada", PuestoLaboral, CodigoLiquidacion)
        Call CargarLiquidacionGanancia4ta("RetencionTotal", PuestoLaboral, CodigoLiquidacion)
        Call CargarLiquidacionGanancia4ta("RetencionAcumulada", PuestoLaboral, CodigoLiquidacion)
    End With
    
    SQL = ""
    dblImporteCalculado = 0
    dblImporteAcumulado = 0
    
End Sub

Public Sub CargarLiquidacionGanancia4ta(PasoLiquidacion As String, PuestoLaboral As String, _
CodigoLiquidacion As String, Optional LiquidacionFinal As Boolean = False)

    Dim Ajustar As Integer
    Dim dblDiferenciaLiquidacion As Double
    Dim dblImporteCalculado As Double
    Dim dblImporteAcumulado As Double
    Dim dblRdoNetoAntesDeDeducciones As Double
    Dim strCLPrevio As String
    Dim i As Integer
    
    strCLPrevio = BuscarCodigoLiquidacionAnterior(CodigoLiquidacion)
    
    With LiquidacionGanancia4ta
        Select Case PasoLiquidacion
        Case "RentaImponible"
            '1)Renta Imponible
            dblImporteCalculado = CalcularHaberBruto(PuestoLaboral, CodigoLiquidacion)
            dblDiferenciaLiquidacion = DiferenciaHaberBrutoAcumuladoSISPERvsSLAVE(PuestoLaboral, CodigoLiquidacion, dblImporteCalculado)
            If dblDiferenciaLiquidacion <> 0 Then
                Ajustar = MsgBox("Se ha detectado una diferencia de $ " & dblDiferenciaLiquidacion & " en el importe acumulado de Haber Bruto entre SISPER y SLAVE" & vbCrLf & "Desea ajustar el Haber Óptimo del período en base a la diferencia detectada?", vbQuestion + vbYesNo, "DIFERENCIA ACUMULADA SISPER VS SLAVE")
                If Ajustar = 6 Then
                    dblImporteCalculado = dblImporteCalculado + dblDiferenciaLiquidacion
                Else
                    dblDiferenciaLiquidacion = 0
                End If
            End If
            dblImporteAcumulado = dblImporteCalculado
            .txtHaberOptimo.Text = De_Num_a_Tx_01(dblImporteCalculado)
            dblImporteCalculado = CalcularPluriempleo(PuestoLaboral, CodigoLiquidacion)
            dblImporteAcumulado = dblImporteAcumulado + dblImporteCalculado
            .txtPluriempleo.Text = De_Num_a_Tx_01(dblImporteCalculado)
            dblImporteCalculado = CalcularAjuste(PuestoLaboral, CodigoLiquidacion)
            dblImporteAcumulado = dblImporteAcumulado + dblImporteCalculado
            .txtAjuste.Text = De_Num_a_Tx_01(dblImporteCalculado)
            dblImporteAcumulado = Round(dblImporteAcumulado, 2)
            .txtSubtotalSueldo.Text = De_Num_a_Tx_01(dblImporteAcumulado)
        Case "DescuentosRecibo"
            '2)Descuentos Recibo INVICO
            If dblDiferenciaLiquidacion = 0 Then
                dblImporteCalculado = CalcularJubilacion(PuestoLaboral, CodigoLiquidacion)
                dblImporteAcumulado = dblImporteCalculado
                .txtJubilacion.Text = De_Num_a_Tx_01(dblImporteCalculado)
                dblImporteCalculado = CalcularObraSocial(PuestoLaboral, CodigoLiquidacion)
                dblImporteAcumulado = dblImporteAcumulado + dblImporteCalculado
                .txtObraSocial.Text = De_Num_a_Tx_01(dblImporteCalculado)
            Else
                'Recalculo Descuento Jubilación y O.Social en función del Haber Óptimo
                dblImporteCalculado = De_Txt_a_Num_01(.txtHaberOptimo.Text) * 0.185
                dblImporteAcumulado = dblImporteCalculado
                .txtJubilacion.Text = De_Num_a_Tx_01(Round(dblImporteCalculado, 2))
                dblImporteCalculado = De_Txt_a_Num_01(.txtHaberOptimo.Text) * 0.05
                dblImporteAcumulado = dblImporteAcumulado + dblImporteCalculado
                .txtObraSocial.Text = De_Num_a_Tx_01(Round(dblImporteCalculado, 2))
                dblDiferenciaLiquidacion = 0
            End If
            dblImporteCalculado = CalcularAdherenteObraSocial(PuestoLaboral, CodigoLiquidacion)
            dblImporteAcumulado = dblImporteAcumulado + dblImporteCalculado
            .txtAdherente.Text = De_Num_a_Tx_01(dblImporteCalculado)
            dblImporteCalculado = CalcularDescuentoSeguroDeVidaObligatorio(PuestoLaboral, CodigoLiquidacion)
            dblImporteAcumulado = dblImporteAcumulado + dblImporteCalculado
            .txtSeguroObligatorio.Text = De_Num_a_Tx_01(dblImporteCalculado)
            dblImporteCalculado = CalcularDescuentoCuotaSindical(PuestoLaboral, CodigoLiquidacion)
            dblDiferenciaLiquidacion = DiferenciaCuotaSindicalAcumuladaSISPERvsSLAVE(PuestoLaboral, CodigoLiquidacion, dblImporteCalculado)
            If dblDiferenciaLiquidacion <> 0 Then
                Ajustar = MsgBox("Se ha detectado una diferencia de $ " & dblDiferenciaLiquidacion & " en el importe acumulado de Cuota Sindical entre SISPER y SLAVE" & vbCrLf & "Desea ajustar la Cuota Sindical del período en base a la diferencia detectada?", vbQuestion + vbYesNo, "DIFERENCIA ACUMULADA SISPER VS SLAVE")
                If Ajustar = 6 Then
                    dblImporteCalculado = dblImporteCalculado + dblDiferenciaLiquidacion
                Else
                    dblDiferenciaLiquidacion = 0
                End If
            End If
            dblImporteAcumulado = dblImporteAcumulado + dblImporteCalculado
            .txtCuotaSindical.Text = De_Num_a_Tx_01(dblImporteCalculado)
            dblImporteCalculado = CalcularDescuentoSeguroDeVidaOptativo(PuestoLaboral, CodigoLiquidacion)
            dblImporteAcumulado = dblImporteAcumulado + dblImporteCalculado
            .txtSeguroOptativo.Text = De_Num_a_Tx_01(dblImporteCalculado)
            dblImporteCalculado = CalcularOtrosDescuentos(PuestoLaboral, CodigoLiquidacion)
            dblImporteAcumulado = dblImporteAcumulado + dblImporteCalculado
            .txtOtrosDescuentos.Text = De_Num_a_Tx_01(dblImporteCalculado)
            dblImporteAcumulado = Round(dblImporteAcumulado, 2)
            .txtSubtotalDescuento.Text = De_Num_a_Tx_01(dblImporteAcumulado)
        Case "RecalculoDescuentosObligatorios"
            'Recalculo Descuento Jubilación y O.Social en función del Haber Óptimo
            'Verificamos si los Aportes Personales guardan relación con el Haber Óptimo
            dblImporteCalculado = De_Txt_a_Num_01(.txtHaberOptimo.Text) * 0.185
            .txtJubilacion.Text = De_Num_a_Tx_01(Round(dblImporteCalculado, 2))
            dblImporteCalculado = De_Txt_a_Num_01(.txtHaberOptimo.Text) * 0.05
            .txtObraSocial.Text = De_Num_a_Tx_01(Round(dblImporteCalculado, 2))
        Case "SubtotalSueldo"
            'Permite calcular el Subtotal de Renta Imponible
            dblImporteCalculado = De_Txt_a_Num_01(.txtHaberOptimo.Text) + De_Txt_a_Num_01(.txtAjuste.Text) _
            + De_Txt_a_Num_01(.txtPluriempleo.Text)
            .txtSubtotalSueldo.Text = De_Num_a_Tx_01(Round(dblImporteCalculado, 2))
        Case "SubtotalDescuentos"
            'Permite calcular el Subtotal Descuentos de Recibo INVICO
            dblImporteCalculado = De_Txt_a_Num_01(.txtJubilacion.Text) + De_Txt_a_Num_01(.txtObraSocial.Text) _
            + De_Txt_a_Num_01(.txtAdherente.Text) + De_Txt_a_Num_01(.txtSeguroObligatorio.Text) _
            + De_Txt_a_Num_01(.txtCuotaSindical.Text) + De_Txt_a_Num_01(.txtSeguroOptativo.Text) _
            + De_Txt_a_Num_01(.txtOtrosDescuentos.Text)
            .txtSubtotalDescuento.Text = De_Num_a_Tx_01(Round(dblImporteCalculado, 2))
        Case "RentayDescuentosAcumulados"
            '3)Renta y Desc. Acum. a la liquidación anterior (todo aquello que no esta incluido en deducciones generales)
            dblImporteCalculado = CalcularRentaAcumulada(PuestoLaboral, CodigoLiquidacion, False)
            .txtRentaAcumulada.Text = De_Num_a_Tx_01(dblImporteCalculado)
            dblImporteCalculado = CalcularDescuentoAcumulado(PuestoLaboral, CodigoLiquidacion, False)
            .txtDescuentoAcumulado.Text = De_Num_a_Tx_01(dblImporteCalculado)
        Case "SueldoNeto"
            '4)Gcia. del periodo (todo aquello que no esta incluido en deducciones generales)
            'Sumamos el subtotal Sueldo
            dblImporteCalculado = De_Txt_a_Num_01(.txtSubtotalSueldo.Text)
            dblImporteAcumulado = dblImporteCalculado
            'Le restamos los descuentos obligatorios y adherente obra social
            dblImporteCalculado = De_Txt_a_Num_01(.txtJubilacion.Text) + De_Txt_a_Num_01(.txtObraSocial.Text) _
            + De_Txt_a_Num_01(.txtAdherente.Text) + De_Txt_a_Num_01(.txtSeguroObligatorio.Text)
            dblImporteAcumulado = dblImporteAcumulado - dblImporteCalculado
            dblImporteAcumulado = Round(dblImporteAcumulado, 2)
            .txtGananciaPeriodo.Text = De_Num_a_Tx_01(dblImporteAcumulado)
        Case "GananciaNetaAcumulada"
            '5) Calculamos Ganancia Neta Acumulada antes de Deducciones Personales y Generales
            'Partimos de la Renta acumulada hasta la liquidación anterior
            dblImporteCalculado = De_Txt_a_Num_01(.txtRentaAcumulada.Text)
            dblImporteAcumulado = dblImporteCalculado
            'Le restamos los descuentos acumulados hasta la liquidación previa (todos aquellos que no son deducciones generales)
            dblImporteCalculado = De_Txt_a_Num_01(.txtDescuentoAcumulado.Text)
            dblImporteAcumulado = dblImporteAcumulado - dblImporteCalculado
            'Le sumamos la ganancia del período
            dblImporteCalculado = De_Txt_a_Num_01(.txtGananciaPeriodo.Text)
            dblImporteAcumulado = dblImporteAcumulado + dblImporteCalculado
            'Convertimos a texto el total
            dblImporteAcumulado = Round(dblImporteAcumulado, 2)
            .txtGananciaNeta.Text = De_Num_a_Tx_01(dblImporteAcumulado)
        Case "DeduccionesGenerales"
            '6)Configuramos y Cargamos Deducciones Generales (Tabla NO EDITABLE)
            dblRdoNetoAntesDeDeducciones = De_Txt_a_Num_01(.txtGananciaNeta.Text)
            ConfigurardgDeduccionesGeneralesLG4ta
            dblImporteCalculado = Round(De_Txt_a_Num_01(.txtSeguroOptativo.Text), 2)
            dblImporteCalculado = CalcularDeduccionSeguroDeVida(PuestoLaboral, CodigoLiquidacion, dblImporteCalculado)
            dblImporteAcumulado = dblImporteCalculado
            Call CargardgDeduccionesGeneralesLG4taIndividual("Seguro de Vida", dblImporteCalculado)
            dblRdoNetoAntesDeDeducciones = dblRdoNetoAntesDeDeducciones - dblImporteCalculado
            dblImporteCalculado = CalcularDeduccionServicioDomestico(PuestoLaboral, CodigoLiquidacion)
            dblImporteAcumulado = dblImporteAcumulado + dblImporteCalculado
            Call CargardgDeduccionesGeneralesLG4taIndividual("Servicio Doméstido", dblImporteCalculado)
            dblRdoNetoAntesDeDeducciones = dblRdoNetoAntesDeDeducciones - dblImporteCalculado
            dblImporteCalculado = CalcularDeduccionAlquiler(PuestoLaboral, CodigoLiquidacion)
            dblImporteAcumulado = dblImporteAcumulado + dblImporteCalculado
            Call CargardgDeduccionesGeneralesLG4taIndividual("Alquiler", dblImporteCalculado)
            dblRdoNetoAntesDeDeducciones = dblRdoNetoAntesDeDeducciones - dblImporteCalculado
            dblImporteCalculado = Round(De_Txt_a_Num_01(.txtCuotaSindical.Text), 2)
            dblImporteCalculado = CalcularDeduccionCuotaSindical(PuestoLaboral, CodigoLiquidacion, dblImporteCalculado)
            dblImporteAcumulado = dblImporteAcumulado + dblImporteCalculado
            Call CargardgDeduccionesGeneralesLG4taIndividual("Cuota Sindical", dblImporteCalculado)
            dblRdoNetoAntesDeDeducciones = dblRdoNetoAntesDeDeducciones - dblImporteCalculado
            dblImporteCalculado = CalcularDeduccionCuotaMedicoAsistencial(PuestoLaboral, CodigoLiquidacion, , dblRdoNetoAntesDeDeducciones)
            dblImporteAcumulado = dblImporteAcumulado + dblImporteCalculado
            Call CargardgDeduccionesGeneralesLG4taIndividual("Cuota Médica Asist.", dblImporteCalculado)
            dblImporteCalculado = CalcularDeduccionDonaciones(PuestoLaboral, CodigoLiquidacion, , dblRdoNetoAntesDeDeducciones)
            dblImporteAcumulado = dblImporteAcumulado + dblImporteCalculado
            Call CargardgDeduccionesGeneralesLG4taIndividual("Donaciones", dblImporteCalculado)
            If LiquidacionFinal = False Then
                dblImporteCalculado = 0
            Else
                dblImporteCalculado = CalcularDeduccionHonorariosMedicos(PuestoLaboral, CodigoLiquidacion, , dblRdoNetoAntesDeDeducciones)
            End If
            dblImporteAcumulado = dblImporteAcumulado + dblImporteCalculado
            Call CargardgDeduccionesGeneralesLG4taIndividual("Honorarios Medicos", dblImporteCalculado)
            Call CargardgDeduccionesGeneralesLG4taIndividual("Total Mensual", dblImporteAcumulado)
            'Buscamos el total de deducciones acumuladas hasta la liquidacion anterior y le sumamos al total mensual
            dblImporteCalculado = CalcularDeduccionesGeneralesAcumuladas(PuestoLaboral, CodigoLiquidacion, False)
            dblImporteAcumulado = dblImporteAcumulado + dblImporteCalculado
            Call CargardgDeduccionesGeneralesLG4taIndividual("Total Acumulado", dblImporteAcumulado)
        Case "DeduccionesPersonales"
            '7)Configuramos y Cargamos Deducciones Personales (Tabla NO EDITABLE)
            dblRdoNetoAntesDeDeducciones = De_Txt_a_Num_01(.txtGananciaNeta.Text) _
            - De_Txt_a_Num_01(.dgDeduccionesGenerales.TextMatrix(.dgDeduccionesGenerales.Rows - 1, 1))
            ConfigurardgDeduccionesPersonalesLG4ta
            dblImporteCalculado = CalcularDeduccionMinimoNoImponible(PuestoLaboral, CodigoLiquidacion)
            dblRdoNetoAntesDeDeducciones = dblRdoNetoAntesDeDeducciones - dblImporteCalculado
            dblImporteAcumulado = dblImporteCalculado
            Call CargardgDeduccionesPersonalesLG4taIndividual("Mín. no Imponible", dblImporteCalculado)
            dblImporteCalculado = CalcularDeduccionConyuge(PuestoLaboral, CodigoLiquidacion)
            dblRdoNetoAntesDeDeducciones = dblRdoNetoAntesDeDeducciones - dblImporteCalculado
            dblImporteAcumulado = dblImporteAcumulado + dblImporteCalculado
            Call CargardgDeduccionesPersonalesLG4taIndividual("Conyuge", dblImporteCalculado)
            dblImporteCalculado = CalcularDeduccionHijo(PuestoLaboral, CodigoLiquidacion)
            dblRdoNetoAntesDeDeducciones = dblRdoNetoAntesDeDeducciones - dblImporteCalculado
            dblImporteAcumulado = dblImporteAcumulado + dblImporteCalculado
            Call CargardgDeduccionesPersonalesLG4taIndividual("Hijo/s", dblImporteCalculado)
            dblImporteCalculado = CalcularDeduccionOtrasCargasDeFamilia(PuestoLaboral, CodigoLiquidacion)
            dblRdoNetoAntesDeDeducciones = dblRdoNetoAntesDeDeducciones - dblImporteCalculado _
            - CalcularDeduccionesPersonalesAcumuladas(PuestoLaboral, CodigoLiquidacion, False, True)
            dblImporteAcumulado = dblImporteAcumulado + dblImporteCalculado
            Call CargardgDeduccionesPersonalesLG4taIndividual("Otro/s Flia.", dblImporteCalculado)
            dblImporteCalculado = CalcularDeduccionEspecial(PuestoLaboral, CodigoLiquidacion, , dblRdoNetoAntesDeDeducciones)
            dblImporteAcumulado = dblImporteAcumulado + dblImporteCalculado
            Call CargardgDeduccionesPersonalesLG4taIndividual("Deducción Especial", dblImporteCalculado)
            Call CargardgDeduccionesPersonalesLG4taIndividual("Total Mensual", dblImporteAcumulado)
            'Buscamos el total de deducciones acumuladas hasta la liquidacion anterior y le sumamos al total mensual
            dblImporteCalculado = CalcularDeduccionesPersonalesAcumuladas(PuestoLaboral, CodigoLiquidacion, False)
            dblImporteAcumulado = dblImporteAcumulado + dblImporteCalculado
            Call CargardgDeduccionesPersonalesLG4taIndividual("Total Acumulado", dblImporteAcumulado)
        Case "RetencionTotal"
            '8)Retención total que debería practicarse
            '8.1)Base Imponible
            '8.1.a)Ganancia Neta
            dblImporteCalculado = De_Txt_a_Num_01(.txtGananciaNeta.Text)
            dblImporteAcumulado = dblImporteCalculado
            '8.1.b)Deducciones Personales Acumuladas
            i = .dgDeduccionesPersonales.Rows - 1
            dblImporteCalculado = De_Txt_a_Num_01(.dgDeduccionesPersonales.TextMatrix(i, 1))
            dblImporteAcumulado = dblImporteAcumulado - dblImporteCalculado
            '8.1.c)Deducciones Generales Acumuladas
            i = .dgDeduccionesGenerales.Rows - 1
            dblImporteCalculado = De_Txt_a_Num_01(.dgDeduccionesGenerales.TextMatrix(i, 1))
            dblImporteAcumulado = dblImporteAcumulado - dblImporteCalculado
            dblImporteAcumulado = Round(dblImporteAcumulado, 2)
            .txtBaseImponible.Text = De_Num_a_Tx_01(dblImporteAcumulado)
            If dblImporteAcumulado > 0 Then
                '8.2)Porcentaje Aplicable (dblImporteAcumulado = Base Imponible)
                dblImporteCalculado = CalcularAlicuotaAplicable(PuestoLaboral, CodigoLiquidacion, dblImporteAcumulado)
                dblImporteCalculado = dblImporteCalculado
                .txtPorcentajeAplicable.Text = De_Num_a_Tx_01(dblImporteCalculado * 100, True) & " %"
                '8.3)Importe Variable (dblImporteAcumulado = Base Imponible y dblImporteCalculado = Alicuota)
                dblImporteCalculado = CalcularImporteVariable(PuestoLaboral, CodigoLiquidacion, dblImporteCalculado, dblImporteAcumulado)
                .txtSumaVariable.Text = De_Num_a_Tx_01(dblImporteCalculado)
                '8.4)Importe Fijo (dblImporteAcumulado = Base Imponible)
                dblImporteCalculado = CalcularImporteFijo(PuestoLaboral, CodigoLiquidacion, dblImporteAcumulado)
                .txtSumaFija.Text = De_Num_a_Tx_01(dblImporteCalculado)
                '8.5)Subtotal Retención (Retención que debería practicarse)
                dblImporteCalculado = De_Txt_a_Num_01(.txtSumaVariable.Text) + De_Txt_a_Num_01(.txtSumaFija.Text)
                .txtSubtotalRentencion.Text = De_Num_a_Tx_01(dblImporteCalculado)
            Else
                .txtPorcentajeAplicable.Text = De_Num_a_Tx_01(0)
                .txtSumaVariable.Text = De_Num_a_Tx_01(0)
                .txtSumaFija.Text = De_Num_a_Tx_01(0)
                .txtSubtotalRentencion.Text = De_Num_a_Tx_01(0)
            End If
        Case "RetencionAcumulada"
            '9)Retención Acumulada
            dblImporteCalculado = CalcularRetencionAcumulada(PuestoLaboral, CodigoLiquidacion, False)
            .txtRetencionAcumulada.Text = De_Num_a_Tx_01(dblImporteCalculado)
        Case "AjusteRetencion"
            '10)Ajuste Retención Ganancias
            dblImporteCalculado = CalcularAjusteRetencion(PuestoLaboral, CodigoLiquidacion)
            .txtAjuesteRetencion.Text = De_Num_a_Tx_01(dblImporteCalculado)
        Case "RetencionPeriodo"
            '11)Retención del Período
            dblImporteCalculado = De_Txt_a_Num_01(.txtSubtotalRentencion.Text) - _
            De_Txt_a_Num_01(.txtRetencionAcumulada.Text) + De_Txt_a_Num_01(.txtAjuesteRetencion.Text)
            .txtRetencionPeriodo.Text = De_Num_a_Tx_01(dblImporteCalculado)
        End Select
    End With

    dblImporteCalculado = 0
    dblImporteAcumulado = 0
    dblRdoNetoAntesDeDeducciones = 0
    strCLPrevio = ""
    i = 0

End Sub

Public Sub EditandoLiquidacionGanancia4taSIRADIG(PuestoLaboral As String, CodigoLiquidacion As String)

    Dim SQL As String
    Dim strPeriodo As String
    Dim dblImporteCalculado As Double
    Dim dblImporteMensualAcumulado As Double
    Dim strCodigoSIRADIG As String
    Dim strConceptoSIRADIG As String
    Dim strDenominacion As String
    Dim dblImporteAcumulado As Double
    
    strPeriodo = BuscarPeriodoLiquidacion(CodigoLiquidacion)
    
    'Creamos la consulta base que va a cargar el formulario
    SQL = "Select * From LIQUIDACIONGANANCIAS4TACATEGORIA" _
    & " Where CODIGOLIQUIDACION = '" & CodigoLiquidacion _
    & "' And PUESTOLABORAL = '" & PuestoLaboral & "'"
    
    dblImporteAcumulado = 0
    
    With LiquidacionGanancia4taSIRADIG
        'Seguro de Vida Optativo del recibo requiere tratamiento especial, comenzamos cargando eso
        dblImporteCalculado = CalcularDescuentoSeguroDeVidaOptativo(PuestoLaboral, CodigoLiquidacion)
        .txtSeguroOptativo.Text = De_Num_a_Tx_01(dblImporteCalculado)
        'Abrimos el recordset que va a cargar el resto del formulario
        Set rstBuscarSlave = New ADODB.Recordset
        rstBuscarSlave.Open SQL, dbSlave, adOpenDynamic, adLockReadOnly
        'Renta Imponible
        .txtHaberOptimo.Text = De_Num_a_Tx_01(rstBuscarSlave!HaberOptimo)
        .txtPluriempleo.Text = De_Num_a_Tx_01(rstBuscarSlave!Pluriempleo)
        .txtAjuste.Text = De_Num_a_Tx_01(rstBuscarSlave!Ajuste)
        'Descuentos Recibo
        .txtJubilacion.Text = De_Num_a_Tx_01(rstBuscarSlave!Jubilacion)
        .txtObraSocial.Text = De_Num_a_Tx_01(rstBuscarSlave!ObraSocial)
        .txtAdherente.Text = De_Num_a_Tx_01(rstBuscarSlave!AdherenteObraSocial)
        .txtSeguroObligatorio.Text = De_Num_a_Tx_01(rstBuscarSlave!SeguroDeVidaObligatorio)
        .txtCuotaSindical.Text = De_Num_a_Tx_01(rstBuscarSlave!CuotaSindical)
        .txtOtrosDescuentos = De_Num_a_Tx_01(Round(0, 2)) 'No prevista en BD
        'Ajuste Retencion
        .txtAjuesteRetencion.Text = De_Num_a_Tx_01(rstBuscarSlave!AjusteRetencion)
        'Retencion Periodo
        .txtRetencionPeriodo.Text = De_Num_a_Tx_01(rstBuscarSlave!Retencion)
        'Cerramos del recordset porque no lo necesitamos más
        rstBuscarSlave.Close
        Set rstBuscarSlave = Nothing
        
        'Cargamos la Ganancia Bruta Acumulada
        .txtGananciaBrutaAcumulada.Text = De_Num_a_Tx_01(CalcularRentaAcumulada(PuestoLaboral, _
        CodigoLiquidacion, True))
        
        'Cargamos las Deducciones Generales
        Dim strDeducciones() As String
        ReDim strDeducciones(7)
        strDeducciones(0) = "Jubilacion"
        strDeducciones(1) = "ObraSocial"
        strDeducciones(2) = "AdherenteObraSocial"
        strDeducciones(3) = "SeguroDeVidaObligatorio"
        strDeducciones(4) = "SeguroDeVidaOptativo"
        strDeducciones(5) = "CuotaSindical"
        strDeducciones(6) = "ServicioDomestico"
        strDeducciones(7) = "Alquileres"
        
        ConfigurardgDeduccionesGeneralesLG4taSIRADIG
        For d = 0 To UBound(strDeducciones)
            strCodigoSIRADIG = EquipararConceptoDeduccionSISPERconCodigoSIRADIG(strDeducciones(d))
            strConceptoSIRADIG = BuscarDenominacionDeduccionSIRADIG(strCodigoSIRADIG)
            dblImporteCalculado = ImporteRegistradoDeduccionEspecifica(strDeducciones(d), _
            PuestoLaboral, CodigoLiquidacion)
            dblImporteMensualAcumulado = dblImporteMensualAcumulado + dblImporteCalculado
            dblImporteAcumulado = ImporteRegistradoAcumuladoDeduccionEspecifica(strDeducciones(d), _
            PuestoLaboral, strPeriodo, CodigoLiquidacion)
            If dblImporteCalculado <> 0 Or dblImporteAcumulado <> 0 Then
                Call CargardgDeduccionesGeneralesLG4taSIRADIG(strCodigoSIRADIG, strConceptoSIRADIG, _
                dblImporteCalculado, dblImporteAcumulado)
            End If
        Next
        dblImporteAcumulado = CalcularDeduccionesGeneralesAcumuladas(PuestoLaboral, CodigoLiquidacion, True, True, True)
        Call CargardgDeduccionesGeneralesLG4taSIRADIG("ST", "-SUBTOTAL-", dblImporteMensualAcumulado, dblImporteAcumulado)

        ReDim strDeducciones(2)
        strDeducciones(0) = "CuotaMedicoAsistencial"
        strDeducciones(1) = "Donaciones"
        strDeducciones(2) = "HonorariosMedicos"
        For d = 0 To UBound(strDeducciones)
            strCodigoSIRADIG = EquipararConceptoDeduccionSISPERconCodigoSIRADIG(strDeducciones(d))
            strConceptoSIRADIG = BuscarDenominacionDeduccionSIRADIG(strCodigoSIRADIG)
            dblImporteCalculado = ImporteRegistradoDeduccionEspecifica(strDeducciones(d), _
            PuestoLaboral, CodigoLiquidacion)
            dblImporteMensualAcumulado = dblImporteMensualAcumulado + dblImporteCalculado
            dblImporteAcumulado = ImporteRegistradoAcumuladoDeduccionEspecifica(strDeducciones(d), _
            PuestoLaboral, strPeriodo, CodigoLiquidacion)
            If dblImporteCalculado <> 0 Or dblImporteAcumulado <> 0 Then
                Call CargardgDeduccionesGeneralesLG4taSIRADIG(strCodigoSIRADIG, strConceptoSIRADIG, _
                dblImporteCalculado, dblImporteAcumulado)
            End If
            If strDeducciones(d) = "HonorariosMedicos" And dblImporteCalculado <> 0 Then
                .chkLiquidacionFinal.Value = 1
            End If
        Next
        dblImporteAcumulado = CalcularDeduccionesGeneralesAcumuladas(PuestoLaboral, CodigoLiquidacion, True, False, True)
        Call CargardgDeduccionesGeneralesLG4taSIRADIG("T", "-TOTAL-", dblImporteMensualAcumulado, dblImporteAcumulado)
        .txtDeduccionesGeneralesAcumuladas.Text = De_Num_a_Tx_01(dblImporteAcumulado)

        'Cargamos las Deducciones Personales
        dblImporteMensualAcumulado = 0
        ReDim strDeducciones(3)
        strDeducciones(0) = "MNI"
        strDeducciones(1) = "C"
        strDeducciones(2) = "H"
        strDeducciones(3) = "OCF"

        ConfigurardgDeduccionesPersonalesLG4taSIRADIG
        For d = 0 To UBound(strDeducciones)
'            strCodigoSIRADIG = EquipararConceptoDeduccionSISPERconCodigoSIRADIG(strDeducciones(d))
            strDenominacion = BuscarDenominacionDeduccionPersonalSLAVE(strDeducciones(d))
            dblImporteCalculado = ImporteRegistradoDeduccionEspecifica(strDenominacion, _
            PuestoLaboral, CodigoLiquidacion)
            dblImporteMensualAcumulado = dblImporteMensualAcumulado + dblImporteCalculado
            dblImporteAcumulado = ImporteRegistradoAcumuladoDeduccionEspecifica(strDenominacion, _
            PuestoLaboral, strPeriodo, CodigoLiquidacion)
'            If dblImporteCalculado <> 0 Or dblImporteAcumulado <> 0 Then
                Call CargardgDeduccionesPersonalesLG4taSIRADIG(strDeducciones(d), strDenominacion, _
                dblImporteCalculado, dblImporteAcumulado)
'            End If
        Next
        dblImporteAcumulado = CalcularDeduccionesPersonalesAcumuladas(PuestoLaboral, CodigoLiquidacion, True, True)
        Call CargardgDeduccionesPersonalesLG4taSIRADIG("ST", "-SUBTOTAL-", dblImporteMensualAcumulado, dblImporteAcumulado)
        .txtDeduccionesPersonalesAcumuladas.Text = De_Num_a_Tx_01(dblImporteAcumulado)
        'Agregamos por separado la Deducción Especial
        strDenominacion = BuscarDenominacionDeduccionPersonalSLAVE("DE")
        dblImporteCalculado = ImporteRegistradoDeduccionEspecifica(strDenominacion, _
        PuestoLaboral, CodigoLiquidacion)
        dblImporteMensualAcumulado = dblImporteMensualAcumulado + dblImporteCalculado
        dblImporteAcumulado = ImporteRegistradoAcumuladoDeduccionEspecifica(strDenominacion, _
        PuestoLaboral, strPeriodo, CodigoLiquidacion)
        Call CargardgDeduccionesPersonalesLG4taSIRADIG("DE", strDenominacion, _
        dblImporteCalculado, dblImporteAcumulado)
        dblImporteAcumulado = CalcularDeduccionesPersonalesAcumuladas(PuestoLaboral, CodigoLiquidacion, True, False)
        Call CargardgDeduccionesPersonalesLG4taSIRADIG("T", "-TOTAL-", dblImporteMensualAcumulado, dblImporteAcumulado)
        .txtDeduccionesPersonalesAcumuladas.Text = De_Num_a_Tx_01(dblImporteAcumulado)

        'Procedemos a completar el resto de los datos del formulario
        Call CargarLiquidacionGanancia4taSIRADIG("SubtotalSueldo", PuestoLaboral, CodigoLiquidacion)
        Call CargarLiquidacionGanancia4taSIRADIG("SubtotalDescuentos", PuestoLaboral, CodigoLiquidacion)
        Call CargarLiquidacionGanancia4taSIRADIG("RetencionTotal", PuestoLaboral, CodigoLiquidacion)
        Call CargarLiquidacionGanancia4taSIRADIG("RetencionAcumulada", PuestoLaboral, CodigoLiquidacion)
    End With
    
    SQL = ""
    dblImporteCalculado = 0
    dblImporteAcumulado = 0
    
End Sub

Public Sub CargarLiquidacionGanancia4taSIRADIG(PasoLiquidacion As String, PuestoLaboral As String, _
CodigoLiquidacion As String, Optional LiquidacionFinal As Boolean = False)

    Dim Ajustar As Integer
    Dim dblDiferenciaLiquidacion As Double
    Dim dblImporteCalculado As Double
    Dim dblImporteAcumulado As Double
    Dim dblImporteMensualAcumulado As Double
    Dim dblRdoNetoAntesDeDeducciones As Double
    Dim strCLPrevio As String
    Dim strPeriodo As String
    Dim i As Integer
    Dim strDeducciones() As String
    Dim strCodigoSIRADIG As String
    Dim strConceptoSIRADIG As String
    Dim strIDLiquidacion As String
    Dim strDenominacion As String
    
    
    strCLPrevio = BuscarCodigoLiquidacionAnterior(CodigoLiquidacion)
    strPeriodo = BuscarPeriodoLiquidacion(CodigoLiquidacion)
    
    With LiquidacionGanancia4taSIRADIG
        
        If .cmbPresentacionesSIRADIG.Text = "Ninguno" Then
            strIDLiquidacion = .cmbPresentacionesSIRADIG.Text
        Else
            strIDLiquidacion = Right(.cmbPresentacionesSIRADIG.Text, 9)
            strIDLiquidacion = Left(strIDLiquidacion, 8)
        End If
        
        Select Case PasoLiquidacion
        Case "RentaImponible"
            '1)Renta Imponible
            dblImporteCalculado = CalcularHaberBruto(PuestoLaboral, CodigoLiquidacion)
            dblDiferenciaLiquidacion = DiferenciaHaberBrutoAcumuladoSISPERvsSLAVE(PuestoLaboral, CodigoLiquidacion, dblImporteCalculado)
            If dblDiferenciaLiquidacion <> 0 Then
                Ajustar = MsgBox("Se ha detectado una diferencia de $ " & dblDiferenciaLiquidacion & " en el importe acumulado de Haber Bruto entre SISPER y SLAVE" & vbCrLf & "Desea ajustar el Haber Óptimo del período en base a la diferencia detectada?", vbQuestion + vbYesNo, "DIFERENCIA ACUMULADA SISPER VS SLAVE")
                If Ajustar = 6 Then
                    dblImporteCalculado = dblImporteCalculado + dblDiferenciaLiquidacion
                Else
                    dblDiferenciaLiquidacion = 0
                End If
            End If
            dblImporteAcumulado = dblImporteCalculado
            .txtHaberOptimo.Text = De_Num_a_Tx_01(dblImporteCalculado)
            dblImporteCalculado = CalcularPluriempleo(PuestoLaboral, CodigoLiquidacion)
            dblImporteAcumulado = dblImporteAcumulado + dblImporteCalculado
            .txtPluriempleo.Text = De_Num_a_Tx_01(dblImporteCalculado)
            dblImporteCalculado = CalcularAjuste(PuestoLaboral, CodigoLiquidacion)
            dblImporteAcumulado = dblImporteAcumulado + dblImporteCalculado
            .txtAjuste.Text = De_Num_a_Tx_01(dblImporteCalculado)
            dblImporteAcumulado = Round(dblImporteAcumulado, 2)
            .txtSubtotalSueldo.Text = De_Num_a_Tx_01(dblImporteAcumulado)
            If dblDiferenciaLiquidacion <> 0 Then
                dblDiferenciaLiquidacionPublica = dblDiferenciaLiquidacion
            End If
        Case "DescuentosRecibo"
            '2)Descuentos Recibo INVICO
            If dblDiferenciaLiquidacionPublica = 0 Then
                dblImporteCalculado = CalcularJubilacion(PuestoLaboral, CodigoLiquidacion)
                dblImporteAcumulado = dblImporteCalculado
                .txtJubilacion.Text = De_Num_a_Tx_01(dblImporteCalculado)
                dblImporteCalculado = CalcularObraSocial(PuestoLaboral, CodigoLiquidacion)
                dblImporteAcumulado = dblImporteAcumulado + dblImporteCalculado
                .txtObraSocial.Text = De_Num_a_Tx_01(dblImporteCalculado)
            Else
                'Recalculo Descuento Jubilación y O.Social en función del Haber Óptimo
                dblImporteCalculado = De_Txt_a_Num_01(.txtHaberOptimo.Text) * 0.185
                dblImporteAcumulado = dblImporteCalculado
                .txtJubilacion.Text = De_Num_a_Tx_01(Round(dblImporteCalculado, 2))
                dblImporteCalculado = De_Txt_a_Num_01(.txtHaberOptimo.Text) * 0.05
                dblImporteAcumulado = dblImporteAcumulado + dblImporteCalculado
                .txtObraSocial.Text = De_Num_a_Tx_01(Round(dblImporteCalculado, 2))
                dblDiferenciaLiquidacionPublica = 0
            End If
            dblImporteCalculado = CalcularAdherenteObraSocial(PuestoLaboral, CodigoLiquidacion)
            dblImporteAcumulado = dblImporteAcumulado + dblImporteCalculado
            .txtAdherente.Text = De_Num_a_Tx_01(dblImporteCalculado)
            dblImporteCalculado = CalcularDescuentoSeguroDeVidaObligatorio(PuestoLaboral, CodigoLiquidacion)
            dblImporteAcumulado = dblImporteAcumulado + dblImporteCalculado
            .txtSeguroObligatorio.Text = De_Num_a_Tx_01(dblImporteCalculado)
            dblImporteCalculado = CalcularDescuentoCuotaSindical(PuestoLaboral, CodigoLiquidacion)
            dblDiferenciaLiquidacion = DiferenciaCuotaSindicalAcumuladaSISPERvsSLAVE(PuestoLaboral, CodigoLiquidacion, dblImporteCalculado)
            If dblDiferenciaLiquidacion <> 0 Then
                Ajustar = MsgBox("Se ha detectado una diferencia de $ " & dblDiferenciaLiquidacion & " en el importe acumulado de Cuota Sindical entre SISPER y SLAVE" & vbCrLf & "Desea ajustar la Cuota Sindical del período en base a la diferencia detectada?", vbQuestion + vbYesNo, "DIFERENCIA ACUMULADA SISPER VS SLAVE")
                If Ajustar = 6 Then
                    dblImporteCalculado = dblImporteCalculado + dblDiferenciaLiquidacion
                Else
                    dblDiferenciaLiquidacion = 0
                End If
            End If
            dblImporteAcumulado = dblImporteAcumulado + dblImporteCalculado
            .txtCuotaSindical.Text = De_Num_a_Tx_01(dblImporteCalculado)
            dblImporteCalculado = CalcularDescuentoSeguroDeVidaOptativo(PuestoLaboral, CodigoLiquidacion)
            dblImporteAcumulado = dblImporteAcumulado + dblImporteCalculado
            .txtSeguroOptativo.Text = De_Num_a_Tx_01(dblImporteCalculado)
            dblImporteCalculado = CalcularOtrosDescuentos(PuestoLaboral, CodigoLiquidacion)
            dblImporteAcumulado = dblImporteAcumulado + dblImporteCalculado
            .txtOtrosDescuentos.Text = De_Num_a_Tx_01(dblImporteCalculado)
            dblImporteAcumulado = Round(dblImporteAcumulado, 2)
            .txtSubtotalDescuento.Text = De_Num_a_Tx_01(dblImporteAcumulado)
        Case "RecalculoDescuentosObligatorios"
            'Recalculo Descuento Jubilación y O.Social en función del Haber Óptimo
            'Verificamos si los Aportes Personales guardan relación con el Haber Óptimo
            dblImporteCalculado = De_Txt_a_Num_01(.txtHaberOptimo.Text) * 0.185
            .txtJubilacion.Text = De_Num_a_Tx_01(Round(dblImporteCalculado, 2))
            dblImporteCalculado = De_Txt_a_Num_01(.txtHaberOptimo.Text) * 0.05
            .txtObraSocial.Text = De_Num_a_Tx_01(Round(dblImporteCalculado, 2))
        Case "SubtotalSueldo"
            'Permite calcular el Subtotal de Renta Imponible
            dblImporteCalculado = De_Txt_a_Num_01(.txtHaberOptimo.Text) + De_Txt_a_Num_01(.txtAjuste.Text) _
            + De_Txt_a_Num_01(.txtPluriempleo.Text)
            .txtSubtotalSueldo.Text = De_Num_a_Tx_01(Round(dblImporteCalculado, 2))
        Case "SubtotalDescuentos"
            'Permite calcular el Subtotal Descuentos de Recibo INVICO
            dblImporteCalculado = De_Txt_a_Num_01(.txtJubilacion.Text) + De_Txt_a_Num_01(.txtObraSocial.Text) _
            + De_Txt_a_Num_01(.txtAdherente.Text) + De_Txt_a_Num_01(.txtSeguroObligatorio.Text) _
            + De_Txt_a_Num_01(.txtCuotaSindical.Text) + De_Txt_a_Num_01(.txtSeguroOptativo.Text) _
            + De_Txt_a_Num_01(.txtOtrosDescuentos.Text)
            .txtSubtotalDescuento.Text = De_Num_a_Tx_01(Round(dblImporteCalculado, 2))
        Case "GananciaBrutaAcumulada"
            '3)Gcia. Bruta Acumulada
            dblImporteCalculado = CalcularRentaAcumulada(PuestoLaboral, CodigoLiquidacion, False)
            dblImporteCalculado = dblImporteCalculado + De_Txt_a_Num_01(.txtSubtotalSueldo.Text)
            .txtGananciaBrutaAcumulada.Text = De_Num_a_Tx_01(dblImporteCalculado)
        Case "DeduccionesGenerales"
            '6)Configuramos y Cargamos Deducciones Generales (Tabla NO EDITABLE)
            ConfigurardgDeduccionesGeneralesLG4taSIRADIG
            ReDim strDeducciones(1)
            'Agregamos el concepto Jubilación
            strDeducciones(0) = "Jubilacion"
            strCodigoSIRADIG = EquipararConceptoDeduccionSISPERconCodigoSIRADIG(strDeducciones(0))
            strConceptoSIRADIG = BuscarDenominacionDeduccionSIRADIG(strCodigoSIRADIG)
            dblImporteCalculado = De_Txt_a_Num_01(.txtJubilacion.Text)
            dblImporteMensualAcumulado = dblImporteMensualAcumulado + dblImporteCalculado
            dblImporteAcumulado = ImporteRegistradoAcumuladoDeduccionEspecifica(strDeducciones(0), _
            PuestoLaboral, strPeriodo, CodigoLiquidacion, False) + dblImporteCalculado
            If dblImporteCalculado <> 0 Or dblImporteAcumulado <> 0 Then
                Call CargardgDeduccionesGeneralesLG4taSIRADIG(strCodigoSIRADIG, strConceptoSIRADIG, _
                dblImporteCalculado, dblImporteAcumulado)
            End If
            'Agregamos el concepto Obra Social
            strDeducciones(0) = "ObraSocial"
            strCodigoSIRADIG = EquipararConceptoDeduccionSISPERconCodigoSIRADIG(strDeducciones(0))
            strConceptoSIRADIG = BuscarDenominacionDeduccionSIRADIG(strCodigoSIRADIG)
            dblImporteCalculado = De_Txt_a_Num_01(.txtObraSocial.Text)
            dblImporteMensualAcumulado = dblImporteMensualAcumulado + dblImporteCalculado
            dblImporteAcumulado = ImporteRegistradoAcumuladoDeduccionEspecifica(strDeducciones(0), _
            PuestoLaboral, strPeriodo, CodigoLiquidacion, False) + dblImporteCalculado
            If dblImporteCalculado <> 0 Or dblImporteAcumulado <> 0 Then
                Call CargardgDeduccionesGeneralesLG4taSIRADIG(strCodigoSIRADIG, strConceptoSIRADIG, _
                dblImporteCalculado, dblImporteAcumulado)
            End If
            'Agregamos el concepto Adherente Obra Social
            strDeducciones(0) = "AdherenteObraSocial"
            strCodigoSIRADIG = EquipararConceptoDeduccionSISPERconCodigoSIRADIG(strDeducciones(0))
            strConceptoSIRADIG = BuscarDenominacionDeduccionSIRADIG(strCodigoSIRADIG)
            dblImporteCalculado = De_Txt_a_Num_01(.txtAdherente.Text)
            dblImporteMensualAcumulado = dblImporteMensualAcumulado + dblImporteCalculado
            dblImporteAcumulado = ImporteRegistradoAcumuladoDeduccionEspecifica(strDeducciones(0), _
            PuestoLaboral, strPeriodo, CodigoLiquidacion, False) + dblImporteCalculado
            If dblImporteCalculado <> 0 Or dblImporteAcumulado <> 0 Then
                Call CargardgDeduccionesGeneralesLG4taSIRADIG(strCodigoSIRADIG, strConceptoSIRADIG, _
                dblImporteCalculado, dblImporteAcumulado)
            End If
            'Agregamos el concepto Seguro De Vida Obligatorio
            strDeducciones(0) = "SeguroDeVidaObligatorio"
            strCodigoSIRADIG = EquipararConceptoDeduccionSISPERconCodigoSIRADIG(strDeducciones(0))
            strConceptoSIRADIG = BuscarDenominacionDeduccionSIRADIG(strCodigoSIRADIG)
            dblImporteCalculado = CalcularDescuentoSeguroDeVidaObligatorioAcumulado(PuestoLaboral, _
            CodigoLiquidacion, False) - ImporteRegistradoAcumuladoDeduccionEspecifica(strDeducciones(0), _
            PuestoLaboral, strPeriodo, CodigoLiquidacion, False) + De_Txt_a_Num_01(.txtSeguroObligatorio.Text)
            dblImporteMensualAcumulado = dblImporteMensualAcumulado + dblImporteCalculado
            dblImporteAcumulado = ImporteRegistradoAcumuladoDeduccionEspecifica(strDeducciones(0), _
            PuestoLaboral, strPeriodo, CodigoLiquidacion, False) + dblImporteCalculado
            If dblImporteCalculado <> 0 Or dblImporteAcumulado <> 0 Then
                Call CargardgDeduccionesGeneralesLG4taSIRADIG(strCodigoSIRADIG, strConceptoSIRADIG, _
                dblImporteCalculado, dblImporteAcumulado)
            End If
            'Agregamos el concepto Seguro de Vida Optativo (NO ME CONVENCE)
            strDeducciones(0) = "SeguroDeVidaOptativo"
            strCodigoSIRADIG = EquipararConceptoDeduccionSISPERconCodigoSIRADIG(strDeducciones(0))
            strConceptoSIRADIG = BuscarDenominacionDeduccionSIRADIG(strCodigoSIRADIG)
            'Verificamos si tiene Seguro Optativo en Recibo de Sueldo
            dblImporteCalculado = CalcularDescuentoSeguroDeVidaOptativoAcumulado(PuestoLaboral, _
            CodigoLiquidacion, LiquidacionFinal)
            If dblImporteCalculado <> 0 Then
                dblImporteCalculado = dblImporteCalculado - ImporteRegistradoAcumuladoDeduccionEspecifica(strDeducciones(0), _
                PuestoLaboral, strPeriodo, CodigoLiquidacion, False) + De_Txt_a_Num_01(.txtSeguroOptativo.Text)
            ElseIf De_Txt_a_Num_01(.txtSeguroOptativo.Text) > 0 Then
                dblImporteCalculado = De_Txt_a_Num_01(.txtSeguroOptativo.Text)
            Else 'Buscamos el Seguro de Vida informado por F.572 Web si no tiene Seg.Vida Opt. en recibo
                dblImporteCalculado = CalcularDeduccionGeneralEspecificaSIRADIG(strIDLiquidacion, _
                strCodigoSIRADIG, PuestoLaboral, CodigoLiquidacion)
            End If
'            If dblImporteCalculado = 0 Then 'Buscamos el Seguro de Vida ACUMULADO en liquidaciones anteriores
'                dblImporteCalculado = CalcularDeduccionSeguroDeVida(PuestoLaboral, _
'                CodigoLiquidacion, dblImporteCalculado)
'            End If
            dblImporteMensualAcumulado = dblImporteMensualAcumulado + dblImporteCalculado
            dblImporteAcumulado = ImporteRegistradoAcumuladoDeduccionEspecifica(strDeducciones(0), _
            PuestoLaboral, strPeriodo, CodigoLiquidacion, False) + dblImporteCalculado
            If dblImporteCalculado <> 0 Or dblImporteAcumulado <> 0 Then
                Call CargardgDeduccionesGeneralesLG4taSIRADIG(strCodigoSIRADIG, strConceptoSIRADIG, _
                dblImporteCalculado, dblImporteAcumulado)
            End If
            'Agregamos el concepto Cuota Sindical
            strDeducciones(0) = "CuotaSindical"
            strCodigoSIRADIG = EquipararConceptoDeduccionSISPERconCodigoSIRADIG(strDeducciones(0))
            strConceptoSIRADIG = BuscarDenominacionDeduccionSIRADIG(strCodigoSIRADIG)
            If LiquidacionFinal = True Then
                'El nuevo procediemiento (genera problemas si me.txtcuotasindical es mayor a 0 en la liquidacion final)
                dblImporteCalculado = CalcularDescuentoCuotaSindicalAcumulado(PuestoLaboral, _
                CodigoLiquidacion, LiquidacionFinal, LiquidacionFinal)
                If dblImporteCalculado <> 0 Then
                    dblImporteCalculado = dblImporteCalculado - ImporteRegistradoAcumuladoDeduccionEspecifica(strDeducciones(0), _
                    PuestoLaboral, strPeriodo, CodigoLiquidacion, False) + De_Txt_a_Num_01(.txtCuotaSindical.Text)
                ElseIf De_Txt_a_Num_01(.txtCuotaSindical.Text) > 0 Then
                    dblImporteCalculado = De_Txt_a_Num_01(.txtCuotaSindical.Text)
    '            Else 'Buscamos el Seguro de Vida informado por F.572 Web si no tiene Seg.Vida Opt. en recibo
    '                dblImporteCalculado = CalcularDeduccionGeneralEspecificaSIRADIG(strIDLiquidacion, _
    '                strCodigoSIRADIG, PuestoLaboral, CodigoLiquidacion)
                End If
                dblImporteMensualAcumulado = dblImporteMensualAcumulado + dblImporteCalculado
                dblImporteAcumulado = ImporteRegistradoAcumuladoDeduccionEspecifica(strDeducciones(0), _
                PuestoLaboral, strPeriodo, CodigoLiquidacion, False) + dblImporteCalculado
            Else
                'El anterior procedimiento (no funciona en la liquidación final)
                dblImporteCalculado = De_Txt_a_Num_01(.txtCuotaSindical.Text)
                dblImporteMensualAcumulado = dblImporteMensualAcumulado + dblImporteCalculado
                dblImporteAcumulado = ImporteRegistradoAcumuladoDeduccionEspecifica(strDeducciones(0), _
                PuestoLaboral, strPeriodo, CodigoLiquidacion, False) + dblImporteCalculado
            End If
            If dblImporteCalculado <> 0 Or dblImporteAcumulado <> 0 Then
                Call CargardgDeduccionesGeneralesLG4taSIRADIG(strCodigoSIRADIG, strConceptoSIRADIG, _
                dblImporteCalculado, dblImporteAcumulado)
            End If
            'Agregamos las restantes deducciones que están fuera de recibo y por SIRADIG
            ReDim strDeducciones(1)
            strDeducciones(0) = "ServicioDomestico"
            strDeducciones(1) = "Alquileres"
            For d = 0 To UBound(strDeducciones)
                strCodigoSIRADIG = EquipararConceptoDeduccionSISPERconCodigoSIRADIG(strDeducciones(d))
                strConceptoSIRADIG = BuscarDenominacionDeduccionSIRADIG(strCodigoSIRADIG)
                dblImporteCalculado = CalcularDeduccionGeneralEspecificaSIRADIG(strIDLiquidacion, _
                strCodigoSIRADIG, PuestoLaboral, CodigoLiquidacion)
                dblImporteMensualAcumulado = dblImporteMensualAcumulado + dblImporteCalculado
                dblImporteAcumulado = ImporteRegistradoAcumuladoDeduccionEspecifica(strDeducciones(d), _
                PuestoLaboral, strPeriodo, CodigoLiquidacion, False) + dblImporteCalculado
                If dblImporteCalculado <> 0 Or dblImporteAcumulado <> 0 Then
                    Call CargardgDeduccionesGeneralesLG4taSIRADIG(strCodigoSIRADIG, strConceptoSIRADIG, _
                    dblImporteCalculado, dblImporteAcumulado)
                End If
            Next
            'Agregamos el Subtotal Deducciones Generales
            dblImporteAcumulado = CalcularDeduccionesGeneralesAcumuladas(PuestoLaboral, CodigoLiquidacion, False, True, True) + dblImporteMensualAcumulado
            Call CargardgDeduccionesGeneralesLG4taSIRADIG("ST", "-SUBTOTAL-", dblImporteMensualAcumulado, dblImporteAcumulado)
            'Calculamos la Ganancia Neta antes de las siguientes deducciones
            dblRdoNetoAntesDeDeducciones = De_Txt_a_Num_01(.txtGananciaBrutaAcumulada.Text)
            dblRdoNetoAntesDeDeducciones = dblRdoNetoAntesDeDeducciones - dblImporteAcumulado
            ReDim strDeducciones(2)
            strDeducciones(0) = "CuotaMedicoAsistencial"
            strDeducciones(1) = "Donaciones"
            strDeducciones(2) = "HonorariosMedicos"
            For d = 0 To UBound(strDeducciones)
                strCodigoSIRADIG = EquipararConceptoDeduccionSISPERconCodigoSIRADIG(strDeducciones(d))
                strConceptoSIRADIG = BuscarDenominacionDeduccionSIRADIG(strCodigoSIRADIG)
                If strDeducciones(d) = "HonorariosMedicos" And LiquidacionFinal = False Then
                    dblImporteCalculado = 0
                Else
                    dblImporteCalculado = CalcularDeduccionGeneralEspecificaSIRADIG(strIDLiquidacion, _
                    strCodigoSIRADIG, PuestoLaboral, CodigoLiquidacion, dblRdoNetoAntesDeDeducciones)
                End If
                dblImporteMensualAcumulado = dblImporteMensualAcumulado + dblImporteCalculado
                dblImporteAcumulado = ImporteRegistradoAcumuladoDeduccionEspecifica(strDeducciones(d), _
                PuestoLaboral, strPeriodo, CodigoLiquidacion, False) + dblImporteCalculado
                If dblImporteCalculado <> 0 Or dblImporteAcumulado <> 0 Then
                    Call CargardgDeduccionesGeneralesLG4taSIRADIG(strCodigoSIRADIG, strConceptoSIRADIG, _
                    dblImporteCalculado, dblImporteAcumulado)
                End If
            Next
            dblImporteAcumulado = CalcularDeduccionesGeneralesAcumuladas(PuestoLaboral, CodigoLiquidacion, False, False, True) + dblImporteMensualAcumulado
            Call CargardgDeduccionesGeneralesLG4taSIRADIG("T", "-TOTAL-", dblImporteMensualAcumulado, dblImporteAcumulado)
            .txtDeduccionesGeneralesAcumuladas.Text = De_Num_a_Tx_01(dblImporteAcumulado)
    
        Case "DeduccionesPersonales"
            'Cargamos las Deducciones Personales
            dblImporteMensualAcumulado = 0
            ReDim strDeducciones(3)
            strDeducciones(0) = "MNI"
            strDeducciones(1) = "C"
            strDeducciones(2) = "H"
            strDeducciones(3) = "OCF"
            ConfigurardgDeduccionesPersonalesLG4taSIRADIG
            For d = 0 To UBound(strDeducciones)
                strDenominacion = BuscarDenominacionDeduccionPersonalSLAVE(strDeducciones(d))
                dblImporteCalculado = CalcularDeduccionPersonalEspecificaSIRADIG(strIDLiquidacion, _
                strDenominacion, PuestoLaboral, CodigoLiquidacion)
                dblImporteMensualAcumulado = dblImporteMensualAcumulado + dblImporteCalculado
                dblImporteAcumulado = ImporteRegistradoAcumuladoDeduccionEspecifica(strDenominacion, _
                PuestoLaboral, strPeriodo, CodigoLiquidacion) + dblImporteCalculado
                Call CargardgDeduccionesPersonalesLG4taSIRADIG(strDeducciones(d), strDenominacion, _
                dblImporteCalculado, dblImporteAcumulado)
            Next
            dblImporteAcumulado = CalcularDeduccionesPersonalesAcumuladas(PuestoLaboral, CodigoLiquidacion, False, True) + dblImporteMensualAcumulado
            Call CargardgDeduccionesPersonalesLG4taSIRADIG("ST", "-SUBTOTAL-", dblImporteMensualAcumulado, dblImporteAcumulado)
            .txtDeduccionesPersonalesAcumuladas.Text = De_Num_a_Tx_01(dblImporteAcumulado)
            'Calculamos la Ganancia Neta antes de las siguiente deduccion
            dblRdoNetoAntesDeDeducciones = De_Txt_a_Num_01(.txtGananciaBrutaAcumulada.Text)
            dblRdoNetoAntesDeDeducciones = dblRdoNetoAntesDeDeducciones - De_Txt_a_Num_01(.txtDeduccionesGeneralesAcumuladas.Text)
            dblRdoNetoAntesDeDeducciones = dblRdoNetoAntesDeDeducciones - dblImporteAcumulado
            'Agregamos por separado la Deducción Especial
            strDenominacion = BuscarDenominacionDeduccionPersonalSLAVE("DE")
            dblImporteCalculado = CalcularDeduccionPersonalEspecificaSIRADIG(strIDLiquidacion, _
            strDenominacion, PuestoLaboral, CodigoLiquidacion, dblRdoNetoAntesDeDeducciones)
            dblImporteMensualAcumulado = dblImporteMensualAcumulado + dblImporteCalculado
            dblImporteAcumulado = ImporteRegistradoAcumuladoDeduccionEspecifica(strDenominacion, _
            PuestoLaboral, strPeriodo, CodigoLiquidacion) + dblImporteCalculado
            Call CargardgDeduccionesPersonalesLG4taSIRADIG("DE", strDenominacion, _
            dblImporteCalculado, dblImporteAcumulado)
            dblImporteAcumulado = CalcularDeduccionesPersonalesAcumuladas(PuestoLaboral, CodigoLiquidacion, False, False) + dblImporteMensualAcumulado
            Call CargardgDeduccionesPersonalesLG4taSIRADIG("T", "-TOTAL-", dblImporteMensualAcumulado, dblImporteAcumulado)
            .txtDeduccionesPersonalesAcumuladas.Text = De_Num_a_Tx_01(dblImporteAcumulado)
        
        Case "RetencionTotal"
            '8)Retención total que debería practicarse
            '8.1)Base Imponible
            '8.1.a)Ganancia Bruta
            dblImporteCalculado = De_Txt_a_Num_01(.txtGananciaBrutaAcumulada.Text)
            dblImporteAcumulado = dblImporteCalculado
            '8.1.b)Deducciones Generales Acumuladas
            dblImporteCalculado = De_Txt_a_Num_01(.txtDeduccionesGeneralesAcumuladas.Text)
            dblImporteAcumulado = dblImporteAcumulado - dblImporteCalculado
            '8.1.c)Deducciones Personales Acumuladas
            dblImporteCalculado = De_Txt_a_Num_01(.txtDeduccionesPersonalesAcumuladas.Text)
            dblImporteAcumulado = dblImporteAcumulado - dblImporteCalculado
            dblImporteAcumulado = Round(dblImporteAcumulado, 2)
            .txtBaseImponible.Text = De_Num_a_Tx_01(dblImporteAcumulado)
            If dblImporteAcumulado > 0 Then
                '8.2)Porcentaje Aplicable (dblImporteAcumulado = Base Imponible)
                dblImporteCalculado = CalcularAlicuotaAplicable(PuestoLaboral, CodigoLiquidacion, dblImporteAcumulado)
                dblImporteCalculado = dblImporteCalculado
                .txtPorcentajeAplicable.Text = De_Num_a_Tx_01(dblImporteCalculado * 100, True) & " %"
                '8.3)Importe Variable (dblImporteAcumulado = Base Imponible y dblImporteCalculado = Alicuota)
                dblImporteCalculado = CalcularImporteVariable(PuestoLaboral, CodigoLiquidacion, dblImporteCalculado, dblImporteAcumulado)
                .txtSumaVariable.Text = De_Num_a_Tx_01(dblImporteCalculado)
                '8.4)Importe Fijo (dblImporteAcumulado = Base Imponible)
                dblImporteCalculado = CalcularImporteFijo(PuestoLaboral, CodigoLiquidacion, dblImporteAcumulado)
                .txtSumaFija.Text = De_Num_a_Tx_01(dblImporteCalculado)
                '8.5)Subtotal Retención (Retención que debería practicarse)
                dblImporteCalculado = De_Txt_a_Num_01(.txtSumaVariable.Text) + De_Txt_a_Num_01(.txtSumaFija.Text)
                .txtSubtotalRentencion.Text = De_Num_a_Tx_01(dblImporteCalculado)
            Else
                .txtPorcentajeAplicable.Text = De_Num_a_Tx_01(0)
                .txtSumaVariable.Text = De_Num_a_Tx_01(0)
                .txtSumaFija.Text = De_Num_a_Tx_01(0)
                .txtSubtotalRentencion.Text = De_Num_a_Tx_01(0)
            End If
        Case "RetencionAcumulada"
            '9)Retención Acumulada
            dblImporteCalculado = CalcularRetencionAcumulada(PuestoLaboral, CodigoLiquidacion, False)
            .txtRetencionAcumulada.Text = De_Num_a_Tx_01(dblImporteCalculado)
        Case "AjusteRetencion"
            '10)Ajuste Retención Ganancias
            dblImporteCalculado = CalcularAjusteRetencion(PuestoLaboral, CodigoLiquidacion)
            .txtAjuesteRetencion.Text = De_Num_a_Tx_01(dblImporteCalculado)
        Case "RetencionPeriodo"
            '11)Retención del Período
            dblImporteCalculado = De_Txt_a_Num_01(.txtSubtotalRentencion.Text) - _
            De_Txt_a_Num_01(.txtRetencionAcumulada.Text) + De_Txt_a_Num_01(.txtAjuesteRetencion.Text)
            .txtRetencionPeriodo.Text = De_Num_a_Tx_01(dblImporteCalculado)
        End Select
    End With

    dblImporteCalculado = 0
    dblImporteAcumulado = 0
    dblRdoNetoAntesDeDeducciones = 0
    strCLPrevio = ""
    i = 0

End Sub

