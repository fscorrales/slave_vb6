Attribute VB_Name = "AgregarRegistros"
Public Sub GenerarAgente()

    Dim SQL As String

    If ValidarAgente = True Then
        Set rstRegistroSlave = New ADODB.Recordset
        With CargaAgente
            If strEditandoAgente = "" Then
                rstRegistroSlave.Open "AGENTES", dbSlave, adOpenForwardOnly, adLockOptimistic
                rstRegistroSlave.AddNew
            Else
                SQL = "Select * from AGENTES Where PUESTOLABORAL = " & "'" & strEditandoAgente & "'"
                rstRegistroSlave.Open SQL, dbSlave, adOpenForwardOnly, adLockOptimistic
                SQL = ""
                strEditandoAgente = ""
            End If
            rstRegistroSlave!PuestoLaboral = .txtPuestoLaboral.Text
            rstRegistroSlave!CUIL = .txtCUIL.Text
            rstRegistroSlave!NombreCompleto = .txtDescripcion.Text
            rstRegistroSlave!Legajo = .txtLegajo.Text
            rstRegistroSlave!Activado = .chkActivado.Value
            rstRegistroSlave.Update
        End With
        rstRegistroSlave.Close
        Set rstRegistroSlave = Nothing
        Unload CargaAgente
        ListadoAgentes.Show
        Call ConfigurardgAgentes(ListadoAgentes)
        Call CargardgAgentes(ListadoAgentes)
    End If
    
End Sub

Public Sub GenerarConcepto()

    Dim SQL As String

    If ValidarConcepto = True Then
        Set rstRegistroSlave = New ADODB.Recordset
        With CargaConcepto
            If strEditandoConcepto = "" Then
                rstRegistroSlave.Open "CONCEPTOS", dbSlave, adOpenForwardOnly, adLockOptimistic
                rstRegistroSlave.AddNew
            Else
                SQL = "Select * from CONCEPTOS Where CODIGO = " & "'" & strEditandoConcepto & "'"
                rstRegistroSlave.Open SQL, dbSlave, adOpenForwardOnly, adLockOptimistic
                SQL = ""
                strEditandoConcepto = ""
            End If
            rstRegistroSlave!Codigo = Format(.txtCodigo.Text, "0000")
            rstRegistroSlave!Denominacion = .txtDenominacion.Text
            rstRegistroSlave.Update
        End With
        rstRegistroSlave.Close
        Set rstRegistroSlave = Nothing
        Unload CargaConcepto
        ListadoConceptos.Show
        ConfigurardgConceptos
        CargardgConceptos
    End If
    
End Sub

Public Sub GenerarNormaEscalaGanancias()

    Dim SQL As String
    Dim datFecha As Date

    If ValidarNormaEscalaGanancias = True Then
        With CargaNormaEscalaGanancias
            'Por las dudas convertirmos a date la fecha almacenada como string
            datFecha = DateTime.DateSerial(Right(.txtFecha.Text, 4), Mid(.txtFecha.Text, 4, 2), Left(.txtFecha.Text, 2))
            If strEditandoNormaEscalaGanancias = "" Then
                Set rstRegistroSlave = New ADODB.Recordset
                rstRegistroSlave.Open "ESCALAAPLICABLEGANANCIAS", dbSlave, adOpenForwardOnly, adLockOptimistic
                rstRegistroSlave.AddNew
                rstRegistroSlave!NormaLegal = .txtNormaLegal.Text
                rstRegistroSlave!Fecha = datFecha
                'Generamos una nueva norma con una sola escala en 0 que luego hay que editar
                rstRegistroSlave!Tramo = "0"
                rstRegistroSlave!ImporteMaximo = 0
                rstRegistroSlave!ImporteFijo = 0
                rstRegistroSlave!ImporteVariable = 0
                rstRegistroSlave.Update
                rstRegistroSlave.Close
                Set rstRegistroSlave = Nothing
            Else
                'Identificamos la Norma Legal a editar
                Dim strNormaAEditar As String
                strNormaAEditar = strEditandoNormaEscalaGanancias
                strEditandoNormaEscalaGanancias = ""
                'Identificamos la Fecha de la Norma Legal a editar
                Dim datFechaAEditar As Date
                SQL = "Select FECHA From ESCALAAPLICABLEGANANCIAS" _
                & " Where NORMALEGAL = '" & strNormaAEditar _
                & "' Group By NORMALEGAL, FECHA"
                Set rstBuscarSlave = New ADODB.Recordset
                rstBuscarSlave.Open SQL, dbSlave, adOpenForwardOnly, adLockOptimistic
                datFechaAEditar = rstBuscarSlave!Fecha
                rstBuscarSlave.Close
                Set rstBuscarSlave = Nothing
                'Para poder editar una norma legal existente debemos copiar _
                la vieja norma ,con el numero de norma y/o fecha modificados, y luego _
                borrar todo lo relacionado con la norma previa a la edicion
                'SQL INSERT INTO SELECT Statement
                SQL = "Insert Into ESCALAAPLICABLEGANANCIAS" _
                & " (NormaLegal, Fecha, Tramo, ImporteMaximo, ImporteFijo, ImporteVariable)" _
                & " Select '" & .txtNormaLegal & "', #" & Format(datFecha, "MM/DD/YYYY") & "#, Tramo, ImporteMaximo, ImporteFijo, ImporteVariable" _
                & " From ESCALAAPLICABLEGANANCIAS Where NORMALEGAL = '" & strNormaAEditar & "'"
                dbSlave.BeginTrans
                dbSlave.Execute SQL
                dbSlave.CommitTrans
                SQL = "Delete * from ESCALAAPLICABLEGANANCIAS" _
                & " Where NORMALEGAL = " & "'" & strNormaAEditar _
                & "' And FECHA = #" & Format(datFecha, "MM/DD/YYYY") & "#"
                dbSlave.BeginTrans
                dbSlave.Execute SQL
                dbSlave.CommitTrans
            End If
            
        End With
        Unload CargaNormaEscalaGanancias
        ListadoEscalaGanancias.Show
        ConfigurardgNormasEscalaGanancias
        CargardgNormasEscalaGanancias
        ConfigurardgEscalaGanancias
        CargardgEscalaGanancias (ListadoEscalaGanancias.dgNormasEscalaGanancias.TextMatrix(1, 0))
    End If
    
End Sub

Public Sub GenerarEscalaGanancias()

    Dim SQL As String
    Dim datFecha As Date

    If ValidarEscalaGanancias = True Then
        Set rstRegistroSlave = New ADODB.Recordset
        With CargaEscalaGanancias
            datFecha = DateTime.DateSerial(Right(.txtFecha.Text, 4), Mid(.txtFecha.Text, 4, 2), Left(.txtFecha.Text, 2))
            If strEditandoEscalaGanancias = "" Then
                rstRegistroSlave.Open "ESCALAAPLICABLEGANANCIAS", dbSlave, adOpenForwardOnly, adLockOptimistic
                rstRegistroSlave.AddNew
            Else
                SQL = "Select * from ESCALAAPLICABLEGANANCIAS" _
                & " Where TRAMO = '" & strEditandoEscalaGanancias _
                & "' And NORMALEGAL = '" & .txtNormaLegal.Text & "'"
                rstRegistroSlave.Open SQL, dbSlave, adOpenForwardOnly, adLockOptimistic
                SQL = ""
                strEditandoEscalaGanancias = ""
            End If
            rstRegistroSlave!NormaLegal = .txtNormaLegal.Text
            rstRegistroSlave!Fecha = datFecha
            rstRegistroSlave!Tramo = .txtTramo.Text
            rstRegistroSlave!ImporteMaximo = De_Txt_a_Num_01(.txtImporteMaximo.Text)
            Debug.Print De_Txt_a_Num_01(.txtImporteFijo.Text)
            rstRegistroSlave!ImporteFijo = De_Txt_a_Num_01(.txtImporteFijo.Text)
            rstRegistroSlave!ImporteVariable = .txtImporteVariable.Text / 100
            rstRegistroSlave.Update
        End With
        rstRegistroSlave.Close
        Set rstRegistroSlave = Nothing
        Unload CargaEscalaGanancias
        ListadoEscalaGanancias.Show
        ConfigurardgNormasEscalaGanancias
        CargardgNormasEscalaGanancias
        ConfigurardgEscalaGanancias
        CargardgEscalaGanancias (ListadoEscalaGanancias.dgNormasEscalaGanancias.TextMatrix(1, 0))
    End If
    
End Sub

Public Sub GenerarLimitesDeducciones()

    Dim SQL As String

    If ValidarLimitesDeducciones = True Then
        Set rstRegistroSlave = New ADODB.Recordset
        With CargaLimitesDeducciones
            If strEditandoLimitesDeducciones = "" Then
                rstRegistroSlave.Open "DEDUCCIONES4TACATEGORIA", dbSlave, adOpenForwardOnly, adLockOptimistic
                rstRegistroSlave.AddNew
            Else
                SQL = "Select * from DEDUCCIONES4TACATEGORIA Where NORMALEGAL = " & "'" & strEditandoLimitesDeducciones & "'"
                rstRegistroSlave.Open SQL, dbSlave, adOpenForwardOnly, adLockOptimistic
                SQL = ""
                strEditandoLimitesDeducciones = ""
            End If
            rstRegistroSlave!NormaLegal = .txtNormaLegal.Text
            rstRegistroSlave!Fecha = .txtFecha.Text
            rstRegistroSlave!MinimoNoImponible = De_Txt_a_Num_01(.txtMinimoNoImponible.Text)
            rstRegistroSlave!DeduccionEspecial = De_Txt_a_Num_01(.txtDeduccionEspecial.Text)
            rstRegistroSlave!Hijo = De_Txt_a_Num_01(.txtHijo.Text)
            rstRegistroSlave!Conyuge = De_Txt_a_Num_01(.txtConyuge.Text)
            rstRegistroSlave!OtrasCargasDeFamilia = De_Txt_a_Num_01(.txtOtrasCargas.Text)
            rstRegistroSlave!ServicioDomestico = De_Txt_a_Num_01(.txtServicioDomestico.Text)
            rstRegistroSlave!SeguroDeVida = De_Txt_a_Num_01(.txtSeguroDeVida.Text)
            rstRegistroSlave!Alquileres = De_Txt_a_Num_01(.txtAlquileres.Text)
            rstRegistroSlave!HonorariosMedicos = .txtHonorariosMedicos.Text / 100
            rstRegistroSlave!CuotaMedicoAsistencial = .txtCuotaMedico.Text / 100
            rstRegistroSlave!Donaciones = .txtDonaciones.Text / 100
            rstRegistroSlave.Update
        End With
        rstRegistroSlave.Close
        Set rstRegistroSlave = Nothing
        Unload CargaLimitesDeducciones
        ListadoLimitesDeducciones.Show
        ConfigurardgLimitesDeducciones
        CargardgLimitesDeducciones
    End If
    
End Sub

Public Sub GenerarParentesco()

    Dim SQL As String

    If ValidarParentesco = True Then
        Set rstRegistroSlave = New ADODB.Recordset
        With CargaParentesco
            If strEditandoParentesco = "" Then
                rstRegistroSlave.Open "ASIGNACIONESFAMILIARES", dbSlave, adOpenForwardOnly, adLockOptimistic
                rstRegistroSlave.AddNew
            Else
                SQL = "Select * from ASIGNACIONESFAMILIARES Where CODIGO = " & "'" & strEditandoParentesco & "'"
                rstRegistroSlave.Open SQL, dbSlave, adOpenForwardOnly, adLockOptimistic
                SQL = ""
                strEditandoParentesco = ""
            End If
            rstRegistroSlave!Codigo = Format(.txtCodigo.Text, "00")
            rstRegistroSlave!Denominacion = .txtDenominacion.Text
            rstRegistroSlave!Importe = De_Txt_a_Num_01(.txtImporte.Text)
            rstRegistroSlave.Update
        End With
        rstRegistroSlave.Close
        Set rstRegistroSlave = Nothing
        Unload CargaParentesco
        ListadoParentesco.Show
        ConfigurardgParentesco
        CargardgParentesco
    End If
    
End Sub

Public Sub GenerarDeduccionesGenerales()

    Dim SQL As String

    If ValidarDeduccionesGenerales = True Then
        Set rstRegistroSlave = New ADODB.Recordset
        With CargaDeduccionesGenerales
            If strEditandoDeduccionesGenerales = "" Then
                rstRegistroSlave.Open "IMPORTEDEDUCCIONESGENERALES", dbSlave, adOpenForwardOnly, adLockOptimistic
                rstRegistroSlave.AddNew
            Else
                SQL = "Select * from IMPORTEDEDUCCIONESGENERALES Where PUESTOLABORAL= '" & .txtPuestoLaboral.Text & "' And FECHA = #" & Format(strEditandoDeduccionesGenerales, "YYYY/MM/DD") & "#"
                rstRegistroSlave.Open SQL, dbSlave, adOpenForwardOnly, adLockOptimistic
                SQL = ""
                strEditandoDeduccionesGenerales = ""
            End If
            rstRegistroSlave!PuestoLaboral = .txtPuestoLaboral.Text
            rstRegistroSlave!Fecha = .txtFecha.Text
            rstRegistroSlave!ServicioDomestico = De_Txt_a_Num_01(.txtServicioDomestico.Text)
            rstRegistroSlave!SeguroDeVida = De_Txt_a_Num_01(.txtSeguroDeVida.Text)
            rstRegistroSlave!Alquileres = De_Txt_a_Num_01(.txtAlquileres.Text)
            rstRegistroSlave!HonorariosMedicos = De_Txt_a_Num_01(.txtHonorariosMedicos.Text)
            rstRegistroSlave!CuotaMedicoAsistencial = De_Txt_a_Num_01(.txtCuotaMedico.Text)
            rstRegistroSlave!Donaciones = De_Txt_a_Num_01(.txtDonaciones.Text)
            rstRegistroSlave.Update
        End With
        rstRegistroSlave.Close
        Set rstRegistroSlave = Nothing
        Unload CargaDeduccionesGenerales
        ListadoDeduccionesGenerales.Show
        Call ConfigurardgAgentes(ListadoDeduccionesGenerales)
        Call CargardgAgentes(ListadoDeduccionesGenerales)
        ConfigurardgDeduccionesGenerales
        Call CargardgDeduccionesGenerales(ListadoDeduccionesGenerales.dgAgentes.TextMatrix(1, 2))
    End If
    
End Sub

Public Sub GenerarDeduccionesPersonales()

    Dim SQL As String
    Dim i As Integer
    Dim strNombreFamiliar As String
       
    Set rstRegistroSlave = New ADODB.Recordset
    With CargaDeduccionesPersonales
        For i = 0 To .lstFamiliaresDeduciblesGanancias.ListCount - 1
            .lstFamiliaresDeduciblesGanancias.ListIndex = i
            strNombreFamiliar = .lstFamiliaresDeduciblesGanancias.Text
            SQL = "Select * from CARGASDEFAMILIA Where PUESTOLABORAL= '" & .txtPuestoLaboral.Text & "' And NombreCompleto= '" & strNombreFamiliar & "'"
            rstRegistroSlave.Open SQL, dbSlave, adOpenForwardOnly, adLockOptimistic
            SQL = ""
            rstRegistroSlave!DeducibleGanancias = True
            rstRegistroSlave.Update
            rstRegistroSlave.Close
        Next i
    End With

    i = 0
    Set rstRegistroSlave = Nothing
    Unload CargaDeduccionesPersonales
    ListadoDeduccionesPersonales.Show
    Call ConfigurardgAgentes(ListadoDeduccionesPersonales)
    Call CargardgAgentes(ListadoDeduccionesPersonales)
    ConfigurardgDeduccionesPersonales
    Call CargardgDeduccionesPersonales(ListadoDeduccionesPersonales.dgAgentes.TextMatrix(1, 2))
    
End Sub

Public Sub GenerarCodigoLiquidacion()

    Dim SQL As String

    If ValidarCodigoLiquidacion = True Then
        Set rstRegistroSlave = New ADODB.Recordset
        With CargaCodigoLiquidacion
            If strEditandoCodigoLiquidacion = "" Then
                rstRegistroSlave.Open "CODIGOLIQUIDACIONES", dbSlave, adOpenForwardOnly, adLockOptimistic
                rstRegistroSlave.AddNew
            Else
                SQL = "Select * from CODIGOLIQUIDACIONES Where CODIGO = " & "'" & strEditandoCodigoLiquidacion & "'"
                rstRegistroSlave.Open SQL, dbSlave, adOpenForwardOnly, adLockOptimistic
                SQL = ""
                strEditandoCodigoLiquidacion = ""
            End If
            rstRegistroSlave!Codigo = .txtCodigo.Text
            rstRegistroSlave!Periodo = .txtPeriodo.Text
            rstRegistroSlave!Descripcion = .txtDescripcion.Text
            rstRegistroSlave!MontoExento = De_Txt_a_Num_01(.txtMontoExento.Text)
            rstRegistroSlave.Update
        End With
        rstRegistroSlave.Close
        Set rstRegistroSlave = Nothing
        Unload CargaCodigoLiquidacion
        ListadoCodigoLiquidaciones.Show
        ConfigurardgCodigoLiquidacion
        CargardgCodigoLiquidacion
    End If
    
End Sub

Public Sub GenerarFamiliar()

    Dim SQL As String

    If ValidarFamiliar = True Then
        Set rstRegistroSlave = New ADODB.Recordset
        With CargaFamiliar
            If strEditandoFamiliar = "" Then
                rstRegistroSlave.Open "CARGASDEFAMILIA", dbSlave, adOpenForwardOnly, adLockOptimistic
                rstRegistroSlave.AddNew
            Else
                SQL = "Select * from CARGASDEFAMILIA Where PUESTOLABORAL= '" & .txtPuestoLaboral.Text & "' And DNI = '" & strEditandoFamiliar & "'"
                rstRegistroSlave.Open SQL, dbSlave, adOpenForwardOnly, adLockOptimistic
                SQL = ""
                strEditandoFamiliar = ""
            End If
            rstRegistroSlave!PuestoLaboral = .txtPuestoLaboral
            rstRegistroSlave!DNI = .txtDNI.Text
            rstRegistroSlave!NombreCompleto = .txtDescripcionFamiliar.Text
            Set rstBuscarSlave = New ADODB.Recordset
            SQL = "Select * from ASIGNACIONESFAMILIARES Where PARENTESCO= '" & .cmbParentesco.Text & "'"
            rstBuscarSlave.Open SQL, dbSlave, adOpenForwardOnly, adLockOptimistic
            rstRegistroSlave!CodigoParentesco = rstBuscarSlave!Codigo
            rstBuscarSlave.Close
            Set rstBuscarSlave = Nothing
            rstRegistroSlave!FechaAlta = .txtFechaAlta.Text
            rstRegistroSlave!NivelDeEstudio = .cmbNivelDeEstudio.Text
            rstRegistroSlave!CobraSalario = .chkCobraSalario.Value
            rstRegistroSlave!AdherenteObraSocial = .chkAdherente.Value
            rstRegistroSlave!Discapacitado = .chkDiscapacitado.Value
            rstRegistroSlave!DeducibleGanancias = .chkGanancias.Value
            rstRegistroSlave.Update
        End With
        rstRegistroSlave.Close
        Set rstRegistroSlave = Nothing
        Unload CargaFamiliar
        ListadoFamiliares.Show
        Call ConfigurardgAgentes(ListadoFamiliares)
        Call CargardgAgentes(ListadoFamiliares)
        ConfigurardgFamiliares
        Call CargardgFamiliares(ListadoFamiliares.dgAgentes.TextMatrix(1, 2))
    End If
    
End Sub

Public Sub GenerarLiquidacionGanancia4taViejo()

    Dim SQL As String

    If ValidarLiquidacionGanancias = True Then
        Set rstRegistroSlave = New ADODB.Recordset
        With LiquidacionGanancia4ta
            If bolEditandoRetencionGanancias = False Then
                rstRegistroSlave.Open "LIQUIDACIONGANANCIAS4TACATEGORIA", dbSlave, adOpenForwardOnly, adLockOptimistic
                rstRegistroSlave.AddNew
            Else
                SQL = "Select * from LIQUIDACIONGANANCIAS4TACATEGORIA Where PUESTOLABORAL = " & "'" & .txtPuestoLaboral.Text & "' " _
                & "And CODIGOLIQUIDACION = " & "'" & .txtCodigoLiquidacion.Text & "'"
                rstRegistroSlave.Open SQL, dbSlave, adOpenForwardOnly, adLockOptimistic
                SQL = ""
                bolEditandoRetencionGanancias = False
            End If
            rstRegistroSlave!CodigoLiquidacion = .txtCodigoLiquidacion.Text
            rstRegistroSlave!PuestoLaboral = .txtPuestoLaboral.Text
            rstRegistroSlave!HaberOptimo = De_Txt_a_Num_01(.txtHaberOptimo.Text)
            rstRegistroSlave!Pluriempleo = De_Txt_a_Num_01(.txtPluriempleo.Text)
            rstRegistroSlave!Ajuste = De_Txt_a_Num_01(.txtAjuste.Text)
            rstRegistroSlave!Jubilacion = De_Txt_a_Num_01(.txtJubilacion.Text)
            rstRegistroSlave!ObraSocial = De_Txt_a_Num_01(.txtObraSocial.Text)
            rstRegistroSlave!AdherenteObraSocial = De_Txt_a_Num_01(.txtAdherente.Text)
            rstRegistroSlave!SeguroDeVidaObligatorio = De_Txt_a_Num_01(.txtSeguroObligatorio.Text)
            rstRegistroSlave!SeguroDeVidaOptativo = De_Txt_a_Num_01(.txtSeguroOptativo.Text) + De_Txt_a_Num_01(.dgDeduccionesGenerales.TextMatrix(2, 1))
            rstRegistroSlave!CuotaSindical = De_Txt_a_Num_01(.txtCuotaSindical.Text)
            rstRegistroSlave!ServicioDomestico = De_Txt_a_Num_01(.dgDeduccionesGenerales.TextMatrix(1, 1))
            rstRegistroSlave!CuotaMedicoAsistencial = De_Txt_a_Num_01(.dgDeduccionesGenerales.TextMatrix(3, 1))
            rstRegistroSlave!Donaciones = De_Txt_a_Num_01(.dgDeduccionesGenerales.TextMatrix(4, 1))
            rstRegistroSlave!HonorariosMedicos = 0 'Determinar qué hacer en Liquidación Final / Anual
            rstRegistroSlave!MinimoNoImponible = De_Txt_a_Num_01(.dgDeduccionesPersonales.TextMatrix(1, 1))
            rstRegistroSlave!DeduccionEspecial = De_Txt_a_Num_01(.dgDeduccionesPersonales.TextMatrix(6, 1))
            rstRegistroSlave!Conyuge = De_Txt_a_Num_01(.dgDeduccionesPersonales.TextMatrix(3, 1))
            rstRegistroSlave!Hijo = De_Txt_a_Num_01(.dgDeduccionesPersonales.TextMatrix(4, 1))
            rstRegistroSlave!OtrasCargasDeFamilia = De_Txt_a_Num_01(.dgDeduccionesPersonales.TextMatrix(5, 1))
            rstRegistroSlave!AjusteRetencion = De_Txt_a_Num_01(.txtAjuesteRetencion.Text)
            rstRegistroSlave!Retencion = De_Txt_a_Num_01(.txtRetencionPeriodo.Text)
            rstRegistroSlave.Update
        End With
        rstRegistroSlave.Close
        Set rstRegistroSlave = Nothing
        Unload LiquidacionGanancia4ta
        ListadoLiquidacionGanancias.Show
        ConfigurardgCodigosLiquidacionesGanancias
        CargardgCodigosLiquidacionesGanancias
        ConfigurardgAgentesRetenidos
        Call CargardgAgentesRetenidos(ListadoLiquidacionGanancias.dgCodigosLiquidacionesGanancias.TextMatrix(1, 0))

    End If
    
End Sub

Public Sub GenerarLiquidacionGanancia4ta()

    Dim SQL As String
    Dim strPL As String
    Dim strCL As String
    Dim strCLPrevio As String
    Dim strPeriodo As String
    Dim bolLiquidacionFinal As Boolean
    Dim dblRdoAntesDeCMAyDyHM As Double
    Dim dblRdoNetoAntesDeDeduccionEspecial As Double
    Dim dblImporteCalculado As Double
    
    If ValidarLiquidacionGanancias = True Then
        Set rstRegistroSlave = New ADODB.Recordset
        With LiquidacionGanancia4ta
            bolLiquidacionFinal = .chkLiquidacionFinal.Value
            strPL = .txtPuestoLaboral.Text
            strCL = Left(.txtCodigoLiquidacion.Text, 4)
            strCLPrevio = BuscarCodigoLiquidacionAnterior(strCL, True)
            strPeriodo = BuscarPeriodoLiquidacion(strCL)
            If bolEditandoRetencionGanancias = False Then
                rstRegistroSlave.Open "LIQUIDACIONGANANCIAS4TACATEGORIA", dbSlave, adOpenForwardOnly, adLockOptimistic
                rstRegistroSlave.AddNew
            Else
                SQL = "Select * from LIQUIDACIONGANANCIAS4TACATEGORIA" _
                & " Where PUESTOLABORAL = " & "'" & strPL _
                & "' And CODIGOLIQUIDACION = " & "'" & strCL & "'"
                rstRegistroSlave.Open SQL, dbSlave, adOpenForwardOnly, adLockOptimistic
                SQL = ""
                bolEditandoRetencionGanancias = False
            End If
            rstRegistroSlave!CodigoLiquidacion = strCL
            rstRegistroSlave!PuestoLaboral = strPL
            rstRegistroSlave!HaberOptimo = De_Txt_a_Num_01(.txtHaberOptimo.Text)
            rstRegistroSlave!Pluriempleo = De_Txt_a_Num_01(.txtPluriempleo.Text)
            rstRegistroSlave!Ajuste = De_Txt_a_Num_01(.txtAjuste.Text)
            rstRegistroSlave!Jubilacion = De_Txt_a_Num_01(.txtJubilacion.Text)
            rstRegistroSlave!ObraSocial = De_Txt_a_Num_01(.txtObraSocial.Text)
            rstRegistroSlave!AdherenteObraSocial = De_Txt_a_Num_01(.txtAdherente.Text)
            rstRegistroSlave!SeguroDeVidaObligatorio = De_Txt_a_Num_01(.txtSeguroObligatorio.Text)
            dblRdoAntesDeCMAyDyHM = De_Txt_a_Num_01(.txtGananciaNeta.Text)
            dblImporteCalculado = CalcularDeduccionSeguroDeVida(strPL, strCL, De_Txt_a_Num_01(.txtSeguroOptativo.Text))
            rstRegistroSlave!SeguroDeVidaOptativo = dblImporteCalculado
            dblRdoAntesDeCMAyDyHM = dblRdoAntesDeCMAyDyHM - dblImporteCalculado _
            - ImporteRegistradoAcumuladoDeduccionEspecifica("SeguroDeVidaOptativo", strPL, strPeriodo, strCLPrevio)
            dblImporteCalculado = CalcularDeduccionCuotaSindical(strPL, strCL, De_Txt_a_Num_01(.txtCuotaSindical.Text))
            rstRegistroSlave!CuotaSindical = dblImporteCalculado
            dblRdoAntesDeCMAyDyHM = dblRdoAntesDeCMAyDyHM - dblImporteCalculado _
            - ImporteRegistradoAcumuladoDeduccionEspecifica("CuotaSindical", strPL, strPeriodo, strCLPrevio)
            dblImporteCalculado = CalcularDeduccionServicioDomestico(strPL, strCL)
            rstRegistroSlave!ServicioDomestico = dblImporteCalculado
            dblRdoAntesDeCMAyDyHM = dblRdoAntesDeCMAyDyHM - dblImporteCalculado _
            - ImporteRegistradoAcumuladoDeduccionEspecifica("ServicioDomestico", strPL, strPeriodo, strCLPrevio)
            dblImporteCalculado = CalcularDeduccionAlquiler(strPL, strCL)
            rstRegistroSlave!Alquileres = dblImporteCalculado
            dblRdoAntesDeCMAyDyHM = dblRdoAntesDeCMAyDyHM - dblImporteCalculado _
            - ImporteRegistradoAcumuladoDeduccionEspecifica("Alquileres", strPL, strPeriodo, strCLPrevio)
            dblRdoAntesDeDeduccionEspecial = dblRdoAntesDeCMAyDyHM
            dblImporteCalculado = CalcularDeduccionCuotaMedicoAsistencial(strPL, strCL, , dblRdoAntesDeCMAyDyHM)
            rstRegistroSlave!CuotaMedicoAsistencial = dblImporteCalculado
            dblRdoAntesDeDeduccionEspecial = dblRdoAntesDeDeduccionEspecial - dblImporteCalculado _
            - ImporteRegistradoAcumuladoDeduccionEspecifica("CuotaMedicoAsistencial", strPL, strPeriodo, strCLPrevio)
            dblImporteCalculado = CalcularDeduccionDonaciones(strPL, strCL, , dblRdoAntesDeCMAyDyHM)
            rstRegistroSlave!Donaciones = dblImporteCalculado
            dblRdoAntesDeDeduccionEspecial = dblRdoAntesDeDeduccionEspecial - dblImporteCalculado _
            - ImporteRegistradoAcumuladoDeduccionEspecifica("Donaciones", strPL, strPeriodo, strCLPrevio)
            If bolLiquidacionFinal = False Then
                dblImporteCalculado = 0
            Else
                dblImporteCalculado = CalcularDeduccionHonorariosMedicos(strPL, strCL, , dblRdoAntesDeCMAyDyHM)
            End If
            rstRegistroSlave!HonorariosMedicos = dblImporteCalculado
            dblRdoAntesDeDeduccionEspecial = dblRdoAntesDeDeduccionEspecial - dblImporteCalculado
            dblImporteCalculado = CalcularDeduccionMinimoNoImponible(strPL, strCL)
            rstRegistroSlave!MinimoNoImponible = dblImporteCalculado
            dblRdoAntesDeDeduccionEspecial = dblRdoAntesDeDeduccionEspecial - dblImporteCalculado _
            - ImporteRegistradoAcumuladoDeduccionEspecifica("MinimoNoImponible", strPL, strPeriodo, strCLPrevio)
            dblImporteCalculado = CalcularDeduccionConyuge(strPL, strCL)
            rstRegistroSlave!Conyuge = dblImporteCalculado
            dblRdoAntesDeDeduccionEspecial = dblRdoAntesDeDeduccionEspecial - dblImporteCalculado _
            - ImporteRegistradoAcumuladoDeduccionEspecifica("Conyuge", strPL, strPeriodo, strCLPrevio)
            dblImporteCalculado = CalcularDeduccionHijo(strPL, strCL)
            rstRegistroSlave!Hijo = dblImporteCalculado
            dblRdoAntesDeDeduccionEspecial = dblRdoAntesDeDeduccionEspecial - dblImporteCalculado _
            - ImporteRegistradoAcumuladoDeduccionEspecifica("Hijo", strPL, strPeriodo, strCLPrevio)
            dblImporteCalculado = CalcularDeduccionOtrasCargasDeFamilia(strPL, strCL)
            rstRegistroSlave!OtrasCargasDeFamilia = CalcularDeduccionOtrasCargasDeFamilia(strPL, strCL)
            dblRdoAntesDeDeduccionEspecial = dblRdoAntesDeDeduccionEspecial - dblImporteCalculado _
            - ImporteRegistradoAcumuladoDeduccionEspecifica("OtrasCargasDeFamilia", strPL, strPeriodo, strCLPrevio)
            dblImporteCalculado = CalcularDeduccionEspecial(strPL, strCL, , dblRdoAntesDeDeduccionEspecial)
            rstRegistroSlave!DeduccionEspecial = dblImporteCalculado
            rstRegistroSlave!AjusteRetencion = De_Txt_a_Num_01(.txtAjuesteRetencion.Text)
            rstRegistroSlave!Retencion = De_Txt_a_Num_01(.txtRetencionPeriodo.Text)
            rstRegistroSlave.Update
        End With
        rstRegistroSlave.Close
        Set rstRegistroSlave = Nothing
        Unload LiquidacionGanancia4ta
        ListadoLiquidacionGanancias.Show
        ConfigurardgCodigosLiquidacionesGanancias
        CargardgCodigosLiquidacionesGanancias
        ConfigurardgAgentesRetenidos
        Call CargardgAgentesRetenidos(ListadoLiquidacionGanancias.dgCodigosLiquidacionesGanancias.TextMatrix(1, 0))

    End If
    
End Sub

Public Sub GenerarLiquidacionGanancia4taSIRADIG()

    Dim SQL As String
    Dim strPL As String
    Dim strCL As String
    Dim i As Integer
    Dim strDenominacion As String
    
'    Dim strCLPrevio As String
    Dim strPeriodo As String
    Dim bolLiquidacionFinal As Boolean
'    Dim dblRdoAntesDeCMAyDyHM As Double
'    Dim dblRdoNetoAntesDeDeduccionEspecial As Double
'    Dim dblImporteCalculado As Double
    
    If ValidarLiquidacionGananciasSIRADIG = True Then
        Set rstRegistroSlave = New ADODB.Recordset
        With LiquidacionGanancia4taSIRADIG
            bolLiquidacionFinal = .chkLiquidacionFinal.Value
            strPL = .txtPuestoLaboral.Text
            strCL = Left(.txtCodigoLiquidacion.Text, 4)
'            strCLPrevio = BuscarCodigoLiquidacionAnterior(strCL, True)
            strPeriodo = BuscarPeriodoLiquidacion(strCL)
            If bolEditandoRetencionGanancias = False Then
                rstRegistroSlave.Open "LIQUIDACIONGANANCIAS4TACATEGORIA", dbSlave, adOpenForwardOnly, adLockOptimistic
                rstRegistroSlave.AddNew
            Else
                SQL = "Select * from LIQUIDACIONGANANCIAS4TACATEGORIA" _
                & " Where PUESTOLABORAL = " & "'" & strPL _
                & "' And CODIGOLIQUIDACION = " & "'" & strCL & "'"
                rstRegistroSlave.Open SQL, dbSlave, adOpenForwardOnly, adLockOptimistic
                SQL = ""
                bolEditandoRetencionGanancias = False
            End If
            'Precargamos los registros en 0
            For i = 1 To rstRegistroSlave.Fields.Count
                rstRegistroSlave.Fields(i - 1) = 0
            Next i
            'Identificación de la Liquidacion y el Agente
            rstRegistroSlave!CodigoLiquidacion = strCL
            rstRegistroSlave!PuestoLaboral = strPL
            'Importe Bruto
            rstRegistroSlave!HaberOptimo = De_Txt_a_Num_01(.txtHaberOptimo.Text)
            rstRegistroSlave!Pluriempleo = De_Txt_a_Num_01(.txtPluriempleo.Text)
            rstRegistroSlave!Ajuste = De_Txt_a_Num_01(.txtAjuste.Text)
            'Deducciones Generales
            i = .dgDeduccionesGenerales.Rows
            If i > 2 Then
                For i = 2 To .dgDeduccionesGenerales.Rows
                    strDenominacion = .dgDeduccionesGenerales.TextMatrix((i - 1), 0)
                    If strDenominacion <> "ST" And strDenominacion <> "T" Then
                        strDenominacion = EquipararCodigoDeduccionSIRADIGconConceptoSISPER(strDenominacion)
                        rstRegistroSlave.Fields(strDenominacion) = De_Txt_a_Num_01(.dgDeduccionesGenerales.TextMatrix((i - 1), 2))
                    End If
                Next i
            End If
            'Deducciones Personales
            i = .dgDeduccionesPersonales.Rows
            If i > 2 Then
                For i = 2 To .dgDeduccionesPersonales.Rows
                    strDenominacion = .dgDeduccionesPersonales.TextMatrix((i - 1), 0)
                    If strDenominacion <> "ST" And strDenominacion <> "T" Then
                        strDenominacion = .dgDeduccionesPersonales.TextMatrix((i - 1), 1)
                        rstRegistroSlave.Fields(strDenominacion) = De_Txt_a_Num_01(.dgDeduccionesPersonales.TextMatrix((i - 1), 2))
                    End If
                Next i
            End If
            'Retencion
            rstRegistroSlave!AjusteRetencion = De_Txt_a_Num_01(.txtAjuesteRetencion.Text)
            rstRegistroSlave!Retencion = De_Txt_a_Num_01(.txtRetencionPeriodo.Text)
            rstRegistroSlave.Update
        End With
        rstRegistroSlave.Close
        Set rstRegistroSlave = Nothing
        Unload LiquidacionGanancia4taSIRADIG
        ListadoLiquidacionGanancias.Show
        ConfigurardgCodigosLiquidacionesGanancias
        CargardgCodigosLiquidacionesGanancias
        ConfigurardgAgentesRetenidos
        Call CargardgAgentesRetenidos(ListadoLiquidacionGanancias.dgCodigosLiquidacionesGanancias.TextMatrix(1, 0))

    End If
    
End Sub

Public Sub GenerarComprobanteSIIF()

    Dim SQL As String
    Dim strComprobante As String
    Dim datFecha As Date
    Dim strTipo As String
    Dim X As String

    If ValidarComprobanteSIIF = True Then
        With CargaComprobanteSIIF
            strComprobante = Format(.txtComprobante.Text, "00000") & "/" & Right(.txtFecha.Text, 2)
            datFecha = DateTime.DateSerial(Right(.txtFecha.Text, 4), Mid(.txtFecha.Text, 4, 2), Left(.txtFecha.Text, 2))
            Select Case .txtImputacion.Text
            Case "Honorarios"
                strTipo = "H"
            Case "Comisiones"
                strTipo = "C"
            Case "Horas Extras"
                strTipo = "E"
            Case "Licencia"
                strTipo = "L"
            End Select
            'Insertamos una copia del grupo de registros a editar cambiando el Nro. de Comprobante, Fecha y Tipo
            SQL = "Insert Into LIQUIDACIONHONORARIOS (Comprobante, Fecha, Tipo," _
            & " Proveedor, MontoBruto, Sellos, Seguro, LibramientoPago, IIBB, OtraRetencion, Anticipo, Descuento, Actividad, Partida)" _
            & " Select '" & strComprobante & "', #" & Format(datFecha, "MM/DD/YYYY") & "#, '" & strTipo & "'," _
            & " Proveedor, MontoBruto, Sellos, Seguro, LibramientoPago, IIBB, OtraRetencion, Anticipo, Descuento, Actividad, Partida" _
            & " From LIQUIDACIONHONORARIOS Where Comprobante = " & "'" & strEditandoAutocarga & "'"
            dbSlave.BeginTrans
            dbSlave.Execute SQL
            dbSlave.CommitTrans
            'Modificamos la copia insertada recientemente para incorporar la Actividad y Partida de cada Agente
            If strTipo <> "C" Then
                SQL = "Update LIQUIDACIONHONORARIOS Inner Join PRECARIZADOS" _
                & " On LIQUIDACIONHONORARIOS.Proveedor = PRECARIZADOS.Agentes" _
                & " Set LIQUIDACIONHONORARIOS.Actividad = PRECARIZADOS.Actividad," _
                & " LIQUIDACIONHONORARIOS.Partida = PRECARIZADOS.Partida" _
                & " Where Comprobante = " & "'" & strComprobante & "'"
            Else
                SQL = "Update LIQUIDACIONHONORARIOS Inner Join PRECARIZADOS" _
                & " On LIQUIDACIONHONORARIOS.Proveedor = PRECARIZADOS.Agentes" _
                & " Set LIQUIDACIONHONORARIOS.Actividad = PRECARIZADOS.Actividad," _
                & " LIQUIDACIONHONORARIOS.Partida = " & "'" & "399" & "'" _
                & " Where Comprobante = " & "'" & strComprobante & "'"
            End If
            dbSlave.BeginTrans
            dbSlave.Execute SQL
            dbSlave.CommitTrans
            'Borramos los registros que debían editarse
            SQL = "DELETE FROM LIQUIDACIONHONORARIOS Where COMPROBANTE = " & "'" & strEditandoAutocarga & "'"
            dbSlave.BeginTrans
            dbSlave.Execute SQL
            dbSlave.CommitTrans
            strEditandoAutocarga = ""
        End With
        MsgBox "Los datos fueron agregados"
        ListadoComprobantesSIIF.Show
        Unload CargaComprobanteSIIF
        X = strComprobante
        strComprobante = ""
        ListadoComprobantesSIIF.txtFecha.Text = "20" & Right(X, 2)
        ConfigurardgComprobantesSIIF
        Call CargardgComprobantesSIIF(, "20" & Right(X, 2), X)
        ConfigurardgImputacion
        X = ListadoComprobantesSIIF.dgListadoComprobante.Row
        CargardgImputacion (ListadoComprobantesSIIF.dgListadoComprobante.TextMatrix(X, 0))
        ConfigurardgRetencion
        CargardgRetencion (ListadoComprobantesSIIF.dgListadoComprobante.TextMatrix(X, 0))
        X = ""
        ListadoComprobantesSIIF.dgListadoComprobante.SetFocus
    End If
   
End Sub

Public Sub GenerarPrecarizado()

    Dim SQL As String

    If ValidarPrecarizado = True Then
        Set rstRegistroSlave = New ADODB.Recordset
        With CargaPrecarizado
            If strEditandoPrecarizado = "" Then
                rstRegistroSlave.Open "PRECARIZADOS", dbSlave, adOpenForwardOnly, adLockOptimistic
                rstRegistroSlave.AddNew
            Else
                SQL = "Select * from PRECARIZADOS Where AGENTES = " & "'" & strEditandoPrecarizado & "'"
                rstRegistroSlave.Open SQL, dbSlave, adOpenForwardOnly, adLockOptimistic
                SQL = ""
                strEditandoPrecarizado = ""
            End If
            rstRegistroSlave!AGENTES = .txtNombreCompleto.Text
            rstRegistroSlave!ACTIVIDAD = Left(.mskEstructura.Text, 8)
            rstRegistroSlave!PARTIDA = Right(.mskEstructura.Text, 3)
            rstRegistroSlave.Update
        End With
        rstRegistroSlave.Close
        Set rstRegistroSlave = Nothing
        Unload CargaPrecarizado
        ListadoPrecarizados.Show
        ConfigurardgPrecarizados
        CargardgPrecarizados
    End If
    
End Sub

Public Sub GenerarHaberLiquidado()

    Dim strPL As String
    Dim strCL As String
    Dim dblImporte As Double
    Dim SQL As String

    If ValidarCargaHaberLiquidado = True Then
        Set rstRegistroSlave = New ADODB.Recordset
        With CargaHaberLiquidado
            strPL = .txtPuestoLaboral.Text
            strCL = Left(.txtPeriodo.Text, 4)
            dblImporte = De_Txt_a_Num_01(.txtImporte.Text)
            rstRegistroSlave.Open "LIQUIDACIONSUELDOS", dbSlave, adOpenForwardOnly, adLockOptimistic
            rstRegistroSlave.AddNew
            rstRegistroSlave!PuestoLaboral = strPL
            rstRegistroSlave!CodigoLiquidacion = strCL
            rstRegistroSlave!CodigoConcepto = Left(.cmbConcepto.Text, 4)
            rstRegistroSlave!Importe = dblImporte
            rstRegistroSlave.Update
        End With
        rstRegistroSlave.Close
        'Ajustamos el Haber Óptimo
        SQL = "Select * from LIQUIDACIONSUELDOS " _
        & "Where PUESTOLABORAL = " & "'" & strPL & "'" _
        & "And CODIGOCONCEPTO = " & "'" & "9998" & "'" _
        & "And CODIGOLIQUIDACION = " & "'" & strCL & "'"
        If SQLNoMatch(SQL) Then
            rstRegistroSlave.Open "LIQUIDACIONSUELDOS", dbSlave, adOpenForwardOnly, adLockOptimistic
            rstRegistroSlave.AddNew
            rstRegistroSlave!PuestoLaboral = strPL
            rstRegistroSlave!CodigoConcepto = "9998"
            rstRegistroSlave!CodigoLiquidacion = strCL
            rstRegistroSlave!Importe = dblImporte
        Else
            rstRegistroSlave.Open SQL, dbSlave, adOpenForwardOnly, adLockOptimistic
            rstRegistroSlave!Importe = rstRegistroSlave!Importe + dblImporte
        End If
        rstRegistroSlave.Update
        rstRegistroSlave.Close
        'Ajustamos la Jubilación Personal
        SQL = "Select * from LIQUIDACIONSUELDOS " _
        & "Where PUESTOLABORAL = " & "'" & strPL & "'" _
        & "And CODIGOCONCEPTO = " & "'" & "0208" & "'" _
        & "And CODIGOLIQUIDACION = " & "'" & strCL & "'"
        If SQLNoMatch(SQL) Then
            rstRegistroSlave.Open "LIQUIDACIONSUELDOS", dbSlave, adOpenForwardOnly, adLockOptimistic
            rstRegistroSlave.AddNew
            rstRegistroSlave!PuestoLaboral = strPL
            rstRegistroSlave!CodigoConcepto = "0208"
            rstRegistroSlave!CodigoLiquidacion = strCL
            rstRegistroSlave!Importe = (dblImporte * 0.185)
        Else
            rstRegistroSlave.Open SQL, dbSlave, adOpenForwardOnly, adLockOptimistic
            rstRegistroSlave!Importe = rstRegistroSlave!Importe + (dblImporte * 0.185)
        End If
        rstRegistroSlave.Update
        rstRegistroSlave.Close
        'Ajustamos la Jubilación Estatal
        SQL = "Select * from LIQUIDACIONSUELDOS " _
        & "Where PUESTOLABORAL = " & "'" & strPL & "'" _
        & "And CODIGOCONCEPTO = " & "'" & "0209" & "'" _
        & "And CODIGOLIQUIDACION = " & "'" & strCL & "'"
        If SQLNoMatch(SQL) Then
            rstRegistroSlave.Open "LIQUIDACIONSUELDOS", dbSlave, adOpenForwardOnly, adLockOptimistic
            rstRegistroSlave.AddNew
            rstRegistroSlave!PuestoLaboral = strPL
            rstRegistroSlave!CodigoConcepto = "0209"
            rstRegistroSlave!CodigoLiquidacion = strCL
            rstRegistroSlave!Importe = (dblImporte * 0.185)
        Else
            rstRegistroSlave.Open SQL, dbSlave, adOpenForwardOnly, adLockOptimistic
            rstRegistroSlave!Importe = rstRegistroSlave!Importe + (dblImporte * 0.185)
        End If
        rstRegistroSlave.Update
        rstRegistroSlave.Close
        'Ajustamos la O.Social Personal
        SQL = "Select * from LIQUIDACIONSUELDOS " _
        & "Where PUESTOLABORAL = " & "'" & strPL & "'" _
        & "And CODIGOCONCEPTO = " & "'" & "0212" & "'" _
        & "And CODIGOLIQUIDACION = " & "'" & strCL & "'"
        If SQLNoMatch(SQL) Then
            rstRegistroSlave.Open "LIQUIDACIONSUELDOS", dbSlave, adOpenForwardOnly, adLockOptimistic
            rstRegistroSlave.AddNew
            rstRegistroSlave!PuestoLaboral = strPL
            rstRegistroSlave!CodigoConcepto = "0212"
            rstRegistroSlave!CodigoLiquidacion = strCL
            rstRegistroSlave!Importe = (dblImporte * 0.05)
        Else
            rstRegistroSlave.Open SQL, dbSlave, adOpenForwardOnly, adLockOptimistic
            rstRegistroSlave!Importe = rstRegistroSlave!Importe + (dblImporte * 0.05)
        End If
        rstRegistroSlave.Update
        rstRegistroSlave.Close
        'Ajustamos la O.Social Estatal
        SQL = "Select * from LIQUIDACIONSUELDOS " _
        & "Where PUESTOLABORAL = " & "'" & strPL & "'" _
        & "And CODIGOCONCEPTO = " & "'" & "0213" & "'" _
        & "And CODIGOLIQUIDACION = " & "'" & strCL & "'"
        If SQLNoMatch(SQL) Then
            rstRegistroSlave.Open "LIQUIDACIONSUELDOS", dbSlave, adOpenForwardOnly, adLockOptimistic
            rstRegistroSlave.AddNew
            rstRegistroSlave!PuestoLaboral = strPL
            rstRegistroSlave!CodigoConcepto = "0213"
            rstRegistroSlave!CodigoLiquidacion = strCL
            rstRegistroSlave!Importe = (dblImporte * 0.04)
        Else
            rstRegistroSlave.Open SQL, dbSlave, adOpenForwardOnly, adLockOptimistic
            rstRegistroSlave!Importe = rstRegistroSlave!Importe + (dblImporte * 0.04)
        End If
        rstRegistroSlave.Update
        rstRegistroSlave.Close
        
        Set rstRegistroSlave = Nothing
        ReciboDeSueldo.Show
        ReciboDeSueldo.cmbAgente.Text = CargaHaberLiquidado.txtNombreCompleto.Text
        ReciboDeSueldo.cmbPeriodo.Text = CargaHaberLiquidado.txtPeriodo.Text
        Unload CargaHaberLiquidado
        ConfigurardgHaberesLiquidados
        ConfigurardgDescuentosLiquidados
        Call CargardgHaberesLiquidados(strPL, strCL)
        Call CargardgDescuentosLiquidados(strPL, strCL)
    End If
    
    strPL = ""
    strCL = ""

End Sub

Public Sub GenerarPrecarizadoImputado()

    Dim SQL                 As String
    Dim strComprobante      As String
    Dim strAgente           As String
    Dim dblMontoBruto       As Double
    Dim varEditandoString   As Variant

    If ValidarPrecarizadoImputado = True Then
        Set rstRegistroSlave = New ADODB.Recordset
        With CargaPrecarizadoImputado
            If strEditandoPrecarizadoImputado = "" Then
                rstRegistroSlave.Open "LIQUIDACIONHONORARIOS", dbSlave, adOpenForwardOnly, adLockOptimistic
                rstRegistroSlave.AddNew
            Else
                varEditandoString = Split(strEditandoPrecarizadoImputado, "-")
                strComprobante = varEditandoString(0)
                strAgente = varEditandoString(1)
                dblMontoBruto = De_Txt_a_Num_01(varEditandoString(2))
                SQL = "Select * from LIQUIDACIONHONORARIOS " _
                & "Where COMPROBANTE = " & "'" & strComprobante & "' " _
                & "And PROVEEDOR = " & "'" & strAgente & "' " _
                & "And MONTOBRUTO = " & dblMontoBruto & ""
                rstRegistroSlave.Open SQL, dbSlave, adOpenForwardOnly, adLockOptimistic
                SQL = ""
                strEditandoPrecarizado = ""
                varEditandoString = ""
            End If
            rstRegistroSlave!ACTIVIDAD = Left(.mskEstructuraImputada.Text, 8)
            rstRegistroSlave!PARTIDA = Right(.mskEstructuraImputada.Text, 3)
            rstRegistroSlave!MontoBruto = De_Txt_a_Num_01(.txtMontoBruto.Text)
            rstRegistroSlave.Update
        End With
        rstRegistroSlave.Close
        Set rstRegistroSlave = Nothing
        With ListadoHonorariosImputados
            .Show
            .txtComprobante.Text = strComprobante
            .txtPeriodo.Text = CargaPrecarizadoImputado.txtPeriodo.Text
            ConfigurardgListadoHonorariosImputados
            Call CargardgListadoHonorariosImputados(strComprobante)
        End With
        Unload CargaPrecarizadoImputado
    End If
    
End Sub

Public Sub GenerarCodigoSIRADIG(Tabla As String)

    Dim SQL As String

    If ValidarCodigoSIRADIG(Tabla) = True Then
        Set rstRegistroSlave = New ADODB.Recordset
        With CargaCodigoSIRADIG
            If strEditandoCodigoSIRADIG = "" Then
                rstRegistroSlave.Open Tabla, dbSlave, adOpenForwardOnly, adLockOptimistic
                rstRegistroSlave.AddNew
            Else
                SQL = "Select * from " & Tabla & " Where CODIGO = " & "'" & strEditandoCodigoSIRADIG & "'"
                rstRegistroSlave.Open SQL, dbSlave, adOpenForwardOnly, adLockOptimistic
                SQL = ""
                strEditandoCodigoSIRADIG = ""
            End If
            rstRegistroSlave!Codigo = .txtCodigo.Text
            rstRegistroSlave!Denominacion = .txtDenominacion.Text
            rstRegistroSlave.Update
        End With
        rstRegistroSlave.Close
        Set rstRegistroSlave = Nothing
        strListadoCodigoSIRADIG = strCargaCodigoSIRADIG
        strCargaCodigoSIRADIG = ""
        Unload CargaCodigoSIRADIG
        ListadoCodigosSIRADIG.Show
        ConfigurardgCodigosSIRADIG
        Call CargardgCodigosSIRADIG(strListadoCodigoSIRADIG)
    End If

End Sub
