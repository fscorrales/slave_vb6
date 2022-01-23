Attribute VB_Name = "ConfigurarComboBox"
Public Sub CargarcmbFamiliar()
    
    Set rstListadoSlave = New ADODB.Recordset
    
    rstListadoSlave.Open "Select * From ASIGNACIONESFAMILIARES Order By PARENTESCO", dbSlave, adOpenDynamic, adLockOptimistic
    With CargaFamiliar.cmbParentesco
        If rstListadoSlave.BOF = False Then
            rstListadoSlave.MoveFirst
            While rstListadoSlave.EOF = False
                .AddItem rstListadoSlave!Parentesco
                rstListadoSlave.MoveNext
            Wend
        End If
    End With
    rstListadoSlave.Close
    Set rstListadoSlave = Nothing
    
    With CargaFamiliar.cmbNivelDeEstudio
        .AddItem "Sin Estudios"
        .AddItem "Primario"
        .AddItem "Secundario"
        .AddItem "Universitario"
    End With
    
End Sub

Public Sub CargarCmbImportacionLiquidacionSueldo()

    Set rstListadoSlave = New ADODB.Recordset
    
    rstListadoSlave.Open "Select * From CODIGOLIQUIDACIONES Order By CODIGO Desc", dbSlave, adOpenDynamic, adLockOptimistic
    With ImportacionLiquidacionSueldo.cmbCodigoLiquidacion
        If rstListadoSlave.BOF = False Then
            rstListadoSlave.MoveFirst
            While rstListadoSlave.EOF = False
                .AddItem rstListadoSlave!Codigo & "-" & rstListadoSlave!Descripcion
                rstListadoSlave.MoveNext
            Wend
        End If
    End With
    rstListadoSlave.Close
    Set rstListadoSlave = Nothing
    
End Sub

Public Sub CargarCmbIncorporarConceptoSueldo()

    Set rstListadoSlave = New ADODB.Recordset
    
    rstListadoSlave.Open "Select * From CONCEPTOS Order By CODIGO Asc", dbSlave, adOpenDynamic, adLockOptimistic
    With IncorporarConceptoSueldo.cmbConceptoLiquidacion
        If rstListadoSlave.BOF = False Then
            rstListadoSlave.MoveFirst
            While rstListadoSlave.EOF = False
                .AddItem rstListadoSlave!Codigo & "-" & rstListadoSlave!Denominacion
                rstListadoSlave.MoveNext
            Wend
        End If
    End With
    rstListadoSlave.Close
    
    rstListadoSlave.Open "Select * From CODIGOLIQUIDACIONES Order By CODIGO Desc", dbSlave, adOpenDynamic, adLockOptimistic
    With IncorporarConceptoSueldo.cmbCodigoLiquidacion
        If rstListadoSlave.BOF = False Then
            rstListadoSlave.MoveFirst
            While rstListadoSlave.EOF = False
                .AddItem rstListadoSlave!Codigo & "-" & rstListadoSlave!Descripcion
                rstListadoSlave.MoveNext
            Wend
        End If
    End With
    rstListadoSlave.Close
    Set rstListadoSlave = Nothing
    
End Sub

Public Sub CargarCmbCopiarLiquidacionSisper()

    Dim strCodigoLiquidacion As String
    
    Set rstListadoSlave = New ADODB.Recordset
    
     
    rstListadoSlave.Open "Select * From CODIGOLIQUIDACIONES Order By CODIGO Desc", dbSlave, adOpenDynamic, adLockOptimistic
    With CopiarLiquidacionSisper
        If rstListadoSlave.BOF = False Then
            rstListadoSlave.MoveFirst
            While rstListadoSlave.EOF = False
                strCodigoLiquidacion = rstListadoSlave!Codigo & "-" & rstListadoSlave!Descripcion
                .cmbCodigoLiquidacionOrigen.AddItem strCodigoLiquidacion
                .cmbCodigoLiquidacionDestino.AddItem strCodigoLiquidacion
                rstListadoSlave.MoveNext
            Wend
        End If
    End With
    rstListadoSlave.Close
    Set rstListadoSlave = Nothing
    
End Sub

Public Sub CargarCmbLiquidacionPruebaSisper()

    Dim strCodigoLiquidacion As String
    
    Set rstListadoSlave = New ADODB.Recordset
    
     
    rstListadoSlave.Open "Select * From CODIGOLIQUIDACIONES Order By CODIGO Desc", dbSlave, adOpenDynamic, adLockOptimistic
    With LiquidacionPruebaSISPER
        If rstListadoSlave.BOF = False Then
            rstListadoSlave.MoveFirst
            While rstListadoSlave.EOF = False
                strCodigoLiquidacion = rstListadoSlave!Codigo & "-" & rstListadoSlave!Descripcion
                .cmbCodigoLiquidacionBase.AddItem strCodigoLiquidacion
                .cmbCodigoLiquidacionDestino.AddItem strCodigoLiquidacion
                rstListadoSlave.MoveNext
            Wend
        End If
    End With
    rstListadoSlave.Close
    
    rstListadoSlave.Open "Select * From CONCEPTOS Order By CODIGO Asc", dbSlave, adOpenDynamic, adLockOptimistic
    With LiquidacionPruebaSISPER.cmbConceptoLiquidacion
        If rstListadoSlave.BOF = False Then
            rstListadoSlave.MoveFirst
            While rstListadoSlave.EOF = False
                .AddItem rstListadoSlave!Codigo & "-" & rstListadoSlave!Denominacion
                rstListadoSlave.MoveNext
            Wend
        End If
    End With
    rstListadoSlave.Close
    
    Set rstListadoSlave = Nothing
    
End Sub

Public Sub CargarcmbComprobanteSIIF()
    
    With CargaComprobanteSIIF.txtImputacion
        .AddItem "Honorarios"
        .AddItem "Comisiones"
        .AddItem "Horas Extras"
        .AddItem "Licencia"
    End With
    
End Sub

Public Sub CargarCmbPeriodoResumenAnualGanancias()

    Dim SQL As String
    
    SQL = "Select Right(PERIODO,4) As PeriodoLiquidacion From" _
    & " LIQUIDACIONGANANCIAS4TACATEGORIA Inner Join CODIGOLIQUIDACIONES On" _
    & " LIQUIDACIONGANANCIAS4TACATEGORIA.CodigoLiquidacion = CODIGOLIQUIDACIONES.Codigo" _
    & " Group by Right(PERIODO,4)" _
    & " Order by Right(PERIODO,4) Desc"
    Set rstListadoSlave = New ADODB.Recordset
    rstListadoSlave.Open SQL, dbSlave, adOpenDynamic, adLockOptimistic
    With ResumenAnualGanancias.cmbPeriodo
        If rstListadoSlave.BOF = False Then
            rstListadoSlave.MoveFirst
            While rstListadoSlave.EOF = False
                .AddItem rstListadoSlave!PeriodoLiquidacion
                rstListadoSlave.MoveNext
            Wend
            rstListadoSlave.MoveFirst
            .Text = rstListadoSlave!PeriodoLiquidacion
        End If
    End With
    rstListadoSlave.Close
    Set rstListadoSlave = Nothing
    
End Sub

Public Sub CargarCmbAgenteResumenAnualGanancias(PeriodoLiquidacion As String)

    Dim SQL As String
    
    SQL = "Select NombreCompleto From" _
    & " (LIQUIDACIONGANANCIAS4TACATEGORIA Inner Join CODIGOLIQUIDACIONES On" _
    & " LIQUIDACIONGANANCIAS4TACATEGORIA.CodigoLiquidacion = CODIGOLIQUIDACIONES.Codigo)" _
    & " Inner Join AGENTES On" _
    & " LIQUIDACIONGANANCIAS4TACATEGORIA.PuestoLaboral = AGENTES.PuestoLaboral" _
    & " Where Right(PERIODO,4) = '" & PeriodoLiquidacion & "'" _
    & " Group by NombreCompleto" _
    & " Order by NombreCompleto Asc"
    Set rstListadoSlave = New ADODB.Recordset
    rstListadoSlave.Open SQL, dbSlave, adOpenDynamic, adLockOptimistic
    With ResumenAnualGanancias.cmbAgente
        If rstListadoSlave.BOF = False Then
            .Clear
            rstListadoSlave.MoveFirst
            While rstListadoSlave.EOF = False
                .AddItem rstListadoSlave!NombreCompleto
                rstListadoSlave.MoveNext
            Wend
        End If
    End With
    rstListadoSlave.Close
    Set rstListadoSlave = Nothing
    
End Sub

Public Sub CargarCmbReciboDeSueldo()

    Set rstListadoSlave = New ADODB.Recordset
    
    rstListadoSlave.Open "Select * From CODIGOLIQUIDACIONES Order By CODIGO Desc", dbSlave, adOpenDynamic, adLockOptimistic
    With ReciboDeSueldo.cmbPeriodo
        If rstListadoSlave.BOF = False Then
            rstListadoSlave.MoveFirst
            While rstListadoSlave.EOF = False
                .AddItem rstListadoSlave!Codigo & "-" & rstListadoSlave!Descripcion
                rstListadoSlave.MoveNext
            Wend
        End If
    End With
    rstListadoSlave.Close
    
    rstListadoSlave.Open "Select NombreCompleto From AGENTES Order By NombreCompleto Asc", dbSlave, adOpenDynamic, adLockOptimistic
    With ReciboDeSueldo.cmbAgente
        If rstListadoSlave.BOF = False Then
            rstListadoSlave.MoveFirst
            While rstListadoSlave.EOF = False
                .AddItem rstListadoSlave!NombreCompleto
                rstListadoSlave.MoveNext
            Wend
        End If
    End With
    rstListadoSlave.Close
    
    Set rstListadoSlave = Nothing
    
End Sub

Public Sub CargarCmbCargaHaberLiquidado()

    Set rstListadoSlave = New ADODB.Recordset
    
    rstListadoSlave.Open "Select * From CONCEPTOS Order By CODIGO Asc", dbSlave, adOpenDynamic, adLockOptimistic
    With CargaHaberLiquidado.cmbConcepto
        If rstListadoSlave.BOF = False Then
            rstListadoSlave.MoveFirst
            While rstListadoSlave.EOF = False
                .AddItem rstListadoSlave!Codigo & "-" & rstListadoSlave!Denominacion
                rstListadoSlave.MoveNext
            Wend
        End If
    End With
    rstListadoSlave.Close

    Set rstListadoSlave = Nothing
    
End Sub

Public Sub CargarCmbAgenteListadoF572()

    Dim SQL As String
    
    SQL = "Select NombreCompleto From " _
        & "PresentacionSIRADIG Inner Join Agentes On " _
        & "PresentacionSIRADIG.CUIL = AGENTES.CUIL " _
        & "Group by NombreCompleto " _
        & "Order by NombreCompleto Asc"
    Set rstListadoSlave = New ADODB.Recordset
    rstListadoSlave.Open SQL, dbSlave, adOpenDynamic, adLockOptimistic
    With ListadoF572.cmbAgente
        If rstListadoSlave.BOF = False Then
            .Clear
            rstListadoSlave.MoveFirst
            While rstListadoSlave.EOF = False
                .AddItem rstListadoSlave!NombreCompleto
                rstListadoSlave.MoveNext
            Wend
        End If
    End With
    rstListadoSlave.Close
    Set rstListadoSlave = Nothing
    
End Sub


Public Sub CargarCmbMigrarDeducciones()

    Dim SQL As String
    Dim strPeriodoDDJJ As String
    Dim strPeriodoActual As String
    
    strPeriodoActual = Year(Now())
    SQL = "Select Right(ID,2) as Periodo " _
        & "From PresentacionSIRADIG " _
        & "Group By Right(ID,2) " _
        & "Order By Right(ID,2) Asc"
        
    Set rstListadoSlave = New ADODB.Recordset
    rstListadoSlave.Open SQL, dbSlave, adOpenDynamic, adLockOptimistic
    With MigrarDeducciones
        .cmbPeriodoDDJJOrigen.AddItem "BD Previa"
        .cmbPeriodoDDJJDestino.AddItem strPeriodoActual
        If rstListadoSlave.BOF = False Then
            rstListadoSlave.MoveFirst
            While rstListadoSlave.EOF = False
                strPeriodoDDJJ = "20" & rstListadoSlave!Periodo
                .cmbPeriodoDDJJOrigen.AddItem strPeriodoDDJJ
                If strPeriodoActual <> strPeriodoDDJJ Then
                    .cmbPeriodoDDJJDestino.AddItem strPeriodoDDJJ
                End If
                rstListadoSlave.MoveNext
            Wend
        End If
        .cmbTipoDatos.AddItem "Todas las Deducciones"
        .cmbTipoDatos.AddItem "Solo Deducciones Personales"
        .cmbTipoDatos.AddItem "Solo Deducciones Generales"
    End With
    rstListadoSlave.Close
    Set rstListadoSlave = Nothing
    
End Sub

Public Sub CargarCmbLiquidacionGanancia4taSIRADIG(PuestoLaboral As String, _
CodigoLiquidacion As String)

    Dim SQL As String
    Dim strPeriodoLiquidacion As String
    
    strPeriodoLiquidacion = BuscarPeriodoLiquidacion(CodigoLiquidacion)
    
    SQL = "Select * From " _
        & "PresentacionSIRADIG Inner Join Agentes On " _
        & "PresentacionSIRADIG.CUIL = AGENTES.CUIL " _
        & "Where Agentes.PuestoLaboral = '" & PuestoLaboral & "' " _
        & "And Right(PresentacionSIRADIG.ID,2) = '" & Right(strPeriodoLiquidacion, 2) & "' " _
        & "And Month(PresentacionSIRADIG.Fecha) <= " & Left(strPeriodoLiquidacion, 2) & " " _
        & "Order By PresentacionSIRADIG.Fecha Asc, PresentacionSIRADIG.ID Asc"
    Set rstListadoSlave = New ADODB.Recordset
    rstListadoSlave.Open SQL, dbSlave, adOpenDynamic, adLockOptimistic
    With LiquidacionGanancia4taSIRADIG.cmbPresentacionesSIRADIG
        .Clear
        .AddItem "Ninguno"
        If rstListadoSlave.BOF = False Then
            '.Clear
            rstListadoSlave.MoveFirst
            While rstListadoSlave.EOF = False
                strPeriodoLiquidacion = rstListadoSlave!NroPresentacion & "° -> " & _
                rstListadoSlave!Fecha & "-(" & rstListadoSlave!ID & ")"
                .AddItem strPeriodoLiquidacion
                rstListadoSlave.MoveNext
            Wend
            .Text = strPeriodoLiquidacion
        Else
        .Text = "Ninguno"
        End If
    End With
    rstListadoSlave.Close
    Set rstListadoSlave = Nothing
    
End Sub
