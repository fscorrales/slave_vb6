Attribute VB_Name = "ConfigurarDataGrid"
Public Sub ConfigurardgAgentes(FormularioOrigen As Form)
    
    With FormularioOrigen.dgAgentes
        .Clear
        .Cols = 5
        .Rows = 2
        .TextMatrix(0, 0) = "Activado"
        .ColWidth(0) = 750
        .TextMatrix(0, 1) = "Nombre Completo"
        .ColWidth(1) = 3000
        .TextMatrix(0, 2) = "Puesto"
        .ColWidth(2) = 750
        .TextMatrix(0, 3) = "Legajo"
        .ColWidth(3) = 600
        .TextMatrix(0, 4) = "CUIL / DNI"
        .ColWidth(4) = 1150
        .FixedCols = 0
        .FocusRect = flexFocusHeavy
        .HighLight = flexHighlightWithFocus
        .SelectionMode = flexSelectionByRow
        .AllowUserResizing = flexResizeColumns
        .ColAlignment(0) = 4
        .ColAlignment(1) = 1
        .ColAlignment(2) = 7
        .ColAlignment(3) = 7
        .ColAlignment(4) = 7
    End With
    
End Sub

Public Sub CargardgAgentes(FormularioOrigen As Form)
    
    Dim i As Integer
    i = 0
    FormularioOrigen.dgAgentes.Rows = 2
    Set rstListadoSlave = New ADODB.Recordset
    rstListadoSlave.Open "Select * From AGENTES Order By NombreCompleto", dbSlave, adOpenDynamic, adLockOptimistic
    With FormularioOrigen.dgAgentes
        If rstListadoSlave.BOF = False Then
            rstListadoSlave.MoveFirst
            While rstListadoSlave.EOF = False
                i = i + 1
                .RowHeight(i) = 300
                If rstListadoSlave!Activado Then
                    .TextMatrix(i, 0) = "SI"
                Else
                    .TextMatrix(i, 0) = "NO"
                End If
                .TextMatrix(i, 1) = rstListadoSlave!NombreCompleto
                .TextMatrix(i, 2) = rstListadoSlave!PuestoLaboral
                .TextMatrix(i, 3) = rstListadoSlave!Legajo
                .TextMatrix(i, 4) = rstListadoSlave!CUIL
                rstListadoSlave.MoveNext
                .Rows = .Rows + 1
            Wend
        End If
        .Rows = .Rows - 1
        .SetFocus
    End With
    rstListadoSlave.Close
    Set rstListadoSlave = Nothing
    
End Sub

Public Sub ConfigurardgPrecarizados()
    
    With ListadoPrecarizados.dgPrecarizados
        .Clear
        .Cols = 2
        .Rows = 2
        .TextMatrix(0, 0) = "Nombre Completo"
        .ColWidth(0) = 3000
        .TextMatrix(0, 1) = "Prog-Proy-Act-Part"
        .ColWidth(1) = 1500
        .FixedCols = 1
        .FocusRect = flexFocusHeavy
        .HighLight = flexHighlightWithFocus
        .SelectionMode = flexSelectionByRow
        .AllowUserResizing = flexResizeColumns
        .ColAlignment(0) = 4
        .ColAlignment(1) = 7
    End With
    
End Sub

Public Sub CargardgPrecarizados()
    
    Dim i As Integer
    Dim strEstructuraPresupuestaria As String
    
    i = 0
    ListadoPrecarizados.dgPrecarizados.Rows = 2
    Set rstListadoSlave = New ADODB.Recordset
    rstListadoSlave.Open "Select * From Precarizados Order By Agentes", dbSlave, adOpenDynamic, adLockOptimistic
    With ListadoPrecarizados.dgPrecarizados
        If rstListadoSlave.BOF = False Then
            rstListadoSlave.MoveFirst
            While rstListadoSlave.EOF = False
                i = i + 1
                .RowHeight(i) = 300
                .TextMatrix(i, 0) = rstListadoSlave!AGENTES
                strEstructuraPresupuestaria = rstListadoSlave!ACTIVIDAD & "-" & rstListadoSlave!PARTIDA
                If ValidarEstructuraPresupuestaria(strEstructuraPresupuestaria) = False Then
                    .TextMatrix(i, 1) = strEstructuraPresupuestaria
                    .Row = i
                    .CellBackColor = &H80FF&
                Else
                    .TextMatrix(i, 1) = strEstructuraPresupuestaria
                End If
                rstListadoSlave.MoveNext
                .Rows = .Rows + 1
            Wend
        End If
        .Rows = .Rows - 1
        .Row = 1
        .SetFocus
    End With
    rstListadoSlave.Close
    Set rstListadoSlave = Nothing
    
End Sub


Public Sub ConfigurardgConceptos()
    
    With ListadoConceptos.dgConceptos
        .Clear
        .Cols = 2
        .Rows = 2
        .TextMatrix(0, 0) = "Código"
        .TextMatrix(0, 1) = "Denominación"
        .ColWidth(0) = 600
        .ColWidth(1) = 3400
        .FixedCols = 0
        .FocusRect = flexFocusHeavy
        .HighLight = flexHighlightWithFocus
        .SelectionMode = flexSelectionByRow
        .AllowUserResizing = flexResizeColumns
        .ColAlignment(0) = 4
        .ColAlignment(1) = 1
    End With
    
End Sub

Public Sub CargardgConceptos()
    
    Dim i As Integer
    i = 0
    ListadoConceptos.dgConceptos.Rows = 2
    Set rstListadoSlave = New ADODB.Recordset
    rstListadoSlave.Open "Select * From CONCEPTOS Order By CODIGO", dbSlave, adOpenDynamic, adLockOptimistic
    With ListadoConceptos.dgConceptos
        If rstListadoSlave.BOF = False Then
            rstListadoSlave.MoveFirst
            While rstListadoSlave.EOF = False
                i = i + 1
                .RowHeight(i) = 300
                .TextMatrix(i, 0) = rstListadoSlave!Codigo
                .TextMatrix(i, 1) = rstListadoSlave!Denominacion
                rstListadoSlave.MoveNext
                .Rows = .Rows + 1
            Wend
        End If
        .Rows = .Rows - 1
        .SetFocus
    End With
    rstListadoSlave.Close
    Set rstListadoSlave = Nothing
    
End Sub

Public Sub ConfigurardgNormasEscalaGanancias()
    
    With ListadoEscalaGanancias.dgNormasEscalaGanancias
        .Clear
        .Cols = 2
        .Rows = 2
        .TextMatrix(0, 0) = "Norma"
        .TextMatrix(0, 1) = "Fecha"
        .ColWidth(0) = 1300
        .ColWidth(1) = 1000
        .FixedCols = 0
        .FocusRect = flexFocusHeavy
        .HighLight = flexHighlightWithFocus
        .SelectionMode = flexSelectionByRow
        .AllowUserResizing = flexResizeColumns
        .ColAlignment(0) = 1
        .ColAlignment(1) = 7
    End With
    
End Sub

Public Sub CargardgNormasEscalaGanancias()
    
    Dim SQL As String
    Dim i As Integer
    i = 0
    
    ListadoEscalaGanancias.dgNormasEscalaGanancias.Rows = 2
    SQL = "Select NORMALEGAL, FECHA From ESCALAAPLICABLEGANANCIAS" _
    & " Group by NORMALEGAL, FECHA" _
    & " Order By FECHA Desc"
    Set rstListadoSlave = New ADODB.Recordset
    rstListadoSlave.Open SQL, dbSlave, adOpenDynamic, adLockOptimistic
    With ListadoEscalaGanancias.dgNormasEscalaGanancias
        If rstListadoSlave.BOF = False Then
            rstListadoSlave.MoveFirst
            While rstListadoSlave.EOF = False
                i = i + 1
                .RowHeight(i) = 300
                .TextMatrix(i, 0) = rstListadoSlave!NormaLegal
                .TextMatrix(i, 1) = rstListadoSlave!Fecha
                rstListadoSlave.MoveNext
                .Rows = .Rows + 1
            Wend
        End If
        .Rows = .Rows - 1
        .SetFocus
    End With
    rstListadoSlave.Close
    Set rstListadoSlave = Nothing
    
End Sub

Public Sub ConfigurardgEscalaGanancias()
    
    With ListadoEscalaGanancias.dgEscalaGanancias
        .Clear
        .Cols = 6
        .Rows = 3
        .MergeRow(0) = True
        .MergeCol(0) = True
        .MergeCol(1) = True
        .MergeCells = flexMergeRestrictAll
        .TextMatrix(0, 0) = "Tr"
        .TextMatrix(1, 0) = "Tr"
        .ColWidth(0) = 300
        .TextMatrix(0, 1) = "Gcia.Neta Acum."
        .TextMatrix(1, 1) = "Más de $"
        .ColWidth(1) = 850
        .TextMatrix(0, 2) = "Gcia.Neta Acum."
        .TextMatrix(1, 2) = "A $"
        .ColWidth(2) = 850
        .TextMatrix(0, 3) = "Pagarán"
        .TextMatrix(1, 3) = "$ Fijo"
        .ColWidth(3) = 850
        .TextMatrix(0, 4) = "Pagarán"
        .TextMatrix(1, 4) = "Más el %"
        .ColWidth(4) = 850
        .TextMatrix(0, 5) = "Pagarán"
        .TextMatrix(1, 5) = "Sobre $"
        .ColWidth(5) = 850
        .FixedCols = 0
        .FixedRows = 2
        .FocusRect = flexFocusHeavy
        .HighLight = flexHighlightWithFocus
        .SelectionMode = flexSelectionByRow
        .AllowUserResizing = flexResizeColumns
        .ColAlignmentFixed(0) = 1
        .ColAlignmentFixed(1) = 4
        .ColAlignmentFixed(2) = 4
        .ColAlignmentFixed(3) = 4
        .ColAlignmentFixed(4) = 4
        .ColAlignmentFixed(5) = 4
        .ColAlignment(0) = 7
        .ColAlignment(1) = 7
        .ColAlignment(2) = 7
        .ColAlignment(3) = 7
        .ColAlignment(4) = 7
        .ColAlignment(5) = 7
    End With
    
End Sub

Public Sub CargardgEscalaGanancias(Norma As String)
    
    Dim i As Integer
    Dim curImporte As Currency
    Dim SQL As String
    
    i = 1
    curImporte = 0
    ListadoEscalaGanancias.dgEscalaGanancias.Rows = 3
    SQL = "Select * From ESCALAAPLICABLEGANANCIAS Where NORMALEGAL = '" _
    & Norma & "' Order By TRAMO"
    Set rstListadoSlave = New ADODB.Recordset
    rstListadoSlave.Open SQL, dbSlave, adOpenDynamic, adLockOptimistic
    With ListadoEscalaGanancias.dgEscalaGanancias
        If rstListadoSlave.BOF = False Then
            rstListadoSlave.MoveFirst
            While rstListadoSlave.EOF = False
                i = i + 1
                .RowHeight(i) = 300
                .TextMatrix(i, 0) = rstListadoSlave!Tramo
                .TextMatrix(i, 1) = De_Num_a_Tx_01(curImporte, True)
                .TextMatrix(i, 2) = De_Num_a_Tx_01(rstListadoSlave!ImporteMaximo, True)
                .TextMatrix(i, 3) = De_Num_a_Tx_01(rstListadoSlave!ImporteFijo, False)
                .TextMatrix(i, 4) = rstListadoSlave!ImporteVariable * 100 & " %"
                .TextMatrix(i, 5) = De_Num_a_Tx_01(curImporte, True)
                curImporte = De_Txt_a_Num_01(.TextMatrix(i, 2))
                'currimporte = rstListadoSlave!ImporteMaximo
                rstListadoSlave.MoveNext
                .Rows = .Rows + 1
            Wend
        End If
        .Rows = .Rows - 1
        .SetFocus
    End With
    rstListadoSlave.Close
    Set rstListadoSlave = Nothing
    
End Sub

Public Sub ConfigurardgLimitesDeducciones()
    
    With ListadoLimitesDeducciones.dgLimitesDeducciones
        .Clear
        .Cols = 13
        .Rows = 3
        .MergeRow(0) = True
        .MergeCol(0) = True
        .MergeCol(1) = True
        .MergeCells = flexMergeRestrictAll
        .TextMatrix(0, 0) = "Norma Legal"
        .TextMatrix(1, 0) = "Número"
        .ColWidth(0) = 1000
        .TextMatrix(0, 1) = "Norma Legal"
        .TextMatrix(1, 1) = "Fecha"
        .ColWidth(1) = 1000
        .TextMatrix(0, 2) = "Deducciones Personales"
        .TextMatrix(1, 2) = "Minimo"
        .ColWidth(2) = 1250
        .TextMatrix(0, 3) = "Deducciones Personales"
        .TextMatrix(1, 3) = "Especial"
        .ColWidth(3) = 1250
        .TextMatrix(0, 4) = "Deducciones Personales"
        .TextMatrix(1, 4) = "Hijo"
        .ColWidth(4) = 1250
        .TextMatrix(0, 5) = "Deducciones Personales"
        .TextMatrix(1, 5) = "Conyuge"
        .ColWidth(5) = 1250
        .TextMatrix(0, 6) = "Deducciones Personales"
        .TextMatrix(1, 6) = "Otros"
        .ColWidth(6) = 1250
        .TextMatrix(0, 7) = "Deducciones Generales"
        .TextMatrix(1, 7) = "Serv.Doméstico"
        .ColWidth(7) = 1250
        .TextMatrix(0, 8) = "Deducciones Generales"
        .TextMatrix(1, 8) = "Seguro Vida"
        .ColWidth(8) = 1250
        .TextMatrix(0, 9) = "Deducciones Generales"
        .TextMatrix(1, 9) = "Alquileres"
        .ColWidth(9) = 1250
        .TextMatrix(0, 10) = "Deducciones Generales"
        .TextMatrix(1, 10) = "H. Médicos"
        .ColWidth(10) = 1250
        .TextMatrix(0, 11) = "Deducciones Generales"
        .TextMatrix(1, 11) = "Cuota Médica"
        .ColWidth(11) = 1250
        .TextMatrix(0, 12) = "Deducciones Generales"
        .TextMatrix(1, 12) = "Donaciones"
        .ColWidth(12) = 1250
        .FixedCols = 0
        .FixedRows = 2
        .FocusRect = flexFocusHeavy
        .HighLight = flexHighlightWithFocus
        .SelectionMode = flexSelectionByRow
        .AllowUserResizing = flexResizeColumns
        .ColAlignmentFixed(0) = 4
        .ColAlignmentFixed(1) = 4
        .ColAlignmentFixed(2) = 4
        .ColAlignmentFixed(3) = 4
        .ColAlignmentFixed(4) = 4
        .ColAlignmentFixed(5) = 4
        .ColAlignmentFixed(6) = 4
        .ColAlignmentFixed(7) = 4
        .ColAlignmentFixed(8) = 4
        .ColAlignmentFixed(9) = 4
        .ColAlignmentFixed(10) = 4
        .ColAlignmentFixed(11) = 4
        .ColAlignmentFixed(12) = 4
        .ColAlignment(0) = 7
        .ColAlignment(1) = 7
        .ColAlignment(2) = 7
        .ColAlignment(3) = 7
        .ColAlignment(4) = 7
        .ColAlignment(5) = 7
        .ColAlignment(6) = 7
        .ColAlignment(7) = 7
        .ColAlignment(8) = 7
        .ColAlignment(9) = 7
        .ColAlignment(10) = 7
        .ColAlignment(11) = 7
        .ColAlignment(12) = 7
    End With
    
End Sub

Public Sub CargardgLimitesDeducciones()
    
    Dim i As Integer
    
    i = 1
    ListadoLimitesDeducciones.dgLimitesDeducciones.Rows = 3
    Set rstListadoSlave = New ADODB.Recordset
    rstListadoSlave.Open "Select * From DEDUCCIONES4TACATEGORIA Order By FECHA Desc", dbSlave, adOpenDynamic, adLockOptimistic
    With ListadoLimitesDeducciones.dgLimitesDeducciones
        If rstListadoSlave.BOF = False Then
            rstListadoSlave.MoveFirst
            While rstListadoSlave.EOF = False
                i = i + 1
                .RowHeight(i) = 300
                .TextMatrix(i, 0) = rstListadoSlave!NormaLegal
                .TextMatrix(i, 1) = rstListadoSlave!Fecha
                .TextMatrix(i, 2) = De_Num_a_Tx_01(rstListadoSlave!MinimoNoImponible)
                .TextMatrix(i, 3) = De_Num_a_Tx_01(rstListadoSlave!DeduccionEspecial)
                .TextMatrix(i, 4) = De_Num_a_Tx_01(rstListadoSlave!Hijo)
                .TextMatrix(i, 5) = De_Num_a_Tx_01(rstListadoSlave!Conyuge)
                .TextMatrix(i, 6) = De_Num_a_Tx_01(rstListadoSlave!OtrasCargasDeFamilia)
                .TextMatrix(i, 7) = De_Num_a_Tx_01(rstListadoSlave!ServicioDomestico)
                .TextMatrix(i, 8) = De_Num_a_Tx_01(rstListadoSlave!SeguroDeVida)
                .TextMatrix(i, 9) = De_Num_a_Tx_01(rstListadoSlave!Alquileres)
                .TextMatrix(i, 10) = rstListadoSlave!HonorariosMedicos * 100 & " %"
                .TextMatrix(i, 11) = rstListadoSlave!CuotaMedicoAsistencial * 100 & " %"
                .TextMatrix(i, 12) = rstListadoSlave!Donaciones * 100 & " %"
                rstListadoSlave.MoveNext
                .Rows = .Rows + 1
            Wend
        End If
        .Rows = .Rows - 1
        .SetFocus
    End With
    rstListadoSlave.Close
    Set rstListadoSlave = Nothing
    
End Sub

Public Sub ConfigurardgParentesco()
    
    With ListadoParentesco.dgParentesco
        .Clear
        .Cols = 3
        .Rows = 2
        .TextMatrix(0, 0) = "Código"
        .TextMatrix(0, 1) = "Denominación"
        .TextMatrix(0, 2) = "Importe"
        .ColWidth(0) = 600
        .ColWidth(1) = 2650
        .ColWidth(2) = 750
        .FixedCols = 0
        .FocusRect = flexFocusHeavy
        .HighLight = flexHighlightWithFocus
        .SelectionMode = flexSelectionByRow
        .AllowUserResizing = flexResizeColumns
        .ColAlignment(0) = 4
        .ColAlignment(1) = 1
        .ColAlignment(2) = 1
    
    End With
End Sub

Public Sub CargardgParentesco()
    
    Dim i As Integer
    i = 0
    ListadoParentesco.dgParentesco.Rows = 2
    Set rstListadoSlave = New ADODB.Recordset
    rstListadoSlave.Open "Select * From ASIGNACIONESFAMILIARES Order By CODIGO", dbSlave, adOpenDynamic, adLockOptimistic
    With ListadoParentesco.dgParentesco
        If rstListadoSlave.BOF = False Then
            rstListadoSlave.MoveFirst
            While rstListadoSlave.EOF = False
                i = i + 1
                .RowHeight(i) = 300
                .TextMatrix(i, 0) = rstListadoSlave!Codigo
                .TextMatrix(i, 1) = rstListadoSlave!Parentesco
                .TextMatrix(i, 2) = De_Num_a_Tx_01(rstListadoSlave!Importe)
                rstListadoSlave.MoveNext
                .Rows = .Rows + 1
            Wend
        End If
        .Rows = .Rows - 1
        .SetFocus
    End With
    rstListadoSlave.Close
    Set rstListadoSlave = Nothing
    
End Sub

Public Sub ConfigurardgDeduccionesGenerales()
    
    With ListadoDeduccionesGenerales.dgDeduccionesGenerales
        .Clear
        .Cols = 7
        .Rows = 2
        .TextMatrix(0, 0) = "Fecha"
        .ColWidth(0) = 1000
        .TextMatrix(0, 1) = "Serv.Doméstico"
        .ColWidth(1) = 1200
        .TextMatrix(0, 2) = "Seguro Vida"
        .ColWidth(2) = 1200
        .TextMatrix(0, 3) = "Alquiler"
        .ColWidth(3) = 1200
        .TextMatrix(0, 4) = "Cuota Médico"
        .ColWidth(4) = 1200
        .TextMatrix(0, 5) = "Donaciones"
        .ColWidth(5) = 1200
        .TextMatrix(0, 6) = "H.Médicos"
        .ColWidth(6) = 1200
        .FixedCols = 0
        .FocusRect = flexFocusHeavy
        .HighLight = flexHighlightWithFocus
        .SelectionMode = flexSelectionByRow
        .AllowUserResizing = flexResizeColumns
        .ColAlignment(0) = 4
        .ColAlignment(1) = 7
        .ColAlignment(2) = 7
        .ColAlignment(3) = 7
        .ColAlignment(4) = 7
        .ColAlignment(5) = 7
        .ColAlignment(6) = 7
    End With
    
End Sub

Public Sub CargardgDeduccionesGenerales(PuestoLaboral As String)
    
    Dim i As Integer
    i = 0
    ListadoDeduccionesGenerales.dgDeduccionesGenerales.Rows = 2
    Set rstListadoSlave = New ADODB.Recordset
    rstListadoSlave.Open "Select * From IMPORTEDEDUCCIONESGENERALES Where PUESTOLABORAL = " & "'" & PuestoLaboral & "' Order By Fecha Desc", dbSlave, adOpenDynamic, adLockOptimistic
    With ListadoDeduccionesGenerales.dgDeduccionesGenerales
        If rstListadoSlave.BOF = False Then
            rstListadoSlave.MoveFirst
            While rstListadoSlave.EOF = False
                i = i + 1
                .RowHeight(i) = 300
                .TextMatrix(i, 0) = rstListadoSlave!Fecha
                .TextMatrix(i, 1) = De_Num_a_Tx_01(rstListadoSlave!ServicioDomestico)
                .TextMatrix(i, 2) = De_Num_a_Tx_01(rstListadoSlave!SeguroDeVida)
                .TextMatrix(i, 3) = De_Num_a_Tx_01(rstListadoSlave!Alquileres)
                .TextMatrix(i, 4) = De_Num_a_Tx_01(rstListadoSlave!CuotaMedicoAsistencial)
                .TextMatrix(i, 5) = De_Num_a_Tx_01(rstListadoSlave!Donaciones)
                .TextMatrix(i, 6) = De_Num_a_Tx_01(rstListadoSlave!HonorariosMedicos)
                rstListadoSlave.MoveNext
                .Rows = .Rows + 1
            Wend
        End If
        .Rows = .Rows - 1
    End With
    rstListadoSlave.Close
    Set rstListadoSlave = Nothing
    
End Sub

Public Sub ConfigurardgDeduccionesPersonales()
    
    With ListadoDeduccionesPersonales.dgDeduccionesPersonales
        .Clear
        .Cols = 3
        .Rows = 2
        .TextMatrix(0, 0) = "Fecha"
        .ColWidth(0) = 1000
        .TextMatrix(0, 1) = "Nombre Completo"
        .ColWidth(1) = 2750
        .TextMatrix(0, 2) = "Parentesco"
        .ColWidth(2) = 2000
        .FixedCols = 0
        .FocusRect = flexFocusHeavy
        .HighLight = flexHighlightWithFocus
        .SelectionMode = flexSelectionByRow
        .AllowUserResizing = flexResizeColumns
        .ColAlignment(0) = 4
        .ColAlignment(1) = 1
        .ColAlignment(2) = 1
    End With
    
End Sub

Public Sub CargardgDeduccionesPersonales(PuestoLaboral As String)
    
    Dim i As Integer
    i = 0
    ListadoDeduccionesPersonales.dgDeduccionesPersonales.Rows = 2
    Set rstListadoSlave = New ADODB.Recordset
    rstListadoSlave.Open "Select * From CARGASDEFAMILIA Where PUESTOLABORAL = " & "'" & PuestoLaboral & "' and DEDUCIBLEGANANCIAS = True Order By FechaAlta", dbSlave, adOpenDynamic, adLockOptimistic
    With ListadoDeduccionesPersonales.dgDeduccionesPersonales
        If rstListadoSlave.BOF = False Then
            rstListadoSlave.MoveFirst
            While rstListadoSlave.EOF = False
                i = i + 1
                .RowHeight(i) = 300
                .TextMatrix(i, 0) = rstListadoSlave!FechaAlta
                .TextMatrix(i, 1) = rstListadoSlave!NombreCompleto
                Set rstBuscarSlave = New ADODB.Recordset
                rstBuscarSlave.Open "Select * From ASIGNACIONESFAMILIARES Where CODIGO = " & "'" & rstListadoSlave!CodigoParentesco & "'", dbSlave, adOpenDynamic, adLockOptimistic
                .TextMatrix(i, 2) = rstBuscarSlave!Parentesco
                rstBuscarSlave.Close
                Set rstBuscarSlave = Nothing
                rstListadoSlave.MoveNext
                .Rows = .Rows + 1
            Wend
        End If
        .Rows = .Rows - 1
    End With
    rstListadoSlave.Close
    Set rstListadoSlave = Nothing
    
End Sub

Public Sub ConfigurardgCodigoLiquidacion()
    
    With ListadoCodigoLiquidaciones.dgCodigoLiquidacion
        .Clear
        .Cols = 4
        .Rows = 2
        .TextMatrix(0, 0) = "Código"
        .TextMatrix(0, 1) = "Período"
        .TextMatrix(0, 2) = "Descripción"
        .TextMatrix(0, 3) = "$ Exento"
        .ColWidth(0) = 600
        .ColWidth(1) = 750
        .ColWidth(2) = 2650
        .ColWidth(3) = 850
        .FixedCols = 0
        .FocusRect = flexFocusHeavy
        .HighLight = flexHighlightWithFocus
        .SelectionMode = flexSelectionByRow
        .AllowUserResizing = flexResizeColumns
        .ColAlignment(0) = 4
        .ColAlignment(1) = 1
        .ColAlignment(2) = 1
        .ColAlignment(3) = 7
    
    End With
End Sub

Public Sub CargardgCodigoLiquidacion()
    
    Dim i As Integer
    i = 0
    ListadoCodigoLiquidaciones.dgCodigoLiquidacion.Rows = 2
    Set rstListadoSlave = New ADODB.Recordset
    rstListadoSlave.Open "Select * From CODIGOLIQUIDACIONES Order By CODIGO Desc", dbSlave, adOpenDynamic, adLockOptimistic
    With ListadoCodigoLiquidaciones.dgCodigoLiquidacion
        If rstListadoSlave.BOF = False Then
            rstListadoSlave.MoveFirst
            While rstListadoSlave.EOF = False
                i = i + 1
                .RowHeight(i) = 300
                .TextMatrix(i, 0) = rstListadoSlave!Codigo
                .TextMatrix(i, 1) = rstListadoSlave!Periodo
                .TextMatrix(i, 2) = rstListadoSlave!Descripcion
                .TextMatrix(i, 3) = De_Num_a_Tx_01(rstListadoSlave!MontoExento)
                rstListadoSlave.MoveNext
                .Rows = .Rows + 1
            Wend
        End If
        .Rows = .Rows - 1
        .SetFocus
    End With
    rstListadoSlave.Close
    Set rstListadoSlave = Nothing
    
End Sub

Public Sub ConfigurardgFamiliares()
    
    With ListadoFamiliares.dgFamiliares
        .Clear
        .Cols = 9
        .Rows = 2
        .TextMatrix(0, 0) = "Fecha"
        .ColWidth(0) = 1000
        .TextMatrix(0, 1) = "Nombre Completo"
        .ColWidth(1) = 2750
        .TextMatrix(0, 2) = "Parentesco"
        .ColWidth(2) = 1350
        .TextMatrix(0, 3) = "CUIL / DNI"
        .ColWidth(3) = 1150
        .TextMatrix(0, 4) = "Gcias."
        .ColWidth(4) = 750
        .TextMatrix(0, 5) = "Nivel Estudio"
        .ColWidth(5) = 1150
        .TextMatrix(0, 6) = "Adherente O.S."
        .ColWidth(6) = 1150
        .TextMatrix(0, 7) = "Cobra Salario"
        .ColWidth(7) = 1150
        .TextMatrix(0, 8) = "Discapacitado"
        .ColWidth(8) = 1150
        .FixedCols = 0
        .FocusRect = flexFocusHeavy
        .HighLight = flexHighlightWithFocus
        .SelectionMode = flexSelectionByRow
        .AllowUserResizing = flexResizeColumns
        .ColAlignment(0) = 1
        .ColAlignment(1) = 1
        .ColAlignment(2) = 1
        .ColAlignment(3) = 7
        .ColAlignment(4) = 4
        .ColAlignment(5) = 1
        .ColAlignment(6) = 4
        .ColAlignment(7) = 4
        .ColAlignment(8) = 4
    End With
    
End Sub

Public Sub CargardgFamiliares(PuestoLaboral As String)
    
    Dim i As Integer
    i = 0
    ListadoFamiliares.dgFamiliares.Rows = 2
    Set rstListadoSlave = New ADODB.Recordset
    rstListadoSlave.Open "Select * From CARGASDEFAMILIA Where PUESTOLABORAL = " & "'" & PuestoLaboral & "' Order By FechaAlta", dbSlave, adOpenDynamic, adLockOptimistic
    With ListadoFamiliares.dgFamiliares
        If rstListadoSlave.BOF = False Then
            rstListadoSlave.MoveFirst
            Set rstBuscarSlave = New ADODB.Recordset
            While rstListadoSlave.EOF = False
                i = i + 1
                .RowHeight(i) = 300
                .TextMatrix(i, 0) = rstListadoSlave!FechaAlta
                .TextMatrix(i, 1) = rstListadoSlave!NombreCompleto
                rstBuscarSlave.Open "Select * From ASIGNACIONESFAMILIARES Where CODIGO = " & "'" & rstListadoSlave!CodigoParentesco & "'", dbSlave, adOpenDynamic, adLockOptimistic
                .TextMatrix(i, 2) = rstBuscarSlave!Parentesco
                rstBuscarSlave.Close
                .TextMatrix(i, 3) = rstListadoSlave!DNI
                If rstListadoSlave!DeducibleGanancias = True Then
                    .TextMatrix(i, 4) = "SI"
                Else
                    .TextMatrix(i, 4) = "NO"
                End If
                .TextMatrix(i, 5) = rstListadoSlave!NivelDeEstudio
                If rstListadoSlave!AdherenteObraSocial = True Then
                    .TextMatrix(i, 6) = "SI"
                Else
                    .TextMatrix(i, 6) = "NO"
                End If
                If rstListadoSlave!CobraSalario = True Then
                    .TextMatrix(i, 7) = "SI"
                Else
                    .TextMatrix(i, 7) = "NO"
                End If
                If rstListadoSlave!Discapacitado = True Then
                    .TextMatrix(i, 8) = "SI"
                Else
                    .TextMatrix(i, 8) = "NO"
                End If
                rstListadoSlave.MoveNext
                .Rows = .Rows + 1
            Wend
        End If
        .Rows = .Rows - 1
    End With
    rstListadoSlave.Close
    Set rstListadoSlave = Nothing
    Set rstBuscarSlave = Nothing
    
End Sub

Public Sub ConfigurardgCodigosLiquidacionesGanancias()
    
    With ListadoLiquidacionGanancias.dgCodigosLiquidacionesGanancias
        .Clear
        .Cols = 2
        .Rows = 2
        .TextMatrix(0, 0) = "Código"
        .TextMatrix(0, 1) = "Descripción"
        .ColWidth(0) = 600
        .ColWidth(1) = 1700
        .FixedCols = 0
        .FocusRect = flexFocusHeavy
        .HighLight = flexHighlightWithFocus
        .SelectionMode = flexSelectionByRow
        .AllowUserResizing = flexResizeColumns
        .ColAlignment(0) = 4
        .ColAlignment(1) = 1
    End With
    
End Sub

Public Sub CargardgCodigosLiquidacionesGanancias()
    
    Dim i As Integer
    i = 0
    ListadoLiquidacionGanancias.dgCodigosLiquidacionesGanancias.Rows = 2
    Set rstListadoSlave = New ADODB.Recordset
    rstListadoSlave.Open "Select * From CODIGOLIQUIDACIONES Order By CODIGO Desc", dbSlave, adOpenDynamic, adLockOptimistic
    With ListadoLiquidacionGanancias.dgCodigosLiquidacionesGanancias
        If rstListadoSlave.BOF = False Then
            rstListadoSlave.MoveFirst
            While rstListadoSlave.EOF = False
                i = i + 1
                .RowHeight(i) = 300
                .TextMatrix(i, 0) = rstListadoSlave!Codigo
                .TextMatrix(i, 1) = rstListadoSlave!Descripcion
                rstListadoSlave.MoveNext
                .Rows = .Rows + 1
            Wend
        End If
        .Rows = .Rows - 1
        .SetFocus
    End With
    rstListadoSlave.Close
    Set rstListadoSlave = Nothing
    
End Sub

Public Sub ConfigurardgAgentesRetenidos()
    
    With ListadoLiquidacionGanancias.dgAgentesRetenidos
        .Clear
        .Cols = 3
        .Rows = 2
        .TextMatrix(0, 0) = "Puesto"
        .TextMatrix(0, 1) = "Nombre Completo"
        .TextMatrix(0, 2) = "Retención"
        .ColWidth(0) = 750
        .ColWidth(1) = 2250
        .ColWidth(2) = 850
        .FixedCols = 0
        .FocusRect = flexFocusHeavy
        .HighLight = flexHighlightWithFocus
        .SelectionMode = flexSelectionByRow
        .AllowUserResizing = flexResizeColumns
        .ColAlignment(0) = 1
        .ColAlignment(1) = 1
        .ColAlignment(2) = 7
    End With
    
End Sub

Public Sub CargardgAgentesRetenidosViejo(CodigoLiquidacion As String)
    
    Dim i As Integer
    Dim SQL As String
    Dim SQL2 As String
    Dim dblLimiteGanancia As Double
    Dim dblLimiteGananciaPersonal As Double
    Dim datFecha As Date
    Dim strCodigoAnterior As String
    Dim strCodigoGananciasAnterior As String
    
    i = 0
    ListadoLiquidacionGanancias.dgAgentesRetenidos.Rows = 2
    Set rstListadoSlave = New ADODB.Recordset
    Set rstBuscarSlave = New ADODB.Recordset
    
    'Buscamos los Agentes Retenidos en el Período anterior
    SQL = "Select * From LIQUIDACIONGANANCIAS4TACATEGORIA Where CODIGOLIQUIDACION = '" _
    & CodigoLiquidacion & "' Order By NombreCompleto Asc"
    rstListadoSlave.Open SQL, dbSlave, adOpenDynamic, adLockOptimistic
    With ListadoLiquidacionGanancias.dgAgentesRetenidos
        If rstListadoSlave.BOF = False Then
            rstListadoSlave.MoveFirst
            While rstListadoSlave.EOF = False
                i = i + 1
                .RowHeight(i) = 300
                .TextMatrix(i, 0) = rstListadoSlave!PuestoLaboral
                SQL = "Select * From AGENTES Where PUESTOLABORAL = '" & rstListadoSlave!PuestoLaboral & "'"
                rstBuscarSlave.Open SQL, dbSlave, adOpenDynamic, adLockOptimistic
                .TextMatrix(i, 1) = rstBuscarSlave!NombreCompleto
                rstBuscarSlave.Close
                .TextMatrix(i, 2) = rstListadoSlave!Retencion
                rstListadoSlave.MoveNext
                .Rows = .Rows + 1
            Wend
        End If
        rstListadoSlave.Close
    End With
    
    'Buscamos los Agentes Retenidos en el Período anterior
    SQL = "Select CODIGOLIQUIDACION From LIQUIDACIONGANANCIAS4TACATEGORIA Group by CODIGOLIQUIDACION " _
    & "Having CODIGOLIQUIDACION < '" & CodigoLiquidacion & "' Order by CODIGOLIQUIDACION Desc"
    rstBuscarSlave.Open SQL, dbSlave, adOpenDynamic, adLockOptimistic
    If rstBuscarSlave.BOF = False Then
        strCodigoGananciasAnterior = rstBuscarSlave!CodigoLiquidacion
    Else
        strCodigoGananciasAnterior = CodigoLiquidacion
    End If
    rstBuscarSlave.Close
    SQL = "Select * From LIQUIDACIONGANANCIAS4TACATEGORIA Where CODIGOLIQUIDACION = '" & strCodigoGananciasAnterior & "' " _
    & "Order by PUESTOLABORAL"
    rstListadoSlave.Open SQL, dbSlave, adOpenDynamic, adLockOptimistic
    With ListadoLiquidacionGanancias.dgAgentesRetenidos
        While rstListadoSlave.EOF = False
            SQL = "Select * From LIQUIDACIONGANANCIAS4TACATEGORIA Where CODIGOLIQUIDACION = '" _
            & CodigoLiquidacion & "' And PUESTOLABORAL = '" & rstListadoSlave!PuestoLaboral & "'"
            If SQLNoMatch(SQL) = False Then
                rstListadoSlave.MoveNext 'Si el agente ya tiene liquidación de Ganancias en el período pasamos de largo
            Else
                i = i + 1
                .RowHeight(i) = 300
                .TextMatrix(i, 0) = rstListadoSlave!PuestoLaboral
                SQL = "Select * From AGENTES Where PUESTOLABORAL = '" & rstListadoSlave!PuestoLaboral & "'"
                Set rstBuscarSlave = New ADODB.Recordset
                rstBuscarSlave.Open SQL, dbSlave, adOpenDynamic, adLockOptimistic
                .TextMatrix(i, 1) = rstBuscarSlave!NombreCompleto
                rstBuscarSlave.Close
                Set rstBuscarSlave = Nothing
                .TextMatrix(i, 2) = "Sugerido"
                rstListadoSlave.MoveNext
                .Rows = .Rows + 1
            End If
        Wend
    End With
    rstListadoSlave.Close
    
    'Por último, buscamos aquellos agentes que no fueron retenidos en este período ni en el anterior pero, _
    dado su nivel de ingreso, se sugiere calcular retención
    'El procedimiento actual solo contempla el último sueldo y no el acumulado anual.
    'Buscamos el Período de Liquidación
    Set rstBuscarSlave = New ADODB.Recordset
    SQL = "Select * From CODIGOLIQUIDACIONES Where CODIGO = '" & CodigoLiquidacion & "'"
    rstBuscarSlave.Open SQL, dbSlave, adOpenDynamic, adLockOptimistic
    datFecha = DateTime.DateSerial(Right(rstBuscarSlave!Periodo, 4), Left(rstBuscarSlave!Periodo, 2), 1)
    datFecha = DateAdd("m", 1, datFecha)
    datFecha = DateAdd("d", -1, datFecha)
    rstBuscarSlave.Close
    'Buscamos los Importes de Mínimo no Imponible y Deducción Especial _
    de la Norma más reciente respecto del Período de Liquidación
    SQL = "Select * From DEDUCCIONES4TACATEGORIA Where FECHA <= #" & Format(datFecha, "MM/DD/YYYY") & "# Order by FECHA Desc"
    rstBuscarSlave.Open SQL, dbSlave, adOpenDynamic, adLockOptimistic
    dblLimiteGanancia = Round((rstBuscarSlave!MinimoNoImponible + rstBuscarSlave!DeduccionEspecial) / 13, 0)
    dblLimiteGanancia = Round(((dblLimiteGanancia + 3.25) / 0.765), 0)
    rstBuscarSlave.Close
    'Seleccionamos a aquellos agentes cuyo Haber Óptimo, en la liquidación actual o en la anterior, _
    supere el Limite calculado recientemente
    SQL = "Select CODIGOLIQUIDACION From LIQUIDACIONSUELDOS Where CODIGOLIQUIDACION <= '" & CodigoLiquidacion & "' Order by CODIGOLIQUIDACION Desc"
    rstBuscarSlave.Open SQL, dbSlave, adOpenDynamic, adLockOptimistic
    strCodigoAnterior = rstBuscarSlave!CodigoLiquidacion 'Que pasa con SAC y Complementos?
    rstBuscarSlave.Close
    Set rstBuscarSlave = Nothing
    SQL = "Select * From LIQUIDACIONSUELDOS Where CODIGOLIQUIDACION = '" & strCodigoAnterior & "' " _
    & "And CODIGOCONCEPTO = '0115' And Importe > " & dblLimiteGanancia & " Order by PUESTOLABORAL"
    rstListadoSlave.Open SQL, dbSlave, adOpenDynamic, adLockOptimistic
    With ListadoLiquidacionGanancias.dgAgentesRetenidos
        If rstListadoSlave.BOF = False Then
            rstListadoSlave.MoveFirst
            While rstListadoSlave.EOF = False
                SQL = "Select * From LIQUIDACIONGANANCIAS4TACATEGORIA Where CODIGOLIQUIDACION = '" _
                & CodigoLiquidacion & "' And PUESTOLABORAL = '" & rstListadoSlave!PuestoLaboral & "'"
                SQL2 = "Select * From LIQUIDACIONGANANCIAS4TACATEGORIA Where CODIGOLIQUIDACION = '" & strCodigoGananciasAnterior & "' " _
                & "And PUESTOLABORAL = '" & rstListadoSlave!PuestoLaboral & "'"
                If SQLNoMatch(SQL) = False Then
                    rstListadoSlave.MoveNext 'Si al agente ya tiene liquidación de Ganancias en el período pasamos de largo
                ElseIf SQLNoMatch(SQL2) = False Then
                    rstListadoSlave.MoveNext 'Si al agente ya fue sugerido por estar retenido en período anterior, pasamos de largo
                Else
                    dblLimiteGananciaPersonal = dblLimiteGanancia
                    If TieneConyugeDeducible(rstListadoSlave!PuestoLaboral) = True Then
                        dblLimiteGananciaPersonal = dblLimiteGananciaPersonal + (ImporteDeduccionPersonal("CONYUGE", datFecha) * 12 / 13)
                    End If
                    If CantidadHijosDeducibles(rstListadoSlave!PuestoLaboral, datFecha) > 0 Then
                        dblLimiteGananciaPersonal = dblLimiteGananciaPersonal + (ImporteDeduccionPersonal("HIJO", datFecha) _
                        * CantidadHijosDeducibles(rstListadoSlave!PuestoLaboral, datFecha) * 12 / 13)
                    End If
                    If CantidadOtrasCargasFamiliaDeducibles(rstListadoSlave!PuestoLaboral) > 0 Then
                        dblLimiteGananciaPersonal = dblLimiteGananciaPersonal + (ImporteDeduccionPersonal("OTRASCARGASDEFAMILIA", datFecha) _
                        * CantidadOtrasCargasFamiliaDeducibles(rstListadoSlave!PuestoLaboral) * 12 / 13)
                    End If
                    If TieneDeduccionGeneral("SERVICIODOMESTICO", rstListadoSlave!PuestoLaboral, datFecha) = True Then
                        dblLimiteGananciaPersonal = dblLimiteGananciaPersonal + (ImporteDeduccionGeneral(rstListadoSlave!PuestoLaboral, "SERVICIODOMESTICO", datFecha) * 12 / 13)
                    End If
                    SQL = "Select * from LIQUIDACIONSUELDOS Where ((PUESTOLABORAL = '" & rstListadoSlave!PuestoLaboral & "' And CODIGOLIQUIDACION = '" & strCodigoGananciasAnterior & "' And CODIGOCONCEPTO = '0317') Or " _
                    & "(PUESTOLABORAL = '" & rstListadoSlave!PuestoLaboral & "' And CODIGOLIQUIDACION = '" & strCodigoGananciasAnterior & "' And CODIGOCONCEPTO = '0361') Or " _
                    & "(PUESTOLABORAL = '" & rstListadoSlave!PuestoLaboral & "' And CODIGOLIQUIDACION = '" & strCodigoGananciasAnterior & "' And CODIGOCONCEPTO = '0367') Or " _
                    & "(PUESTOLABORAL = '" & rstListadoSlave!PuestoLaboral & "' And CODIGOLIQUIDACION = '" & strCodigoGananciasAnterior & "' And CODIGOCONCEPTO = '0370') Or " _
                    & "(PUESTOLABORAL = '" & rstListadoSlave!PuestoLaboral & "' And CODIGOLIQUIDACION = '" & strCodigoGananciasAnterior & "' And CODIGOCONCEPTO = '0373') Or " _
                    & "(PUESTOLABORAL = '" & rstListadoSlave!PuestoLaboral & "' And CODIGOLIQUIDACION = '" & strCodigoGananciasAnterior & "' And CODIGOCONCEPTO = '0374'))"
                    If TieneDeduccionGeneral("SEGURODEVIDA", rstListadoSlave!PuestoLaboral, datFecha) = True Or SQLNoMatch(SQL) = False Then
                        If SQLNoMatch(SQL) = False Then
                            Set rstRegistroSlave = New ADODB.Recordset
                            rstRegistroSlave.Open SQL, dbSlave, adOpenDynamic, adLockOptimistic
                            dblLimiteGananciaPersonal = dblLimiteGananciaPersonal _
                            + (ImporteDeduccionGeneral(rstListadoSlave!PuestoLaboral, "SEGURODEVIDA", datFecha, , rstRegistroSlave!Importe) * 12 / 13)
                            rstRegistroSlave.Close
                            Set rstRegistroSlave = Nothing
                        Else
                            dblLimiteGananciaPersonal = dblLimiteGananciaPersonal _
                            + (ImporteDeduccionGeneral(rstListadoSlave!PuestoLaboral, "SEGURODEVIDA", datFecha) * 12 / 13)
                        End If
                    End If
                    If (rstListadoSlave!Importe - dblLimiteGananciaPersonal) > 0 Then
                        If TieneDeduccionGeneral("CUOTAMEDICOASISTENCIAL", rstListadoSlave!PuestoLaboral, datFecha) = True Then
                            dblLimiteGananciaPersonal = dblLimiteGananciaPersonal _
                            + (ImporteDeduccionGeneral(rstListadoSlave!PuestoLaboral, "CUOTAMEDICOASISTENCIAL", datFecha, (rstListadoSlave!Importe - dblLimiteGananciaPersonal)) * 12 / 13)
                        End If
                    End If
                    'FALTÓ CONSIDERAR DONACIONES
                    If rstListadoSlave!Importe < dblLimiteGananciaPersonal Then
                        rstListadoSlave.MoveNext 'No deducible por demás Deducciones Personales y Generales informadas
                    Else
                        i = i + 1
                        .RowHeight(i) = 300
                        .TextMatrix(i, 0) = rstListadoSlave!PuestoLaboral
                        SQL = "Select * From AGENTES Where PUESTOLABORAL = '" & rstListadoSlave!PuestoLaboral & "'"
                        Set rstBuscarSlave = New ADODB.Recordset
                        rstBuscarSlave.Open SQL, dbSlave, adOpenDynamic, adLockOptimistic
                        .TextMatrix(i, 1) = rstBuscarSlave!NombreCompleto
                        rstBuscarSlave.Close
                        Set rstBuscarSlave = Nothing
                        .TextMatrix(i, 2) = "Sugerido"
                        rstListadoSlave.MoveNext
                        .Rows = .Rows + 1
                    End If
                End If
            Wend
        End If
        If .Rows > 2 Then
            .Rows = .Rows - 1
        End If
    End With
    
    i = 0
    SQL = ""
    SQL2 = ""
    dblLimiteGanancia = 0
    datFecha = 0
    strCodigoAnterior = ""
    strCodigoGananciasAnterior = ""
    rstListadoSlave.Close
    Set rstListadoSlave = Nothing
    Set rstBuscarSlave = Nothing
    
End Sub

Public Sub CargardgAgentesRetenidos(CodigoLiquidacion As String)
    
    Dim i As Integer
    Dim SQL As String
    Dim SQLDesactivado As String
    Dim dblLimiteGanancia As Double
    Dim dblImporteMensual As Double
    Dim datFecha As Date
    Dim datFechaControl As Date
    Dim strPeriodo As String
    Dim strCodigoGananciasAnterior As String
    Dim bolMismoEjercicio As Boolean
    
    i = 0
    ListadoLiquidacionGanancias.dgAgentesRetenidos.Rows = 2
    Set rstListadoSlave = New ADODB.Recordset
    
    'Buscamos los Agentes Retenidos en el Período actual
    SQL = "Select LIQUIDACIONGANANCIAS4TACATEGORIA.PuestoLaboral," _
    & " LIQUIDACIONGANANCIAS4TACATEGORIA.Retencion, AGENTES.NombreCompleto" _
    & " From LIQUIDACIONGANANCIAS4TACATEGORIA Inner Join AGENTES" _
    & " On AGENTES.PuestoLaboral = LIQUIDACIONGANANCIAS4TACATEGORIA.PuestoLaboral" _
    & " Where CODIGOLIQUIDACION = '" & CodigoLiquidacion _
    & "' Order By AGENTES.NombreCompleto Asc"
    rstListadoSlave.Open SQL, dbSlave, adOpenDynamic, adLockOptimistic
    With ListadoLiquidacionGanancias.dgAgentesRetenidos
        If rstListadoSlave.BOF = False Then
            rstListadoSlave.MoveFirst
            While rstListadoSlave.EOF = False
                i = i + 1
                .RowHeight(i) = 300
                .TextMatrix(i, 0) = rstListadoSlave!PuestoLaboral
                .TextMatrix(i, 1) = rstListadoSlave!NombreCompleto
                .TextMatrix(i, 2) = rstListadoSlave!Retencion
                SQLDesactivado = "SELECT NombreCompleto FROM Agentes " _
                & "WHERE Activado = False " _
                & "AND NombreCompleto = '" & rstListadoSlave!NombreCompleto & "'"
                If SQLNoMatch(SQLDesactivado) = False Then
                    .col = 1
                    .Row = i
                    .CellBackColor = vbRed
                End If
                rstListadoSlave.MoveNext
                .Rows = .Rows + 1
            Wend
        End If
        rstListadoSlave.Close
    End With
    'Creamos un SQL para filtrar la posterior consulta
    SQL = "Select LIQUIDACIONGANANCIAS4TACATEGORIA.PuestoLaboral" _
    & " From LIQUIDACIONGANANCIAS4TACATEGORIA" _
    & " Where CODIGOLIQUIDACION = '" & CodigoLiquidacion _
    & "' Order By LIQUIDACIONGANANCIAS4TACATEGORIA.PUESTOLABORAL"
    
    'Buscamos los Agentes Retenidos en el Período anterior
    strCodigoGananciasAnterior = BuscarCodigoLiquidacionGananciasAnterior(CodigoLiquidacion)
    If Right(BuscarPeriodoLiquidacion(CodigoLiquidacion), 2) = Right(BuscarPeriodoLiquidacion(strCodigoGananciasAnterior), 2) Then
        bolMismoEjercicio = True
    Else
        bolMismoEjercicio = False
    End If
    If bolMismoEjercicio = True Then
        SQL = "Select LIQUIDACIONGANANCIAS4TACATEGORIA.PuestoLaboral," _
        & " LIQUIDACIONGANANCIAS4TACATEGORIA.Retencion, AGENTES.NombreCompleto" _
        & " From LIQUIDACIONGANANCIAS4TACATEGORIA Inner Join AGENTES" _
        & " On AGENTES.PuestoLaboral = LIQUIDACIONGANANCIAS4TACATEGORIA.PuestoLaboral" _
        & " Where CODIGOLIQUIDACION = '" & strCodigoGananciasAnterior _
        & "' And LIQUIDACIONGANANCIAS4TACATEGORIA.PuestoLaboral Not In (" & SQL _
        & ") Order By AGENTES.NombreCompleto Asc"
        rstListadoSlave.Open SQL, dbSlave, adOpenDynamic, adLockOptimistic
        With ListadoLiquidacionGanancias.dgAgentesRetenidos
            If .Rows > 2 Then
                i = .Rows - 2
            Else
                i = 0
            End If
            If rstListadoSlave.BOF = False Then
                While rstListadoSlave.EOF = False
                    i = i + 1
                    .RowHeight(i) = 300
                    .TextMatrix(i, 0) = rstListadoSlave!PuestoLaboral
                    SQL = "Select * From AGENTES Where PUESTOLABORAL = '" & rstListadoSlave!PuestoLaboral & "'"
                    .TextMatrix(i, 1) = rstListadoSlave!NombreCompleto
                    .TextMatrix(i, 2) = "Ret. Prev."
                    SQLDesactivado = "SELECT NombreCompleto FROM Agentes " _
                    & "WHERE Activado = False " _
                    & "AND NombreCompleto = '" & rstListadoSlave!NombreCompleto & "'"
                    If SQLNoMatch(SQLDesactivado) = False Then
                        .col = 1
                        .Row = i
                        .CellBackColor = vbRed
                    End If
                    rstListadoSlave.MoveNext
                    .Rows = .Rows + 1
                Wend
            End If
        End With
        rstListadoSlave.Close
    End If
    'Por último, buscamos aquellos agentes que no fueron retenidos en este período ni en el anterior pero, _
    dado su nivel de ingreso, se sugiere calcular retención.
    'Buscamos los Importes de Mínimo no Imponible y Deducción Especial _
    de la Norma más reciente respecto del Período de Liquidación
    'Nos posicionamos en el último día del período a liquidar
    strPeriodo = BuscarPeriodoLiquidacion(CodigoLiquidacion)
    datFecha = BuscarUltimoDiaDelPeriodo(strPeriodo)
    'Controlamos lo que debió liquidarse en concepto de Mínimo no Imponible Acumulado al período de liquidación
    dblImporteMensual = ImporteDeduccionPersonal("MINIMONOIMPONIBLE", DateSerial(Year(datFecha) - 1, 12, 31))
    SQL = "Select MINIMONOIMPONIBLE From DEDUCCIONES4TACATEGORIA Where FECHA <= #" & Format(datFecha, "MM/DD/YYYY") & "# " _
    & "And Year(FECHA) = '" & Year(datFecha) & "' Order by FECHA Desc"
    If SQLNoMatch(SQL) = True Then
        dblLimiteGanancia = dblImporteMensual * Month(datFecha)
    Else
        For i = 1 To Month(datFecha)
            datFechaControl = DateSerial(Year(datFecha), i, 1)
            datFechaControl = DateAdd("m", 1, datFechaControl)
            datFechaControl = DateAdd("d", -1, datFechaControl)
            dblImporteMensual = ImporteDeduccionPersonal("MINIMONOIMPONIBLE", datFechaControl)
            dblLimiteGanancia = dblLimiteGanancia + dblImporteMensual
        Next i
        datFechaControl = 0
    End If
    'Controlamos lo que debió liquidarse en concepto de Deducción Especial Acumulado al período de liquidación
    'Sumamos ambos conceptos
    dblImporteMensual = ImporteDeduccionPersonal("DEDUCCIONESPECIAL", DateSerial(Year(datFecha) - 1, 12, 31))
    SQL = "Select DEDUCCIONESPECIAL From DEDUCCIONES4TACATEGORIA Where FECHA <= #" & Format(datFecha, "MM/DD/YYYY") & "# " _
    & "And Year(FECHA) = '" & Year(datFecha) & "' Order by FECHA Desc"
    If SQLNoMatch(SQL) = True Then
        dblLimiteGanancia = dblLimiteGanancia + (dblImporteMensual * Month(datFecha))
    Else
        For i = 1 To Month(datFecha)
            datFechaControl = DateSerial(Year(datFecha), i, 1)
            datFechaControl = DateAdd("m", 1, datFechaControl)
            datFechaControl = DateAdd("d", -1, datFechaControl)
            dblImporteMensual = ImporteDeduccionPersonal("DEDUCCIONESPECIAL", datFechaControl)
            dblLimiteGanancia = dblLimiteGanancia + dblImporteMensual
        Next i
        datFechaControl = 0
    End If
    'Convertimos el valor hallado en un Haber Óptimo Acumulado que actúe como límite
    dblLimiteGanancia = ((dblLimiteGanancia + 3.25) / 0.765)
    'Seleccionamos a aquellos agentes cuyo Haber Óptimo Acumulado supere el Limite calculado recientemente
    'Empezamos por crear una tabla adicional para calcular los Haberes Óptimos Acumulados
    SQL = "Create Table CALCULOSAUXILIARES (" _
    & "PuestoLaboral Text (6), HaberBruto Currency, " _
    & "AsignacionFamiliar Currency, Discapacitado Currency, " _
    & "SAC Currency, HaberOptimo Currency)"
    dbSlave.BeginTrans
    dbSlave.Execute SQL
    dbSlave.CommitTrans
    'Creamos un SQL para filtrar la posterior consulta
    SQL = "Select LIQUIDACIONGANANCIAS4TACATEGORIA.PuestoLaboral" _
    & " From LIQUIDACIONGANANCIAS4TACATEGORIA" _
    & " Where CODIGOLIQUIDACION = '" & CodigoLiquidacion & "'"
    If bolMismoEjercicio = True Then
        SQL = SQL & " Or CODIGOLIQUIDACION = '" & strCodigoGananciasAnterior & "'"
    End If
    'Copiamos el concepto Haber Bruto Acumulado (9998) de la liquidación base a la tabla de cálculos auxiliares
    SQL = "Insert Into CALCULOSAUXILIARES (PuestoLaboral, HaberBruto, AsignacionFamiliar, Discapacitado, SAC, HaberOptimo)" _
    & " Select PuestoLaboral, (Sum(Importe)*(13/12)) As HB, 0, 0, 0, 0" _
    & " From LIQUIDACIONSUELDOS Inner Join CODIGOLIQUIDACIONES" _
    & " On LIQUIDACIONSUELDOS.CodigoLiquidacion  = CODIGOLIQUIDACIONES.Codigo" _
    & " Where Right(PERIODO,4) = '" & Right(strPeriodo, 4) _
    & "' And CODIGO <= '" & CodigoLiquidacion _
    & "' And CODIGOCONCEPTO = '9998'" _
    & " And PuestoLaboral Not In (" & SQL & ")" _
    & " Group by PuestoLaboral" _
    & " Having (Sum(Importe)*(13/12)) > " & De_Num_a_Tx_01(dblLimiteGanancia, , 2)
    dbSlave.BeginTrans
    dbSlave.Execute SQL
    dbSlave.CommitTrans
    'Al parecer, SQL no permite hacer UPDATE con una doble consulta anidada (Inner Join) por que procede a crear otra tabla auxiliar
    SQL = "Create Table CALCULOSAUXILIARES2 (" _
    & "PuestoLaboral Text (6), SumaAcumulada Currency)"
    dbSlave.BeginTrans
    dbSlave.Execute SQL
    dbSlave.CommitTrans
    'Incorporamos las Asignaciones Familiares Acumuladas a CALCULOSAXILIARES2
    SQL = "Insert Into CALCULOSAUXILIARES2 (PuestoLaboral, SumaAcumulada)" _
    & " Select PUESTOLABORAL, (Sum(Importe)*(13/12)) As AF" _
    & " From LIQUIDACIONSUELDOS Inner Join CODIGOLIQUIDACIONES" _
    & " On LIQUIDACIONSUELDOS.CodigoLiquidacion  = CODIGOLIQUIDACIONES.Codigo" _
    & " Where Right(PERIODO,4) = '" & Right(strPeriodo, 4) _
    & "' And CODIGO <= '" & CodigoLiquidacion _
    & "' And CODIGOCONCEPTO = '0003'" _
    & " Group by PuestoLaboral"
    dbSlave.BeginTrans
    dbSlave.Execute SQL
    dbSlave.CommitTrans
    'Modificamos el campo AsignacionFamiliar de la Tabla de Calculos Auxiliares _
    para incorporar la AsignacionFamiliar de la tabla Calculos Auxiliares 2
    SQL = "Update CALCULOSAUXILIARES Inner Join CALCULOSAUXILIARES2" _
    & " On CALCULOSAUXILIARES.PuestoLaboral = CALCULOSAUXILIARES2.PuestoLaboral" _
    & " Set CALCULOSAUXILIARES.AsignacionFamiliar = CALCULOSAUXILIARES2.SUMAACUMULADA"
    dbSlave.BeginTrans
    dbSlave.Execute SQL
    dbSlave.CommitTrans
    'Eliminamos el contenido de la segunda tabla auxiliar
    SQL = "Delete From CALCULOSAUXILIARES2"
    dbSlave.BeginTrans
    dbSlave.Execute SQL
    dbSlave.CommitTrans
    'Incorporamos el concepto Discapacitados Acumulados a CALCULOSAXILIARES2
    SQL = "Insert Into CALCULOSAUXILIARES2 (PuestoLaboral, SumaAcumulada)" _
    & " Select PUESTOLABORAL, (Sum(Importe)*(13/12)) As DISC" _
    & " From LIQUIDACIONSUELDOS Inner Join CODIGOLIQUIDACIONES" _
    & " On LIQUIDACIONSUELDOS.CodigoLiquidacion  = CODIGOLIQUIDACIONES.Codigo" _
    & " Where Right(PERIODO,4) = '" & Right(strPeriodo, 4) _
    & "' And CODIGO <= '" & CodigoLiquidacion _
    & "' And CODIGOCONCEPTO = '0058'" _
    & " Group by PuestoLaboral"
    dbSlave.BeginTrans
    dbSlave.Execute SQL
    dbSlave.CommitTrans
    'Modificamos el campo Discapacitado de la Tabla de Calculos Auxiliares _
    para incorporar Discapacitado de la tabla Calculos Auxiliares 2
    SQL = "Update CALCULOSAUXILIARES Inner Join CALCULOSAUXILIARES2" _
    & " On CALCULOSAUXILIARES.PuestoLaboral = CALCULOSAUXILIARES2.PuestoLaboral" _
    & " Set CALCULOSAUXILIARES.Discapacitado = CALCULOSAUXILIARES2.SUMAACUMULADA"
    dbSlave.BeginTrans
    dbSlave.Execute SQL
    dbSlave.CommitTrans
    'Eliminamos el contenido de la segunda tabla auxiliar
    SQL = "Delete From CALCULOSAUXILIARES2"
    dbSlave.BeginTrans
    dbSlave.Execute SQL
    dbSlave.CommitTrans
    'Incorporamos el concepto SAC Acumulado a CALCULOSAXILIARES2 (HAY QUE VERIFICAR)
    SQL = "Insert Into CALCULOSAUXILIARES2 (PuestoLaboral, SumaAcumulada)" _
    & " Select PUESTOLABORAL, (Sum(Importe)*(13/12)) As SAC" _
    & " From LIQUIDACIONSUELDOS Inner Join CODIGOLIQUIDACIONES" _
    & " On LIQUIDACIONSUELDOS.CodigoLiquidacion  = CODIGOLIQUIDACIONES.Codigo" _
    & " Where Right(PERIODO,4) = '" & Right(strPeriodo, 4) _
    & "' And CODIGO <= '" & CodigoLiquidacion _
    & "' And CODIGOCONCEPTO = '0150'" _
    & " Group by PuestoLaboral"
    dbSlave.BeginTrans
    dbSlave.Execute SQL
    dbSlave.CommitTrans
    'Modificamos el campo SAC de la Tabla de Calculos Auxiliares _
    para incorporar Discapacitado de la tabla Calculos Auxiliares 2
    SQL = "Update CALCULOSAUXILIARES Inner Join CALCULOSAUXILIARES2" _
    & " On CALCULOSAUXILIARES.PuestoLaboral = CALCULOSAUXILIARES2.PuestoLaboral" _
    & " Set CALCULOSAUXILIARES.SAC = CALCULOSAUXILIARES2.SUMAACUMULADA"
    dbSlave.BeginTrans
    dbSlave.Execute SQL
    dbSlave.CommitTrans
    'Borramos la segunda tabla auxiliar
    SQL = "Drop Table CALCULOSAUXILIARES2"
    dbSlave.BeginTrans
    dbSlave.Execute SQL
    dbSlave.CommitTrans
    'Calculamos el Haber Óptimo Acumulado con los datos obtenidos
    SQL = "Update CALCULOSAUXILIARES " _
    & "Set HaberOptimo = Format((HaberBruto - SAC - AsignacionFamiliar - Discapacitado), '#.00')"
    dbSlave.BeginTrans
    dbSlave.Execute SQL
    dbSlave.CommitTrans
    'Utilizamos los datos de la tabla CALCULOSAUXILIARES para llenar el DataGrid
    SQL = "Select CALCULOSAUXILIARES.PuestoLaboral, AGENTES.NombreCompleto" _
    & " From CALCULOSAUXILIARES Inner Join AGENTES" _
    & " On AGENTES.PuestoLaboral = CALCULOSAUXILIARES.PuestoLaboral" _
    & " Where HaberOptimo  > " & De_Num_a_Tx_01(dblLimiteGanancia, , 2) _
    & " Order By AGENTES.NombreCompleto Asc"
    rstListadoSlave.Open SQL, dbSlave, adOpenDynamic, adLockOptimistic
    With ListadoLiquidacionGanancias.dgAgentesRetenidos
        If .Rows > 2 Then
            i = .Rows - 2
        Else
            i = 0
        End If
        If rstListadoSlave.BOF = False Then
            rstListadoSlave.MoveFirst
            While rstListadoSlave.EOF = False
                i = i + 1
                .RowHeight(i) = 300
                .TextMatrix(i, 0) = rstListadoSlave!PuestoLaboral
                .TextMatrix(i, 1) = rstListadoSlave!NombreCompleto
                .TextMatrix(i, 2) = "Sup. MNO"
                SQLDesactivado = "SELECT NombreCompleto FROM Agentes " _
                    & "WHERE Activado = False " _
                    & "AND NombreCompleto = '" & rstListadoSlave!NombreCompleto & "'"
                If SQLNoMatch(SQLDesactivado) = False Then
                    .col = 1
                    .Row = i
                    .CellBackColor = vbRed
                End If
                rstListadoSlave.MoveNext
                .Rows = .Rows + 1
            Wend
        End If
        rstListadoSlave.Close
    End With
    'Por último, borramos la tabla auxiliar
    SQL = "Drop Table CALCULOSAUXILIARES"
    dbSlave.BeginTrans
    dbSlave.Execute SQL
    dbSlave.CommitTrans
    
    With ListadoLiquidacionGanancias.dgAgentesRetenidos
        If .Rows > 2 Then
            .Rows = .Rows - 1
        End If
    End With
    
    i = 0
    SQL = ""
    dblImporteMensual = 0
    dblLimiteGanancia = 0
    datFecha = 0
    datFechaControl = 0
    strPeriodo = ""
    strCodigoGananciasAnterior = ""
    Set rstListadoSlave = Nothing
    
End Sub

Public Sub ConfigurardgDeduccionesPersonalesLG4ta()
    
    With LiquidacionGanancia4ta.dgDeduccionesPersonales
        .Clear
        .Cols = 2
        .Rows = 2
        .TextMatrix(0, 0) = "Concepto"
        .TextMatrix(0, 1) = "Importe"
        .ColWidth(0) = 1600
        .ColWidth(1) = 800
        .FixedCols = 0
        .FocusRect = flexFocusHeavy
        .HighLight = flexHighlightWithFocus
        .SelectionMode = flexSelectionByRow
        .AllowUserResizing = flexResizeColumns
        .ColAlignment(0) = 1
        .ColAlignment(1) = 7
    End With
    
End Sub

'Public Sub CargardgDeduccionesPersonalesLG4ta()
'
'    Dim SQL As String
'    Dim i As Byte
'    Dim datFecha As Date
'    Dim dat24Years As Date
'    Dim datFechaControl As Date
'    Dim dblImporteMensual As Double
'    Dim dblLiquidado As Double
'    Dim dblALiquidarse As Double
'    Dim dblAcumulado As Double
'    Dim intFamiliares As Integer
'
'    'Nos posicionamos en el último día del período a liquidar
'    With LiquidacionGanancia4ta
''        datFecha = DateTime.DateSerial(Right(.txtPeriodo.Text, 4), Left(.txtPeriodo.Text, 2), 1)
''        datFecha = DateAdd("m", 1, datFecha)
''        datFecha = DateAdd("d", -1, datFecha)
'        SQL = BuscarPeriodoLiquidacion(Left(.txtCodigoLiquidacion.Text, 4))
'        datFecha = BuscarUltimoDiaDelPeriodo(SQL)
'    End With
'
'
'
'    Set rstListadoSlave = New ADODB.Recordset
'    With LiquidacionGanancia4ta.dgDeduccionesPersonales
'        .Rows = 9
'        .RowHeight(0) = 300
'        .TextMatrix(1, 0) = "Mín. no Imponible"
'        'Controlamos lo que debió liquidarse en concepto de Mínimo no Imponible
'        dblImporteMensual = ImporteDeduccionPersonal("MINIMONOIMPONIBLE", DateSerial(Year(datFecha) - 1, 12, 31))
'        SQL = "Select MINIMONOIMPONIBLE From DEDUCCIONES4TACATEGORIA Where FECHA <= #" & Format(datFecha, "MM/DD/YYYY") & "# " _
'        & "And Year(Fecha) = '" & Year(datFecha) & "' Order by FECHA Desc"
'        If SQLNoMatch(SQL) = True Then
'            dblALiquidarse = dblImporteMensual * Month(datFecha)
'        Else
'            For i = 1 To Month(datFecha)
'                datFechaControl = DateSerial(Year(datFecha), i, 1)
'                datFechaControl = DateAdd("m", 1, datFechaControl)
'                datFechaControl = DateAdd("d", -1, datFechaControl)
'                dblImporteMensual = ImporteDeduccionPersonal("MINIMONOIMPONIBLE", datFechaControl)
'                dblALiquidarse = dblALiquidarse + dblImporteMensual
'            Next i
'        End If
'        'Controlamos lo que realmente se liquidó en concepto de Mínimo no Imponible
'        SQL = "Select  Sum(LIQUIDACIONGANANCIAS4TACATEGORIA.MinimoNoImponible) AS SumaDeMinimoNoImponible " _
'        & "From CODIGOLIQUIDACIONES Inner Join LIQUIDACIONGANANCIAS4TACATEGORIA On CODIGOLIQUIDACIONES.Codigo = LIQUIDACIONGANANCIAS4TACATEGORIA.CodigoLiquidacion " _
'        & "Where PUESTOLABORAL = '" & LiquidacionGanancia4ta.txtPuestoLaboral.Text & "' " _
'        & "And Right(PERIODO,4) = '" & Right(LiquidacionGanancia4ta.txtPeriodo.Text, 4) & "' And CODIGO < '" & LiquidacionGanancia4ta.txtCodigoLiquidacion.Text & "'"
'        rstListadoSlave.Open SQL, dbSlave, adOpenDynamic, adLockOptimistic
'        If rstListadoSlave.BOF = False And IsNull(rstListadoSlave!SumaDeMinimoNoImponible) = False Then
'            dblLiquidado = rstListadoSlave!SumaDeMinimoNoImponible
'            dblAcumulado = dblAcumulado + dblLiquidado
'        Else
'            dblLiquidado = 0
'        End If
'        rstListadoSlave.Close
'        'Cargamos el Importe de la diferencia entre lo que debió liquidarse y lo que se liquidó en concepto de Mínimo no Imponible
'        .TextMatrix(1, 1) = De_Num_a_Tx_01(dblALiquidarse - dblLiquidado)
'        dblTotal = dblTotal + De_Txt_a_Num_01(.TextMatrix(1, 1))
'        dblALiquidarse = 0
'
'        .TextMatrix(2, 0) = "Cargas de Familia"
'
'        .TextMatrix(3, 0) = " - Conyuge" 'Tener en cuenta que no está previsto el caso de ALTA/BAJA durante el año en curso
'        If TieneConyugeDeducible(LiquidacionGanancia4ta.txtPuestoLaboral.Text) = False Then
'            .TextMatrix(3, 1) = De_Num_a_Tx_01(0)
'        Else
'            'Controlamos lo que debió liquidarse en concepto de Conyuge
'            dblImporteMensual = ImporteDeduccionPersonal("CONYUGE", DateSerial(Year(datFecha) - 1, 12, 31))
'            SQL = "Select CONYUGE From DEDUCCIONES4TACATEGORIA Where FECHA <= #" & Format(datFecha, "MM/DD/YYYY") & "# " _
'            & "And Year(Fecha) = '" & Year(datFecha) & "' Order by FECHA Asc"
'            If SQLNoMatch(SQL) = True Then
'                dblALiquidarse = dblImporteMensual * Month(datFecha)
'            Else
'                For i = 1 To Month(datFecha)
'                    datFechaControl = DateSerial(Year(datFecha), i, 1)
'                    datFechaControl = DateAdd("m", 1, datFechaControl)
'                    datFechaControl = DateAdd("d", -1, datFechaControl)
'                    dblImporteMensual = ImporteDeduccionPersonal("CONYUGE", datFechaControl)
'                    dblALiquidarse = dblALiquidarse + dblImporteMensual
'                Next i
'            End If
'            'Controlamos lo que realmente se liquidó en concepto de Conyuge
'            SQL = "Select  Sum(LIQUIDACIONGANANCIAS4TACATEGORIA.Conyuge) AS SumaDeConyuge " _
'            & "From CODIGOLIQUIDACIONES Inner Join LIQUIDACIONGANANCIAS4TACATEGORIA On CODIGOLIQUIDACIONES.Codigo = LIQUIDACIONGANANCIAS4TACATEGORIA.CodigoLiquidacion " _
'            & "Where PUESTOLABORAL = '" & LiquidacionGanancia4ta.txtPuestoLaboral.Text & "' " _
'            & "And Right(PERIODO,4) = '" & Right(LiquidacionGanancia4ta.txtPeriodo.Text, 4) & "' And CODIGO < '" & LiquidacionGanancia4ta.txtCodigoLiquidacion.Text & "'"
'            rstListadoSlave.Open SQL, dbSlave, adOpenDynamic, adLockOptimistic
'            If rstListadoSlave.BOF = False And IsNull(rstListadoSlave!SumaDeConyuge) = False Then
'                dblLiquidado = rstListadoSlave!SumaDeConyuge
'                dblAcumulado = dblAcumulado + dblLiquidado
'            Else
'                dblLiquidado = 0
'            End If
'            rstListadoSlave.Close
'            'Cargamos el Importe de la diferencia entre lo que debió liquidarse y lo que se liquidó en concepto de Conyuge
'            .TextMatrix(3, 1) = De_Num_a_Tx_01(dblALiquidarse - dblLiquidado)
'            dblTotal = dblTotal + De_Txt_a_Num_01(.TextMatrix(3, 1))
'        End If
'        dblALiquidarse = 0
'
'        .TextMatrix(4, 0) = " - Hijo/s" 'Tener en cuenta que no está previsto el caso de ALTA/BAJA durante el año en curso
'        intFamiliares = CantidadHijosDeducibles(LiquidacionGanancia4ta.txtPuestoLaboral.Text, datFecha)
'        If intFamiliares = 0 Then
'            .TextMatrix(4, 1) = De_Num_a_Tx_01(0)
'            'Controlamos la liquidación acumulada en concepto de Hijo
'            SQL = "Select  Sum(LIQUIDACIONGANANCIAS4TACATEGORIA.Hijo) AS SumaDeHijo " _
'            & "From CODIGOLIQUIDACIONES Inner Join LIQUIDACIONGANANCIAS4TACATEGORIA On CODIGOLIQUIDACIONES.Codigo = LIQUIDACIONGANANCIAS4TACATEGORIA.CodigoLiquidacion " _
'            & "Where PUESTOLABORAL = '" & LiquidacionGanancia4ta.txtPuestoLaboral.Text & "' " _
'            & "And Right(PERIODO,4) = '" & Right(LiquidacionGanancia4ta.txtPeriodo.Text, 4) & "' And CODIGO < '" & LiquidacionGanancia4ta.txtCodigoLiquidacion.Text & "'"
'            rstListadoSlave.Open SQL, dbSlave, adOpenDynamic, adLockOptimistic
'            If rstListadoSlave.BOF = False And IsNull(rstListadoSlave!SumaDeHijo) = False Then
'                dblAcumulado = dblAcumulado + rstListadoSlave!SumaDeHijo
'            End If
'            rstListadoSlave.Close
'        Else
'            'Controlamos lo que debió liquidarse en concepto de Hijo
'            dblImporteMensual = ImporteDeduccionPersonal("HIJO", DateSerial(Year(datFecha) - 1, 12, 31))
'            SQL = "Select HIJO From DEDUCCIONES4TACATEGORIA Where FECHA <= #" & Format(datFecha, "MM/DD/YYYY") & "# " _
'            & "And Year(Fecha) = '" & Year(datFecha) & "' Order by FECHA Desc"
'            If SQLNoMatch(SQL) = True Then
'                For i = 1 To Month(datFecha)
'                    intFamiliares = CantidadHijosDeducibles(LiquidacionGanancia4ta.txtPuestoLaboral.Text, DateSerial(Year(datFecha), i, 1))
'                    dblALiquidarse = dblALiquidarse + (dblImporteMensual * intFamiliares)
'                Next i
'            Else
'                For i = 1 To Month(datFecha)
'                    datFechaControl = DateSerial(Year(datFecha), i, 1)
'                    datFechaControl = DateAdd("m", 1, datFechaControl)
'                    datFechaControl = DateAdd("d", -1, datFechaControl)
'                    dblImporteMensual = ImporteDeduccionPersonal("HIJO", datFechaControl)
'                    intFamiliares = CantidadHijosDeducibles(LiquidacionGanancia4ta.txtPuestoLaboral.Text, DateSerial(Year(datFecha), i, 1))
'                    dblALiquidarse = dblALiquidarse + (dblImporteMensual * intFamiliares)
'                Next i
'            End If
'            'Controlamos lo que realmente se liquidó en concepto de Hijo
'            SQL = "Select  Sum(LIQUIDACIONGANANCIAS4TACATEGORIA.Hijo) AS SumaDeHijo " _
'            & "From CODIGOLIQUIDACIONES Inner Join LIQUIDACIONGANANCIAS4TACATEGORIA On CODIGOLIQUIDACIONES.Codigo = LIQUIDACIONGANANCIAS4TACATEGORIA.CodigoLiquidacion " _
'            & "Where PUESTOLABORAL = '" & LiquidacionGanancia4ta.txtPuestoLaboral.Text & "' " _
'            & "And Right(PERIODO,4) = '" & Right(LiquidacionGanancia4ta.txtPeriodo.Text, 4) & "' And CODIGO < '" & LiquidacionGanancia4ta.txtCodigoLiquidacion.Text & "'"
'            rstListadoSlave.Open SQL, dbSlave, adOpenDynamic, adLockOptimistic
'            If rstListadoSlave.BOF = False And IsNull(rstListadoSlave!SumaDeHijo) = False Then
'                dblLiquidado = rstListadoSlave!SumaDeHijo
'                dblAcumulado = dblAcumulado + dblLiquidado
'            Else
'                dblLiquidado = 0
'            End If
'            rstListadoSlave.Close
'            'Cargamos el Importe de la diferencia entre lo que debió liquidarse y lo que se liquidó en concepto de Hijo
'            .TextMatrix(4, 1) = De_Num_a_Tx_01(dblALiquidarse - dblLiquidado)
'            dblTotal = dblTotal + De_Txt_a_Num_01(.TextMatrix(4, 1))
'        End If
'        dblALiquidarse = 0
'
'        .TextMatrix(5, 0) = " - Otro/s" 'Tener en cuenta que no está previsto el caso de ALTA/BAJA durante el año en curso
'        intFamiliares = CantidadOtrasCargasFamiliaDeducibles(LiquidacionGanancia4ta.txtPuestoLaboral.Text)
'        If intFamiliares = 0 Then
'            .TextMatrix(5, 1) = De_Num_a_Tx_01(0)
'        Else
'            'Controlamos lo que debió liquidarse en concepto de Otras Cargas de Familia
'            dblImporteMensual = ImporteDeduccionPersonal("OTRASCARGASDEFAMILIA", DateSerial(Year(datFecha) - 1, 12, 31))
'            SQL = "Select OTRASCARGASDEFAMILIA From DEDUCCIONES4TACATEGORIA Where FECHA < #" & Format(datFecha, "MM/DD/YYYY") & "# " _
'            & "And Year(Fecha) = '" & Year(datFecha) & "' Order by FECHA Desc"
'            If SQLNoMatch(SQL) = True Then
'                dblALiquidarse = dblImporteMensual * Month(datFecha) * intFamiliares
'            Else
'                For i = 1 To Month(datFecha)
'                    datFechaControl = DateSerial(Year(datFecha), i, 1)
'                    datFechaControl = DateAdd("m", 1, datFechaControl)
'                    datFechaControl = DateAdd("d", -1, datFechaControl)
'                    dblImporteMensual = ImporteDeduccionPersonal("OTRASCARGASDEFAMILIA", datFechaControl)
'                    dblALiquidarse = dblALiquidarse + dblImporteMensual
'                Next i
'                dblALiquidarse = dblALiquidarse * intFamiliares
'            End If
'            'Controlamos lo que realmente se liquidó en concepto de Conyuge
'            SQL = "Select  Sum(LIQUIDACIONGANANCIAS4TACATEGORIA.OtrasCargasDeFamilia) AS SumaDeOtrasCargasDeFamilia " _
'            & "From CODIGOLIQUIDACIONES Inner Join LIQUIDACIONGANANCIAS4TACATEGORIA On CODIGOLIQUIDACIONES.Codigo = LIQUIDACIONGANANCIAS4TACATEGORIA.CodigoLiquidacion " _
'            & "Where PUESTOLABORAL = '" & LiquidacionGanancia4ta.txtPuestoLaboral.Text & "' " _
'            & "And Right(PERIODO,4) = '" & Right(LiquidacionGanancia4ta.txtPeriodo.Text, 4) & "' And CODIGO < '" & LiquidacionGanancia4ta.txtCodigoLiquidacion.Text & "'"
'            rstListadoSlave.Open SQL, dbSlave, adOpenDynamic, adLockOptimistic
'            If rstListadoSlave.BOF = False And IsNull(rstListadoSlave!SumaDeOtrasCargasDeFamilia) = False Then
'                dblLiquidado = rstListadoSlave!SumaDeOtrasCargasDeFamilia
'                dblAcumulado = dblAcumulado + dblLiquidado
'            Else
'                dblLiquidado = 0
'            End If
'            rstListadoSlave.Close
'            'Cargamos el Importe de la diferencia entre lo que debió liquidarse y lo que se liquidó en concepto de Otras Cargas
'            .TextMatrix(5, 1) = De_Num_a_Tx_01(dblALiquidarse - dblLiquidado)
'            dblTotal = dblTotal + De_Txt_a_Num_01(.TextMatrix(5, 1))
'        End If
'        dblALiquidarse = 0
'
'        .TextMatrix(6, 0) = "Deducción Especial"
'        'Controlamos lo que debió liquidarse en concepto de Deducción Especial
'        dblImporteMensual = ImporteDeduccionPersonal("DEDUCCIONESPECIAL", DateSerial(Year(datFecha) - 1, 12, 31))
'        SQL = "Select DEDUCCIONESPECIAL From DEDUCCIONES4TACATEGORIA Where FECHA <= #" & Format(datFecha, "MM/DD/YYYY") & "# " _
'        & "And Year(Fecha) = '" & Year(datFecha) & "' Order by FECHA Asc"
'        If SQLNoMatch(SQL) = True Then
'            dblALiquidarse = dblImporteMensual * Month(datFecha)
'        Else
'            For i = 1 To Month(datFecha)
'                datFechaControl = DateSerial(Year(datFecha), i, 1)
'                datFechaControl = DateAdd("m", 1, datFechaControl)
'                datFechaControl = DateAdd("d", -1, datFechaControl)
'                dblImporteMensual = ImporteDeduccionPersonal("DEDUCCIONESPECIAL", datFechaControl)
'                dblALiquidarse = dblALiquidarse + dblImporteMensual
'            Next i
'        End If
'        'Decreto Especiales del PEN
'        'Decreto 1006/13
'        If Year(datFecha) = 2013 And Month(datFecha) >= 7 Then
'            SQL = "Select * From LIQUIDACIONSUELDOS Where CODIGOLIQUIDACION = '0482' " _
'            & "And PUESTOLABORAL = '" & LiquidacionGanancia4ta.txtPuestoLaboral.Text & "'"
'            If SQLNoMatch(SQL) = False Then
'                'Buscamos el SAC Bruto
'                SQL = "Select * From LIQUIDACIONSUELDOS Where CODIGOLIQUIDACION = '0482' " _
'                & "And PUESTOLABORAL = '" & LiquidacionGanancia4ta.txtPuestoLaboral.Text & "' " _
'                & "And CODIGOCONCEPTO = '0150'"
'                If SQLNoMatch(SQL) = False Then
'                    rstListadoSlave.Open SQL, dbSlave, adOpenDynamic, adLockOptimistic
'                    dblImporteMensual = rstListadoSlave!Importe
'                    rstListadoSlave.Close
'                End If
'                'Buscamos la Retención Jubilación
'                SQL = "Select * From LIQUIDACIONSUELDOS Where CODIGOLIQUIDACION = '0482' " _
'                & "And PUESTOLABORAL = '" & LiquidacionGanancia4ta.txtPuestoLaboral.Text & "' " _
'                & "And CODIGOCONCEPTO = '0208'"
'                If SQLNoMatch(SQL) = False Then
'                    rstListadoSlave.Open SQL, dbSlave, adOpenDynamic, adLockOptimistic
'                    dblImporteMensual = dblImporteMensual - rstListadoSlave!Importe
'                    rstListadoSlave.Close
'                End If
'                'Buscamos la Retención Obra Social
'                SQL = "Select * From LIQUIDACIONSUELDOS Where CODIGOLIQUIDACION = '0482' " _
'                & "And PUESTOLABORAL = '" & LiquidacionGanancia4ta.txtPuestoLaboral.Text & "' " _
'                & "And CODIGOCONCEPTO = '0212'"
'                If SQLNoMatch(SQL) = False Then
'                    rstListadoSlave.Open SQL, dbSlave, adOpenDynamic, adLockOptimistic
'                    dblImporteMensual = dblImporteMensual - rstListadoSlave!Importe
'                    rstListadoSlave.Close
'                End If
'                'Buscamos la Retención Aporte Sindical
'                SQL = "Select * from LIQUIDACIONSUELDOS Where " _
'                & "((PUESTOLABORAL = '" & LiquidacionGanancia4ta.txtPuestoLaboral.Text & "' " _
'                & "And CODIGOLIQUIDACION = '0482' And CODIGOCONCEPTO = '0219') Or " _
'                & "(PUESTOLABORAL = '" & LiquidacionGanancia4ta.txtPuestoLaboral.Text & "' " _
'                & "And CODIGOLIQUIDACION = '0482' And CODIGOCONCEPTO = '0227'))"
'                If SQLNoMatch(SQL) = False Then
'                    rstListadoSlave.Open SQL, dbSlave, adOpenDynamic, adLockOptimistic
'                    dblImporteMensual = dblImporteMensual - rstListadoSlave!Importe
'                    rstListadoSlave.Close
'                End If
'                'Buscamos el Pluriempleo
'                SQL = "Select * from LIQUIDACIONGANANCIAS4TACATEGORIA Where " _
'                & "(PUESTOLABORAL = '" & LiquidacionGanancia4ta.txtPuestoLaboral.Text & "' " _
'                & "And CODIGOLIQUIDACION = '0482')"
'                If SQLNoMatch(SQL) = False Then
'                    rstListadoSlave.Open SQL, dbSlave, adOpenDynamic, adLockOptimistic
'                    dblImporteMensual = dblImporteMensual + rstListadoSlave!Pluriempleo
'                    rstListadoSlave.Close
'                End If
'                dblALiquidarse = dblALiquidarse + dblImporteMensual
'            End If
'        End If
'        'Decreto 2354/14
'        If Year(datFecha) = 2014 And Month(datFecha) >= 12 Then
'            SQL = "Select * From LIQUIDACIONSUELDOS Where CODIGOLIQUIDACION = '0524' " _
'            & "And PUESTOLABORAL = '" & LiquidacionGanancia4ta.txtPuestoLaboral.Text & "'"
'            If SQLNoMatch(SQL) = False Then
'                'Buscamos el SAC Bruto
'                SQL = "Select * From LIQUIDACIONSUELDOS Where CODIGOLIQUIDACION = '0524' " _
'                & "And PUESTOLABORAL = '" & LiquidacionGanancia4ta.txtPuestoLaboral.Text & "' " _
'                & "And CODIGOCONCEPTO = '0150'"
'                If SQLNoMatch(SQL) = False Then
'                    rstListadoSlave.Open SQL, dbSlave, adOpenDynamic, adLockOptimistic
'                    dblImporteMensual = rstListadoSlave!Importe
'                    rstListadoSlave.Close
'                End If
'                'Buscamos la Retención Jubilación
'                SQL = "Select * From LIQUIDACIONSUELDOS Where CODIGOLIQUIDACION = '0524' " _
'                & "And PUESTOLABORAL = '" & LiquidacionGanancia4ta.txtPuestoLaboral.Text & "' " _
'                & "And CODIGOCONCEPTO = '0208'"
'                If SQLNoMatch(SQL) = False Then
'                    rstListadoSlave.Open SQL, dbSlave, adOpenDynamic, adLockOptimistic
'                    dblImporteMensual = dblImporteMensual - rstListadoSlave!Importe
'                    rstListadoSlave.Close
'                End If
'                'Buscamos la Retención Obra Social
'                SQL = "Select * From LIQUIDACIONSUELDOS Where CODIGOLIQUIDACION = '0524' " _
'                & "And PUESTOLABORAL = '" & LiquidacionGanancia4ta.txtPuestoLaboral.Text & "' " _
'                & "And CODIGOCONCEPTO = '0212'"
'                If SQLNoMatch(SQL) = False Then
'                    rstListadoSlave.Open SQL, dbSlave, adOpenDynamic, adLockOptimistic
'                    dblImporteMensual = dblImporteMensual - rstListadoSlave!Importe
'                    rstListadoSlave.Close
'                End If
'                'Buscamos la Retención Aporte Sindical
'                SQL = "Select * from LIQUIDACIONSUELDOS Where " _
'                & "((PUESTOLABORAL = '" & LiquidacionGanancia4ta.txtPuestoLaboral.Text & "' " _
'                & "And CODIGOLIQUIDACION = '0524' And CODIGOCONCEPTO = '0219') Or " _
'                & "(PUESTOLABORAL = '" & LiquidacionGanancia4ta.txtPuestoLaboral.Text & "' " _
'                & "And CODIGOLIQUIDACION = '0524' And CODIGOCONCEPTO = '0227'))"
'                If SQLNoMatch(SQL) = False Then
'                    rstListadoSlave.Open SQL, dbSlave, adOpenDynamic, adLockOptimistic
'                    dblImporteMensual = dblImporteMensual - rstListadoSlave!Importe
'                    rstListadoSlave.Close
'                End If
'                'Buscamos el Pluriempleo
'                SQL = "Select * from LIQUIDACIONGANANCIAS4TACATEGORIA Where " _
'                & "(PUESTOLABORAL = '" & LiquidacionGanancia4ta.txtPuestoLaboral.Text & "' " _
'                & "And CODIGOLIQUIDACION = '0524')"
'                If SQLNoMatch(SQL) = False Then
'                    rstListadoSlave.Open SQL, dbSlave, adOpenDynamic, adLockOptimistic
'                    dblImporteMensual = dblImporteMensual + rstListadoSlave!Pluriempleo
'                    rstListadoSlave.Close
'                End If
'                dblALiquidarse = dblALiquidarse + dblImporteMensual
'            End If
'        End If
'        'Controlamos lo que realmente se liquidó en concepto de Deducción Especial
'        SQL = "Select  Sum(LIQUIDACIONGANANCIAS4TACATEGORIA.DeduccionEspecial) AS SumaDeDeduccionEspecial " _
'        & "From CODIGOLIQUIDACIONES Inner Join LIQUIDACIONGANANCIAS4TACATEGORIA On CODIGOLIQUIDACIONES.Codigo = LIQUIDACIONGANANCIAS4TACATEGORIA.CodigoLiquidacion " _
'        & "Where PUESTOLABORAL = '" & LiquidacionGanancia4ta.txtPuestoLaboral.Text & "' " _
'        & "And Right(PERIODO,4) = '" & Right(LiquidacionGanancia4ta.txtPeriodo.Text, 4) & "' And CODIGO < '" & LiquidacionGanancia4ta.txtCodigoLiquidacion.Text & "'"
'        rstListadoSlave.Open SQL, dbSlave, adOpenDynamic, adLockOptimistic
'        If rstListadoSlave.BOF = False And IsNull(rstListadoSlave!SumaDeDeduccionEspecial) = False Then
'            dblLiquidado = rstListadoSlave!SumaDeDeduccionEspecial
'            dblAcumulado = dblAcumulado + dblLiquidado
'        Else
'            dblLiquidado = 0
'        End If
'        rstListadoSlave.Close
'        'Cargamos el Importe de la diferencia entre lo que debió liquidarse y lo que se liquidó en concepto de Deducción Especial
'        .TextMatrix(6, 1) = De_Num_a_Tx_01(dblALiquidarse - dblLiquidado)
'        dblTotal = dblTotal + De_Txt_a_Num_01(.TextMatrix(6, 1))
'        dblALiquidarse = 0
'
'        'Cargamos el Total de Deducciones
'        .TextMatrix(7, 0) = "Total Mensual"
'        dblTotal = dblTotal * (-1)
'        .TextMatrix(7, 1) = De_Num_a_Tx_01(dblTotal)
'
'        .TextMatrix(8, 0) = "Total Acumulado"
'        dblTotal = dblTotal * (-1)
'        dblAcumulado = (dblAcumulado + dblTotal) * (-1)
'        .TextMatrix(8, 1) = De_Num_a_Tx_01(dblAcumulado)
'
'    End With
'
'    Set rstListadoSlave = Nothing
'    SQL = ""
'    datFecha = 0
'    dat24Years = 0
'    datFechaControl = 0
'    dblALiquidarse = 0
'    dblAcumulado = 0
'    dblLiquidado = 0
'    dblTotal = 0
'    intFamiliares = 0
'
'End Sub

Public Sub CargardgDeduccionesPersonalesLG4taIndividual(Concepto As String, Importe As Double)
    
    Dim i As Integer
    
    With LiquidacionGanancia4ta.dgDeduccionesPersonales
        'Determinamos el número de filas
        i = .Rows
        'Verificamos si tiene más de dos filas
        If i > 2 Then
            'Agregamos una fila
            .Rows = i + 1
            i = .Rows
        Else
            'Verificamos si la segunda fila tiene datos
            If .TextMatrix(1, 0) <> "" Then
                'Si ya tiene datos, agregamos una fila
                .Rows = i + 1
                i = .Rows
            End If
        End If
        'Le restamos 1 a i porque el indíce de filas arranca de 0
        i = i - 1
        .TextMatrix(i, 0) = Concepto
        .TextMatrix(i, 1) = De_Num_a_Tx_01(Round(Importe, 2))
    End With
    
    i = 0
    
End Sub


Public Sub ConfigurardgDeduccionesGeneralesLG4ta()
    
    With LiquidacionGanancia4ta.dgDeduccionesGenerales
        .Clear
        .Cols = 2
        .Rows = 2
        .TextMatrix(0, 0) = "Concepto"
        .TextMatrix(0, 1) = "Importe"
        .ColWidth(0) = 1600
        .ColWidth(1) = 800
        .FixedCols = 0
        .FocusRect = flexFocusHeavy
        .HighLight = flexHighlightWithFocus
        .SelectionMode = flexSelectionByRow
        .AllowUserResizing = flexResizeColumns
        .ColAlignment(0) = 1
        .ColAlignment(1) = 7
    End With
    
End Sub

'Public Sub CargardgDeduccionesGeneralesLG4ta()
'
'    Dim SQL As String
'    Dim datFecha As Date
'    Dim dblLiquidado As Double
'    Dim dblALiquidarse As Double
'    Dim dblAcumulado As Double
'    Dim dblGananciaNeta As Double
'    Dim dblTotal As Double
'
'    With LiquidacionGanancia4ta
'        datFecha = DateTime.DateSerial(Right(.txtPeriodo.Text, 4), Left(.txtPeriodo.Text, 2), 1)
'        datFecha = DateAdd("m", 1, datFecha)
'        datFecha = DateAdd("d", -1, datFecha)
'    End With
'
'    Set rstListadoSlave = New ADODB.Recordset
'    With LiquidacionGanancia4ta.dgDeduccionesGenerales
'        .Rows = 7
'        .RowHeight(0) = 300
'
'        .TextMatrix(1, 0) = "Servicio Doméstico"
'        'Controlamos lo que debió liquidarse en concepto de Servicio Doméstico
'        dblALiquidarse = ImporteDeduccionGeneral(LiquidacionGanancia4ta.txtPuestoLaboral.Text, "SERVICIODOMESTICO", datFecha)
'        dblALiquidarse = dblALiquidarse * Month(datFecha)
'        'Controlamos lo que realmente se liquidó en concepto de Servicio Doméstico
'        SQL = "Select  Sum(LIQUIDACIONGANANCIAS4TACATEGORIA.ServicioDomestico) AS SumaDeServicioDomestico " _
'        & "From CODIGOLIQUIDACIONES Inner Join LIQUIDACIONGANANCIAS4TACATEGORIA On CODIGOLIQUIDACIONES.Codigo = LIQUIDACIONGANANCIAS4TACATEGORIA.CodigoLiquidacion " _
'        & "Where PUESTOLABORAL = '" & LiquidacionGanancia4ta.txtPuestoLaboral.Text & "' " _
'        & "And Right(PERIODO,4) = '" & Right(LiquidacionGanancia4ta.txtPeriodo.Text, 4) & "' And CODIGO < '" & LiquidacionGanancia4ta.txtCodigoLiquidacion.Text & "'"
'        rstListadoSlave.Open SQL, dbSlave, adOpenDynamic, adLockOptimistic
'        If rstListadoSlave.BOF = False And IsNull(rstListadoSlave!SumaDeServicioDomestico) = False Then
'            dblLiquidado = rstListadoSlave!SumaDeServicioDomestico
'            dblAcumulado = dblAcumulado + dblLiquidado
'        Else
'            dblLiquidado = 0
'        End If
'        rstListadoSlave.Close
'        'Cargamos el Importe de la diferencia entre lo que debió liquidarse y lo que se liquidó en concepto de Servicio Doméstico
'        .TextMatrix(1, 1) = De_Num_a_Tx_01(dblALiquidarse - dblLiquidado, , 2)
'        dblTotal = dblTotal + De_Txt_a_Num_01(.TextMatrix(1, 1))
'        dblALiquidarse = 0
'
'        .TextMatrix(2, 0) = "Seguro de Vida"
'        'Controlamos lo que debió liquidarse en concepto de Seguro de Vida
'        dblALiquidarse = ImporteDeduccionGeneral(LiquidacionGanancia4ta.txtPuestoLaboral.Text, "SEGURODEVIDA", datFecha, , De_Txt_a_Num_01(LiquidacionGanancia4ta.txtSeguroOptativo.Text))
'        dblALiquidarse = dblALiquidarse * Month(datFecha)
'        'Controlamos lo que realmente se liquidó en concepto de Seguro de Vida
'        SQL = "Select  Sum(LIQUIDACIONGANANCIAS4TACATEGORIA.SeguroDeVidaOptativo) AS SumaDeSeguroDeVida " _
'        & "From CODIGOLIQUIDACIONES Inner Join LIQUIDACIONGANANCIAS4TACATEGORIA On CODIGOLIQUIDACIONES.Codigo = LIQUIDACIONGANANCIAS4TACATEGORIA.CodigoLiquidacion " _
'        & "Where PUESTOLABORAL = '" & LiquidacionGanancia4ta.txtPuestoLaboral.Text & "' " _
'        & "And Right(PERIODO,4) = '" & Right(LiquidacionGanancia4ta.txtPeriodo.Text, 4) & "' And CODIGO < '" & LiquidacionGanancia4ta.txtCodigoLiquidacion.Text & "'"
'        rstListadoSlave.Open SQL, dbSlave, adOpenDynamic, adLockOptimistic
'        If rstListadoSlave.BOF = False And IsNull(rstListadoSlave!SumaDeSeguroDeVida) = False Then
'            dblLiquidado = rstListadoSlave!SumaDeSeguroDeVida
'            dblAcumulado = dblAcumulado + dblLiquidado
'        Else
'            dblLiquidado = 0
'        End If
'        rstListadoSlave.Close
'        If dblALiquidarse > 0 Then
'            'Cargamos el Importe de la diferencia entre lo que debió liquidarse y lo que se liquidó en concepto de Seguro De Vida _
'            teniendo en cuenta lo que se consideró como deducción
'            dblALiquidarse = (dblALiquidarse - dblLiquidado)
'            .TextMatrix(2, 1) = De_Num_a_Tx_01(dblALiquidarse - De_Txt_a_Num_01(LiquidacionGanancia4ta.txtSeguroOptativo.Text), , 2)
'        Else
'            .TextMatrix(2, 1) = De_Num_a_Tx_01(0, , 2)
'        End If
'        dblTotal = dblTotal + De_Txt_a_Num_01(.TextMatrix(2, 1))
'        dblALiquidarse = 0
'
'        .TextMatrix(3, 0) = "Cuota Médica Asist."
'        'Controlamos lo que debió liquidarse en concepto de Cuota Médico Asistencial
'        dblGananciaNeta = (De_Txt_a_Num_01(LiquidacionGanancia4ta.txtGananciaNeta.Text) - dblAcumulado _
'        - De_Txt_a_Num_01(.TextMatrix(2, 1)) - De_Txt_a_Num_01(.TextMatrix(1, 1)) - De_Txt_a_Num_01(LiquidacionGanancia4ta.txtSeguroOptativo.Text))
'        dblALiquidarse = ImporteDeduccionGeneral(LiquidacionGanancia4ta.txtPuestoLaboral.Text, "CUOTAMEDICOASISTENCIAL", datFecha, dblGananciaNeta)
'        'Controlamos lo que realmente se liquidó en concepto de Mínimo no Imponible
'        SQL = "Select  Sum(LIQUIDACIONGANANCIAS4TACATEGORIA.CuotaMedicoAsistencial) AS SumaDeCuotaMedicoAsistencial " _
'        & "From CODIGOLIQUIDACIONES Inner Join LIQUIDACIONGANANCIAS4TACATEGORIA On CODIGOLIQUIDACIONES.Codigo = LIQUIDACIONGANANCIAS4TACATEGORIA.CodigoLiquidacion " _
'        & "Where PUESTOLABORAL = '" & LiquidacionGanancia4ta.txtPuestoLaboral.Text & "' " _
'        & "And Right(PERIODO,4) = '" & Right(LiquidacionGanancia4ta.txtPeriodo.Text, 4) & "' And CODIGO < '" & LiquidacionGanancia4ta.txtCodigoLiquidacion.Text & "'"
'        rstListadoSlave.Open SQL, dbSlave, adOpenDynamic, adLockOptimistic
'        If rstListadoSlave.BOF = False And IsNull(rstListadoSlave!SumaDeCuotaMedicoAsistencial) = False Then
'            dblLiquidado = rstListadoSlave!SumaDeCuotaMedicoAsistencial
'            dblAcumulado = dblAcumulado + dblLiquidado
'        Else
'            dblLiquidado = 0
'        End If
'        rstListadoSlave.Close
'        'Cargamos el Importe de la diferencia entre lo que debió liquidarse y lo que se liquidó en concepto de Cuota Médico Asistencial
'        .TextMatrix(3, 1) = De_Num_a_Tx_01(dblALiquidarse - dblLiquidado, , 2)
'        dblTotal = dblTotal + De_Txt_a_Num_01(.TextMatrix(3, 1))
'        dblALiquidarse = 0
'
'        .TextMatrix(4, 0) = "Donaciones"
'        'Controlamos lo que debió liquidarse en concepto de Donaciones
'        If dblGananciaNeta = 0 Then
'            dblGananciaNeta = (De_Txt_a_Num_01(LiquidacionGanancia4ta.txtGananciaNeta.Text) - dblAcumulado _
'            - De_Txt_a_Num_01(.TextMatrix(2, 1)) - De_Txt_a_Num_01(.TextMatrix(1, 1)) - De_Txt_a_Num_01(LiquidacionGanancia4ta.txtSeguroOptativo.Text))
'        End If
'        dblALiquidarse = ImporteDeduccionGeneral(LiquidacionGanancia4ta.txtPuestoLaboral.Text, "DONACIONES", datFecha, dblGananciaNeta)
'        'Controlamos lo que realmente se liquidó en concepto de Donaciones
'        SQL = "Select  Sum(LIQUIDACIONGANANCIAS4TACATEGORIA.Donaciones) AS SumaDeDonaciones " _
'        & "From CODIGOLIQUIDACIONES Inner Join LIQUIDACIONGANANCIAS4TACATEGORIA On CODIGOLIQUIDACIONES.Codigo = LIQUIDACIONGANANCIAS4TACATEGORIA.CodigoLiquidacion " _
'        & "Where PUESTOLABORAL = '" & LiquidacionGanancia4ta.txtPuestoLaboral.Text & "' " _
'        & "And Right(PERIODO,4) = '" & Right(LiquidacionGanancia4ta.txtPeriodo.Text, 4) & "' And CODIGO < '" & LiquidacionGanancia4ta.txtCodigoLiquidacion.Text & "'"
'        rstListadoSlave.Open SQL, dbSlave, adOpenDynamic, adLockOptimistic
'        If rstListadoSlave.BOF = False And IsNull(rstListadoSlave!SumaDeDonaciones) = False Then
'            dblLiquidado = rstListadoSlave!SumaDeDonaciones
'            dblAcumulado = dblAcumulado + dblLiquidado
'        Else
'            dblLiquidado = 0
'        End If
'        rstListadoSlave.Close
'        'Cargamos el Importe de la diferencia entre lo que debió liquidarse y lo que se liquidó en concepto de Donaciones
'        .TextMatrix(4, 1) = De_Num_a_Tx_01(dblALiquidarse - dblLiquidado, , 2)
'        dblTotal = dblTotal + De_Txt_a_Num_01(.TextMatrix(4, 1))
'        dblALiquidarse = 0
'
'        .TextMatrix(5, 0) = "Total Mensual"
'        dblTotal = dblTotal * (-1)
'        .TextMatrix(5, 1) = De_Num_a_Tx_01(dblTotal, , 2)
'
'        .TextMatrix(6, 0) = "Total Acumulado"
'        dblTotal = dblTotal * (-1)
'        dblAcumulado = (dblAcumulado + dblTotal) * (-1)
'        .TextMatrix(6, 1) = De_Num_a_Tx_01(dblAcumulado, , 2)
'    End With
'
'    Set rstListadoSlave = Nothing
'    SQL = ""
'    datFecha = 0
'    dblAcumulado = 0
'    dblALiquidarse = 0
'    dblLiquidado = 0
'    dblTotal = 0
'    dblGananciaNeta = 0
'
'End Sub

Public Sub CargardgDeduccionesGeneralesLG4taIndividual(Concepto As String, Importe As Double)
    
    Dim i As Integer
    
    With LiquidacionGanancia4ta.dgDeduccionesGenerales
        'Determinamos el número de filas
        i = .Rows
        'Verificamos si tiene más de dos filas
        If i > 2 Then
            'Agregamos una fila
            .Rows = i + 1
            i = .Rows
        Else
            'Verificamos si la segunda fila tiene datos
            If .TextMatrix(1, 0) <> "" Then
                'Si ya tiene datos, agregamos una fila
                .Rows = i + 1
                i = .Rows
            End If
        End If
        'Le restamos 1 a i porque el indíce de filas arranca de 0
        i = i - 1
        .TextMatrix(i, 0) = Concepto
        .TextMatrix(i, 1) = De_Num_a_Tx_01(Round(Importe, 2))
    End With
    
    i = 0
    
End Sub

Public Sub ConfigurardgLiquidacionSISPER()
    
    With ListadoLiquidacionesSISPER.dgCodigoLiquidacion
        .Clear
        .Cols = 3
        .Rows = 2
        .TextMatrix(0, 0) = "Código"
        .TextMatrix(0, 1) = "Descripción"
        .TextMatrix(0, 2) = "Nro Agentes"
        .ColWidth(0) = 600
        .ColWidth(1) = 2000
        .ColWidth(2) = 1200
        .FixedCols = 0
        .FocusRect = flexFocusHeavy
        .HighLight = flexHighlightWithFocus
        .SelectionMode = flexSelectionByRow
        .AllowUserResizing = flexResizeColumns
        .ColAlignment(0) = 4
        .ColAlignment(1) = 1
        .ColAlignment(2) = 7
    
    End With
End Sub

Public Sub CargardgLiquidacionSISPER()
    
    Dim i As Integer
    Dim SQL As String
    
    i = 0
    ListadoLiquidacionesSISPER.dgCodigoLiquidacion.Rows = 2
    Set rstListadoSlave = New ADODB.Recordset
    rstListadoSlave.Open "Select * From CODIGOLIQUIDACIONES Order By CODIGO Desc", dbSlave, adOpenDynamic, adLockOptimistic
    With ListadoLiquidacionesSISPER.dgCodigoLiquidacion
        If rstListadoSlave.BOF = False Then
            rstListadoSlave.MoveFirst
            While rstListadoSlave.EOF = False
                i = i + 1
                .RowHeight(i) = 300
                .TextMatrix(i, 0) = rstListadoSlave!Codigo
                .TextMatrix(i, 1) = rstListadoSlave!Descripcion
                SQL = "Select PUESTOLABORAL From LIQUIDACIONSUELDOS Where CODIGOLIQUIDACION = " & "'" & rstListadoSlave!Codigo & "' Group by PUESTOLABORAL"
                If SQLNoMatch(SQL) = True Then
                   .TextMatrix(i, 2) = 0
                Else
                    SQL = "Select COUNT(PUESTOLABORAL) As NumeroAgentes From (" & SQL & ")"
                    Set rstBuscarSlave = New ADODB.Recordset
                    rstBuscarSlave.Open SQL, dbSlave, adOpenDynamic, adLockOptimistic
                    .TextMatrix(i, 2) = rstBuscarSlave!NumeroAgentes
                    rstBuscarSlave.Close
                    Set rstBuscarSlave = Nothing
                End If
                rstListadoSlave.MoveNext
                .Rows = .Rows + 1
            Wend
        End If
        .Rows = .Rows - 1
        .SetFocus
    End With
    rstListadoSlave.Close
    Set rstListadoSlave = Nothing
    
End Sub

Public Sub ConfigurardgComprobantesSIIF()

    With ListadoComprobantesSIIF.dgListadoComprobante
        .Clear
        .Cols = 5
        .Rows = 2
        .TextMatrix(0, 0) = "Comprobante"
        .TextMatrix(0, 1) = "Fecha"
        .TextMatrix(0, 2) = "Tipo Imputacion"
        .TextMatrix(0, 3) = "N° Agentes"
        .TextMatrix(0, 4) = "Importe"
        .ColWidth(0) = 1100
        .ColWidth(1) = 1000
        .ColWidth(2) = 3000
        .ColWidth(3) = 1000
        .ColWidth(4) = 1000
        .FixedCols = 0
        .FixedRows = 1
        .FocusRect = flexFocusHeavy
        .HighLight = flexHighlightWithFocus
        .SelectionMode = flexSelectionByRow
        .ColAlignment(0) = 1
        .ColAlignment(1) = 7
        .ColAlignment(2) = 1
        .ColAlignment(3) = 7
        .ColAlignment(4) = 7
    End With
    
End Sub

Public Sub CargardgComprobantesSIIF(Optional ByVal Comprobante As String = "", Optional ByVal Fecha As String = "", Optional ByVal BuscarComprobante As String = 0)

    Dim i As String
    Dim FilaBuscada As String
    Dim SQL As String
    Dim strTipo As String
    
    FilaBuscada = 0
    i = 0
    ListadoComprobantesSIIF.dgListadoComprobante.Rows = 2
    
    Select Case Len(Fecha)
    Case Is = 0
        SQL = "Select COMPROBANTE, FECHA, TIPO, Sum(MONTOBRUTO) As MB, Count(PROVEEDOR) As TP" _
        & " From LIQUIDACIONHONORARIOS" _
        & " Where Left(COMPROBANTE,6) <> '" & "NoSIIF" & "'" _
        & " Group By COMPROBANTE, FECHA, TIPO" _
        & " Order By Fecha DESC,Comprobante DESC"
    Case Is = 4
        SQL = "Select COMPROBANTE, FECHA, TIPO, Sum(MONTOBRUTO) As MB, Count(PROVEEDOR) As TP" _
        & " From LIQUIDACIONHONORARIOS" _
        & " Where Left(COMPROBANTE,6) <> '" & "NoSIIF" & "'" _
        & " And Year(FECHA) = '" & Fecha & "'" _
        & " Group By COMPROBANTE, FECHA, TIPO" _
        & " Order By Fecha DESC,Comprobante DESC"
'    Case Is = 7
'        rstListadoIcaro.Open "SELECT * FROM CARGA Where Year(Fecha) ='" _
'        & Right(Fecha, 4) & "' And Format(Month(Fecha),'00') ='" & Left(Fecha, 2) _
'        & "' ORDER BY Fecha DESC,Comprobante DESC", dbIcaro, adOpenDynamic, adLockOptimistic
'    Case Is = 10
'        rstListadoIcaro.Open "SELECT * FROM CARGA Where Year(Fecha) ='" _
'        & Right(Fecha, 4) & "' And Format(Month(Fecha),'00') ='" & Mid(Fecha, 4, 2) _
'        & "' And Format(Day(Fecha),'00') ='" & Left(Fecha, 2) _
'        & "' ORDER BY Fecha DESC,Comprobante DESC", dbIcaro, adOpenDynamic, adLockOptimistic
    End Select
    
    If SQLNoMatch(SQL) = False Then
        Set rstListadoSlave = New ADODB.Recordset
        rstListadoSlave.Open SQL, dbSlave, adOpenDynamic, adLockOptimistic
        With ListadoComprobantesSIIF.dgListadoComprobante
            If rstListadoSlave.BOF = False Then
                rstListadoSlave.MoveFirst
                While rstListadoSlave.EOF = False
                    i = i + 1
                    .RowHeight(i) = 300
                    If BuscarComprobante <> "0" Then
                        If BuscarComprobante = rstListadoSlave!Comprobante Then
                            FilaBuscada = i
                        End If
                    End If
                    .TextMatrix(i, 0) = rstListadoSlave!Comprobante
                    .TextMatrix(i, 1) = rstListadoSlave!Fecha
                    Select Case rstListadoSlave!Tipo
                    Case "H"
                        strTipo = "Honorarios"
                    Case "C"
                        strTipo = "Comisiones"
                    Case "E"
                        strTipo = "Horas Extras"
                    Case "L"
                        strTipo = "Licencia"
                    End Select
                    .TextMatrix(i, 2) = strTipo
                    .TextMatrix(i, 3) = rstListadoSlave!TP
                    .TextMatrix(i, 4) = FormatNumber(rstListadoSlave!MB, 2)
                    rstListadoSlave.MoveNext
                    .Rows = .Rows + 1
                Wend
            End If
            .Rows = .Rows - 1
            If FilaBuscada <> "0" Then
                .TopRow = FilaBuscada
                .Row = FilaBuscada
            End If
        End With
        rstListadoSlave.Close
        Set rstListadoSlave = Nothing
    End If
    
End Sub

Public Sub ConfigurardgImputacion()

    With ListadoComprobantesSIIF.dgImputacion
        .Clear
        .Cols = 3
        .Rows = 2
        .TextMatrix(0, 0) = "Imputacion"
        .TextMatrix(0, 1) = "Partida"
        .TextMatrix(0, 2) = "Importe"
        .ColWidth(0) = 1100
        .ColWidth(1) = 1000
        .ColWidth(2) = 1000
        .FixedCols = 0
        .FixedRows = 1
        .FocusRect = flexFocusHeavy
        .HighLight = flexHighlightWithFocus
        .SelectionMode = flexSelectionByRow
        .ColAlignment(0) = 1
        .ColAlignment(1) = 1
        .ColAlignment(2) = 7
    End With
    
End Sub

Public Sub CargardgImputacion(Numero As String)

    Dim i As Integer
    Dim SQL As String
    Dim SumaTotal As Double
    
    i = 0
    ListadoComprobantesSIIF.dgImputacion.Rows = 2
    
    SQL = "Select ACTIVIDAD, PARTIDA, SUM(MONTOBRUTO) As SMB" _
    & " FROM LIQUIDACIONHONORARIOS" _
    & " Where COMPROBANTE = " & "'" & Numero & "'" _
    & " Group by ACTIVIDAD, PARTIDA"
    If SQLNoMatch(SQL) = False Then
        Set rstListadoSlave = New ADODB.Recordset
        rstListadoSlave.Open SQL, dbSlave, adOpenDynamic, adLockOptimistic
        With ListadoComprobantesSIIF.dgImputacion
            If rstListadoSlave.EOF = False Then
                rstListadoSlave.MoveFirst
                While rstListadoSlave.EOF = False
                    i = i + 1
                    .RowHeight(i) = 300
                    .TextMatrix(i, 0) = rstListadoSlave!ACTIVIDAD
                    .TextMatrix(i, 1) = rstListadoSlave!PARTIDA
                    .TextMatrix(i, 2) = FormatNumber(rstListadoSlave!SMB, 2)
                    SumaTotal = SumaTotal + rstListadoSlave!SMB
                    rstListadoSlave.MoveNext
                    .Rows = .Rows + 1
                Wend
                ListadoComprobantesSIIF.txtTotalImputacion.Text = FormatNumber(SumaTotal, 2)
            End If
            .Rows = .Rows - 1
        End With
        rstListadoSlave.Close
        Set rstListadoSlave = Nothing
        SumaTotal = 0
    End If
    
End Sub

Public Sub ConfigurardgRetencion()

    With ListadoComprobantesSIIF.dgRetencion
        .Clear
        .Cols = 2
        .Rows = 2
        .TextMatrix(0, 0) = "Código"
        .TextMatrix(0, 1) = "Importe"
        .ColWidth(0) = 2100
        .ColWidth(1) = 1000
        .FixedCols = 0
        .FixedRows = 1
        .FocusRect = flexFocusHeavy
        .HighLight = flexHighlightWithFocus
        .SelectionMode = flexSelectionByRow
        .ColAlignment(0) = 1
        .ColAlignment(1) = 7
    End With
    
End Sub

Public Sub CargardgRetencion(Numero As String)
    
    Dim i As Integer
    Dim SQL As String
    Dim SumaTotal As Double
    
    i = 0
    ListadoComprobantesSIIF.dgRetencion.Rows = 2
    
    SQL = "Select SUM(IIBB) As SIIBB, SUM(Sellos) As SSellos, SUM(LibramientoPago) As SLibramientoPago," _
    & " SUM(Seguro) As SSeguro, SUM(Descuento) As SDescuento" _
    & " FROM LIQUIDACIONHONORARIOS" _
    & " Where COMPROBANTE = " & "'" & Numero & "'"
    If SQLNoMatch(SQL) = False Then
        Set rstListadoSlave = New ADODB.Recordset
        rstListadoSlave.Open SQL, dbSlave, adOpenDynamic, adLockOptimistic
        With ListadoComprobantesSIIF.dgRetencion
            If rstListadoSlave.EOF = False Then
                rstListadoSlave.MoveFirst
                i = i + 1
                .RowHeight(i) = 300
                .TextMatrix(i, 0) = "101"
                .TextMatrix(i, 1) = FormatNumber(rstListadoSlave!SIIBB, 2)
                SumaTotal = SumaTotal + rstListadoSlave!SIIBB
                .Rows = .Rows + 1
                i = i + 1
                .RowHeight(i) = 300
                .TextMatrix(i, 0) = "102"
                .TextMatrix(i, 1) = FormatNumber(rstListadoSlave!SSellos, 2)
                SumaTotal = SumaTotal + rstListadoSlave!SSellos
                .Rows = .Rows + 1
                i = i + 1
                .RowHeight(i) = 300
                .TextMatrix(i, 0) = "104"
                .TextMatrix(i, 1) = FormatNumber(rstListadoSlave!SLibramientoPago, 2)
                SumaTotal = SumaTotal + rstListadoSlave!SLibramientoPago
                .Rows = .Rows + 1
                i = i + 1
                .RowHeight(i) = 300
                .TextMatrix(i, 0) = "413"
                .TextMatrix(i, 1) = FormatNumber(rstListadoSlave!SSeguro, 2)
                SumaTotal = SumaTotal + rstListadoSlave!SSeguro
                .Rows = .Rows + 1
                i = i + 1
                .RowHeight(i) = 300
                .TextMatrix(i, 0) = "999"
                .TextMatrix(i, 1) = FormatNumber(rstListadoSlave!SDescuento, 2)
                SumaTotal = SumaTotal + rstListadoSlave!SDescuento
                ListadoComprobantesSIIF.txtTotalRetencion.Text = FormatNumber(SumaTotal, 2)
            End If
        End With
        rstListadoSlave.Close
        Set rstListadoSlave = Nothing
        SumaTotal = 0
    End If
       
End Sub

Public Sub ConfigurardgAutocarga()
    
    With Autocarga.dgListadoAutocarga
        .Clear
        .Cols = 6
        .Rows = 2
        .TextMatrix(0, 0) = "Comprobante"
        .TextMatrix(0, 1) = "Fecha"
        .TextMatrix(0, 2) = "Monto Bruto"
        .TextMatrix(0, 3) = "Retenciones"
        .TextMatrix(0, 4) = "Líquido"
        .TextMatrix(0, 5) = "Total Agentes"
        .ColWidth(0) = 1
        .ColWidth(1) = 750
        .ColWidth(2) = 2000
        .ColWidth(3) = 2000
        .ColWidth(4) = 2000
        .ColWidth(5) = 2000
        .FixedCols = 0
        .FocusRect = flexFocusHeavy
        .HighLight = flexHighlightWithFocus
        .SelectionMode = flexSelectionByRow
        .AllowUserResizing = flexResizeColumns
        .ColAlignment(0) = 1
        .ColAlignment(1) = 7
        .ColAlignment(2) = 7
        .ColAlignment(3) = 7
        .ColAlignment(4) = 7
        .ColAlignment(5) = 7
    End With
    
End Sub

Public Sub CargardgAutocarga()
    
    Dim i As Integer
    Dim SQL As String
    i = 0
    Dim dblImporte As Double
    Dim strPeriodo As String
    
    Autocarga.dgListadoAutocarga.Rows = 2
    
    SQL = "Select Comprobante, Month(Fecha) As Mes, Year(Fecha) As Ano," _
    & " Sum(MontoBruto) As SMONTOBRUTO, Sum(Sellos) As SSELLOS, Sum(LibramientoPago) As SLIBRAMIENTOPAGO," _
    & " Sum(IIBB) As SIIBB, Sum(OtraRetencion) As SOTRARETENCION, Sum(ANTICIPO) As SANTICIPO," _
    & " Sum(Seguro) As SSEGURO, Sum(Descuento) As SDESCUENTO, Count(Proveedor) As TotalAgentes From LIQUIDACIONHONORARIOS" _
    & " Where Left(COMPROBANTE,6) = '" & "NoSIIF" & "'" _
    & " Group By COMPROBANTE, Year(FECHA), Month(FECHA)" _
    & " Order By Year(FECHA) DESC, Month(FECHA) DESC"
    
    If SQLNoMatch(SQL) = False Then
        Set rstListadoSlave = New ADODB.Recordset
        rstListadoSlave.Open SQL, dbSlave, adOpenDynamic, adLockOptimistic
        With Autocarga.dgListadoAutocarga
            If rstListadoSlave.BOF = False Then
                rstListadoSlave.MoveFirst
                While rstListadoSlave.EOF = False
                    i = i + 1
                    .RowHeight(i) = 300
                    .TextMatrix(i, 0) = rstListadoSlave!Comprobante
                    strPeriodo = rstListadoSlave!Mes & "/" & rstListadoSlave!Ano
                    .TextMatrix(i, 1) = strPeriodo
                    .TextMatrix(i, 2) = De_Num_a_Tx_01(rstListadoSlave!SMONTOBRUTO, , 2)
                    dblImporte = rstListadoSlave!SSellos + rstListadoSlave!SLibramientoPago + rstListadoSlave!SIIBB _
                    + rstListadoSlave!SOTRARETENCION + rstListadoSlave!SANTICIPO + rstListadoSlave!SSeguro + rstListadoSlave!SDescuento
                    .TextMatrix(i, 3) = De_Num_a_Tx_01(dblImporte, , 2)
                    dblImporte = rstListadoSlave!SMONTOBRUTO - dblImporte
                    .TextMatrix(i, 4) = De_Num_a_Tx_01(dblImporte, , 2)
                    .TextMatrix(i, 5) = De_Num_a_Tx_01(rstListadoSlave!TotalAgentes, True)
                    rstListadoSlave.MoveNext
                    .Rows = .Rows + 1
                Wend
            End If
            .Rows = .Rows - 1
        End With
        rstListadoSlave.Close
        Set rstListadoSlave = Nothing
    End If
    
'    With AutocargaCertificados
'        If Len(.dgAutocargaCertificado.TextMatrix(1, 0)) = "0" Then
'            .cmdAgregarCertificado.Enabled = False
'            .cmdEliminarCertificado.Enabled = False
'        Else
'            .cmdAgregarCertificado.Enabled = True
'            .cmdEliminarCertificado.Enabled = True
'        End If
'    End With
    
End Sub

Sub ConfigurardgLiquidacionMensualGanancias(CantidadDeMeses As Integer)

    Dim x As Integer
    x = 1
    
    With ResumenAnualGanancias.dgLiquidacionMensualGanancias
        .Clear
        .Cols = (CantidadDeMeses * 2 + 1)
        .Rows = 37
        .TextMatrix(0, 0) = "CONCEPTO"
        .ColWidth(0) = 2250
        .ColAlignment(0) = 1
        .FixedCols = 1
        .FocusRect = flexFocusHeavy
        .HighLight = flexHighlightWithFocus
        .AllowUserResizing = flexResizeColumns
        'Creamos las columnas restantes y las denominamos
        For x = 1 To (CantidadDeMeses * 2)
            If x Mod 2 = 0 Then
                .TextMatrix(0, x) = "Acumulado"
            Else
                .TextMatrix(0, x) = Format(Format("01/0" & ((x + 1) / 2) & "/2009", "mmmm"), ">")
            End If
            .ColWidth(x) = 1000
            .ColAlignment(x) = 7
        Next x
        x = 0
        'Asigamos el mismo alto a todas las filas
        For x = 0 To (.Rows - 1)
            .RowHeight(x) = 300
        Next x
        'Denominamos las filas
        .TextMatrix(1, 0) = "Remuneración Bruta"
        .TextMatrix(2, 0) = "Ajuste"
        .TextMatrix(3, 0) = "Remuneración Otros Empleos"
        .TextMatrix(5, 0) = "GANANCIA BRUTA"
        .TextMatrix(7, 0) = "Aporte Jubilacion Personal"
        .TextMatrix(8, 0) = "Aporte Obra Social Personal"
        .TextMatrix(9, 0) = "Adherente Obra Social"
        .TextMatrix(10, 0) = "Seguro de Vida Obligatorio"
        .TextMatrix(11, 0) = "Cuota Sindical"
        .TextMatrix(13, 0) = "GANANCIA NETA"
        .TextMatrix(15, 0) = "Seguro de Vida Optativo"
        .TextMatrix(16, 0) = "Servicio Doméstico"
        .TextMatrix(17, 0) = "Alquiler"
        .TextMatrix(18, 0) = "Cuota Médica Asistencial"
        .TextMatrix(19, 0) = "Donaciones"
        .TextMatrix(20, 0) = "Honorarios Médicos"
        .TextMatrix(21, 0) = "Ganancia No Imponible"
        .TextMatrix(22, 0) = "Conyugue"
        .TextMatrix(23, 0) = "Hijos"
        .TextMatrix(24, 0) = "Otras Cargas"
        .TextMatrix(25, 0) = "Deducción Especial"
        .TextMatrix(27, 0) = "GANANCIA IMPONIBLE"
        .TextMatrix(29, 0) = "Alícuota Aplicable"
        .TextMatrix(30, 0) = "Importe Variable"
        .TextMatrix(31, 0) = "Importe Fijo"
        .TextMatrix(32, 0) = "IMPUESTO"
        .TextMatrix(33, 0) = "RETENCIÓN ACUMULADA"
        .TextMatrix(34, 0) = "AJUSTES ACUMULADOS"
        .TextMatrix(36, 0) = "RETENER / DEVOLVER"
    End With
    
    x = 0
    
End Sub

Sub CargardgLiquidacionMensualGanancias(PuestoLaboral As String, AñoLiquidado As String)

    Dim i As Integer
    i = 0
    Dim dblImporteAcumulado As Double
    Dim dblImporteSLAVE As Double
    Dim strCLSlave As String
'    Dim x As Integer
'    x = 0
    Dim SQL As String
    
    
    
    With ResumenAnualGanancias.dgLiquidacionMensualGanancias
        Set rstListadoSlave = New ADODB.Recordset
        'Completamos el Haber Bruto liquidado por SLAVE
        For i = 1 To (.Cols - 1)
            If i Mod 2 <> 0 Then
                SQL = "Select Sum(HaberOptimo) As ImporteMensual From" _
                & " LIQUIDACIONGANANCIAS4TACATEGORIA Inner Join CODIGOLIQUIDACIONES On" _
                & " LIQUIDACIONGANANCIAS4TACATEGORIA.CodigoLiquidacion = CODIGOLIQUIDACIONES.Codigo" _
                & " Where Left(PERIODO, 2) = '" & Format((i + 1) / 2, "00") & "'" _
                & " And Right(PERIODO, 4) = '" & AñoLiquidado & "'" _
                & " And PUESTOLABORAL = '" & PuestoLaboral & "'"
                If SQLNoMatch(SQL) Then
                    dblImporteSLAVE = 0
                Else
                    rstListadoSlave.Open SQL, dbSlave, adOpenDynamic, adLockOptimistic
                    dblImporteSLAVE = rstListadoSlave!ImporteMensual
                    rstListadoSlave.Close
                End If
                .TextMatrix(1, i) = FormatNumber(dblImporteSLAVE, 2)
                dblImporteAcumulado = dblImporteAcumulado + dblImporteSLAVE
            Else
                .TextMatrix(1, i) = FormatNumber(dblImporteAcumulado, 2)
            End If
        Next i
        dblImporteAcumulado = 0
        'Completamos el Ajuste liquidado por SLAVE
        For i = 1 To (.Cols - 1)
            If i Mod 2 <> 0 Then
                SQL = "Select Sum(Ajuste) As ImporteMensual From" _
                & " LIQUIDACIONGANANCIAS4TACATEGORIA Inner Join CODIGOLIQUIDACIONES On" _
                & " LIQUIDACIONGANANCIAS4TACATEGORIA.CodigoLiquidacion = CODIGOLIQUIDACIONES.Codigo" _
                & " Where Left(PERIODO, 2) = '" & Format((i + 1) / 2, "00") & "'" _
                & " And Right(PERIODO, 4) = '" & AñoLiquidado & "'" _
                & " And PUESTOLABORAL = '" & PuestoLaboral & "'"
                If SQLNoMatch(SQL) Then
                    dblImporteSLAVE = 0
                Else
                    rstListadoSlave.Open SQL, dbSlave, adOpenDynamic, adLockOptimistic
                    dblImporteSLAVE = rstListadoSlave!ImporteMensual
                    rstListadoSlave.Close
                End If
                .TextMatrix(2, i) = FormatNumber(dblImporteSLAVE, 2)
                dblImporteAcumulado = dblImporteAcumulado + dblImporteSLAVE
            Else
                .TextMatrix(2, i) = FormatNumber(dblImporteAcumulado, 2)
            End If
        Next i
        dblImporteAcumulado = 0
        'Completamos el Pluriempleo liquidado por SLAVE
        For i = 1 To (.Cols - 1)
            If i Mod 2 <> 0 Then
                SQL = "Select Sum(Pluriempleo) As ImporteMensual From" _
                & " LIQUIDACIONGANANCIAS4TACATEGORIA Inner Join CODIGOLIQUIDACIONES On" _
                & " LIQUIDACIONGANANCIAS4TACATEGORIA.CodigoLiquidacion = CODIGOLIQUIDACIONES.Codigo" _
                & " Where Left(PERIODO, 2) = '" & Format((i + 1) / 2, "00") & "'" _
                & " And Right(PERIODO, 4) = '" & AñoLiquidado & "'" _
                & " And PUESTOLABORAL = '" & PuestoLaboral & "'"
                If SQLNoMatch(SQL) Then
                    dblImporteSLAVE = 0
                Else
                    rstListadoSlave.Open SQL, dbSlave, adOpenDynamic, adLockOptimistic
                    dblImporteSLAVE = rstListadoSlave!ImporteMensual
                    rstListadoSlave.Close
                End If
                .TextMatrix(3, i) = FormatNumber(dblImporteSLAVE, 2)
                dblImporteAcumulado = dblImporteAcumulado + dblImporteSLAVE
            Else
                .TextMatrix(3, i) = FormatNumber(dblImporteAcumulado, 2)
            End If
        Next i
        dblImporteAcumulado = 0
        'Completamos la Ganancia Bruta liquidada por SLAVE
        For i = 1 To (.Cols - 1)
            If i Mod 2 <> 0 Then
                SQL = "Select (Sum(HaberOptimo) + Sum(Ajuste) + Sum(Pluriempleo)) As ImporteMensual From" _
                & " LIQUIDACIONGANANCIAS4TACATEGORIA Inner Join CODIGOLIQUIDACIONES On" _
                & " LIQUIDACIONGANANCIAS4TACATEGORIA.CodigoLiquidacion = CODIGOLIQUIDACIONES.Codigo" _
                & " Where Left(PERIODO, 2) = '" & Format((i + 1) / 2, "00") & "'" _
                & " And Right(PERIODO, 4) = '" & AñoLiquidado & "'" _
                & " And PUESTOLABORAL = '" & PuestoLaboral & "'"
                If SQLNoMatch(SQL) Then
                    dblImporteSLAVE = 0
                Else
                    rstListadoSlave.Open SQL, dbSlave, adOpenDynamic, adLockOptimistic
                    dblImporteSLAVE = rstListadoSlave!ImporteMensual
                    rstListadoSlave.Close
                End If
                .TextMatrix(5, i) = FormatNumber(dblImporteSLAVE, 2)
                dblImporteAcumulado = dblImporteAcumulado + dblImporteSLAVE
            Else
                .TextMatrix(5, i) = FormatNumber(dblImporteAcumulado, 2)
            End If
        Next i
        dblImporteAcumulado = 0
        'Completamos el Aporte Jubilatorio liquidado por SLAVE
        For i = 1 To (.Cols - 1)
            If i Mod 2 <> 0 Then
                SQL = "Select Sum(Jubilacion) As ImporteMensual From" _
                & " LIQUIDACIONGANANCIAS4TACATEGORIA Inner Join CODIGOLIQUIDACIONES On" _
                & " LIQUIDACIONGANANCIAS4TACATEGORIA.CodigoLiquidacion = CODIGOLIQUIDACIONES.Codigo" _
                & " Where Left(PERIODO, 2) = '" & Format((i + 1) / 2, "00") & "'" _
                & " And Right(PERIODO, 4) = '" & AñoLiquidado & "'" _
                & " And PUESTOLABORAL = '" & PuestoLaboral & "'"
                If SQLNoMatch(SQL) Then
                    dblImporteSLAVE = 0
                Else
                    rstListadoSlave.Open SQL, dbSlave, adOpenDynamic, adLockOptimistic
                    dblImporteSLAVE = rstListadoSlave!ImporteMensual
                    rstListadoSlave.Close
                End If
                .TextMatrix(7, i) = FormatNumber(dblImporteSLAVE, 2)
                dblImporteAcumulado = dblImporteAcumulado + dblImporteSLAVE
            Else
                .TextMatrix(7, i) = FormatNumber(dblImporteAcumulado, 2)
            End If
        Next i
        dblImporteAcumulado = 0
        'Completamos el Aporte Obra Social liquidado por SLAVE
        For i = 1 To (.Cols - 1)
            If i Mod 2 <> 0 Then
                SQL = "Select Sum(ObraSocial) As ImporteMensual From" _
                & " LIQUIDACIONGANANCIAS4TACATEGORIA Inner Join CODIGOLIQUIDACIONES On" _
                & " LIQUIDACIONGANANCIAS4TACATEGORIA.CodigoLiquidacion = CODIGOLIQUIDACIONES.Codigo" _
                & " Where Left(PERIODO, 2) = '" & Format((i + 1) / 2, "00") & "'" _
                & " And Right(PERIODO, 4) = '" & AñoLiquidado & "'" _
                & " And PUESTOLABORAL = '" & PuestoLaboral & "'"
                If SQLNoMatch(SQL) Then
                    dblImporteSLAVE = 0
                Else
                    rstListadoSlave.Open SQL, dbSlave, adOpenDynamic, adLockOptimistic
                    dblImporteSLAVE = rstListadoSlave!ImporteMensual
                    rstListadoSlave.Close
                End If
                .TextMatrix(8, i) = FormatNumber(dblImporteSLAVE, 2)
                dblImporteAcumulado = dblImporteAcumulado + dblImporteSLAVE
            Else
                .TextMatrix(8, i) = FormatNumber(dblImporteAcumulado, 2)
            End If
        Next i
        dblImporteAcumulado = 0
        'Completamos el Adherente Obra Social liquidado por SLAVE
        For i = 1 To (.Cols - 1)
            If i Mod 2 <> 0 Then
                SQL = "Select Sum(AdherenteObraSocial) As ImporteMensual From" _
                & " LIQUIDACIONGANANCIAS4TACATEGORIA Inner Join CODIGOLIQUIDACIONES On" _
                & " LIQUIDACIONGANANCIAS4TACATEGORIA.CodigoLiquidacion = CODIGOLIQUIDACIONES.Codigo" _
                & " Where Left(PERIODO, 2) = '" & Format((i + 1) / 2, "00") & "'" _
                & " And Right(PERIODO, 4) = '" & AñoLiquidado & "'" _
                & " And PUESTOLABORAL = '" & PuestoLaboral & "'"
                If SQLNoMatch(SQL) Then
                    dblImporteSLAVE = 0
                Else
                    rstListadoSlave.Open SQL, dbSlave, adOpenDynamic, adLockOptimistic
                    dblImporteSLAVE = rstListadoSlave!ImporteMensual
                    rstListadoSlave.Close
                End If
                .TextMatrix(9, i) = FormatNumber(dblImporteSLAVE, 2)
                dblImporteAcumulado = dblImporteAcumulado + dblImporteSLAVE
            Else
                .TextMatrix(9, i) = FormatNumber(dblImporteAcumulado, 2)
            End If
        Next i
        dblImporteAcumulado = 0
        'Completamos el Seguro de Vida Obligatorio liquidado por SLAVE
        For i = 1 To (.Cols - 1)
            If i Mod 2 <> 0 Then
                SQL = "Select Sum(SeguroDeVidaObligatorio) As ImporteMensual From" _
                & " LIQUIDACIONGANANCIAS4TACATEGORIA Inner Join CODIGOLIQUIDACIONES On" _
                & " LIQUIDACIONGANANCIAS4TACATEGORIA.CodigoLiquidacion = CODIGOLIQUIDACIONES.Codigo" _
                & " Where Left(PERIODO, 2) = '" & Format((i + 1) / 2, "00") & "'" _
                & " And Right(PERIODO, 4) = '" & AñoLiquidado & "'" _
                & " And PUESTOLABORAL = '" & PuestoLaboral & "'"
                If SQLNoMatch(SQL) Then
                    dblImporteSLAVE = 0
                Else
                    rstListadoSlave.Open SQL, dbSlave, adOpenDynamic, adLockOptimistic
                    dblImporteSLAVE = rstListadoSlave!ImporteMensual
                    rstListadoSlave.Close
                End If
                .TextMatrix(10, i) = FormatNumber(dblImporteSLAVE, 2)
                dblImporteAcumulado = dblImporteAcumulado + dblImporteSLAVE
            Else
                .TextMatrix(10, i) = FormatNumber(dblImporteAcumulado, 2)
            End If
        Next i
        dblImporteAcumulado = 0
        'Completamos el Cuota Sindical liquidada por SLAVE
        For i = 1 To (.Cols - 1)
            If i Mod 2 <> 0 Then
                SQL = "Select Sum(CuotaSindical) As ImporteMensual From" _
                & " LIQUIDACIONGANANCIAS4TACATEGORIA Inner Join CODIGOLIQUIDACIONES On" _
                & " LIQUIDACIONGANANCIAS4TACATEGORIA.CodigoLiquidacion = CODIGOLIQUIDACIONES.Codigo" _
                & " Where Left(PERIODO, 2) = '" & Format((i + 1) / 2, "00") & "'" _
                & " And Right(PERIODO, 4) = '" & AñoLiquidado & "'" _
                & " And PUESTOLABORAL = '" & PuestoLaboral & "'"
                If SQLNoMatch(SQL) Then
                    dblImporteSLAVE = 0
                Else
                    rstListadoSlave.Open SQL, dbSlave, adOpenDynamic, adLockOptimistic
                    dblImporteSLAVE = rstListadoSlave!ImporteMensual
                    rstListadoSlave.Close
                End If
                .TextMatrix(11, i) = FormatNumber(dblImporteSLAVE, 2)
                dblImporteAcumulado = dblImporteAcumulado + dblImporteSLAVE
            Else
                .TextMatrix(11, i) = FormatNumber(dblImporteAcumulado, 2)
            End If
        Next i
        dblImporteAcumulado = 0
        'Completamos la Ganancia Neta liquidada por SLAVE
        For i = 1 To (.Cols - 1)
            If i Mod 2 <> 0 Then
                SQL = "Select (Sum(HaberOptimo) + Sum(Ajuste) + Sum(Pluriempleo)" _
                & " - Sum(Jubilacion) - Sum(ObraSocial) - Sum(AdherenteObraSocial)" _
                & " - Sum(SeguroDeVidaObligatorio) - Sum(CuotaSindical)) As ImporteMensual From" _
                & " LIQUIDACIONGANANCIAS4TACATEGORIA Inner Join CODIGOLIQUIDACIONES On" _
                & " LIQUIDACIONGANANCIAS4TACATEGORIA.CodigoLiquidacion = CODIGOLIQUIDACIONES.Codigo" _
                & " Where Left(PERIODO, 2) = '" & Format((i + 1) / 2, "00") & "'" _
                & " And Right(PERIODO, 4) = '" & AñoLiquidado & "'" _
                & " And PUESTOLABORAL = '" & PuestoLaboral & "'"
                If SQLNoMatch(SQL) Then
                    dblImporteSLAVE = 0
                Else
                    rstListadoSlave.Open SQL, dbSlave, adOpenDynamic, adLockOptimistic
                    dblImporteSLAVE = rstListadoSlave!ImporteMensual
                    rstListadoSlave.Close
                End If
                .TextMatrix(13, i) = FormatNumber(dblImporteSLAVE, 2)
                dblImporteAcumulado = dblImporteAcumulado + dblImporteSLAVE
            Else
                .TextMatrix(13, i) = FormatNumber(dblImporteAcumulado, 2)
            End If
        Next i
        dblImporteAcumulado = 0
        'Completamos el Seguro de Vida Obtativo liquidado por SLAVE
        For i = 1 To (.Cols - 1)
            If i Mod 2 <> 0 Then
                SQL = "Select Sum(SeguroDeVidaOptativo) As ImporteMensual From" _
                & " LIQUIDACIONGANANCIAS4TACATEGORIA Inner Join CODIGOLIQUIDACIONES On" _
                & " LIQUIDACIONGANANCIAS4TACATEGORIA.CodigoLiquidacion = CODIGOLIQUIDACIONES.Codigo" _
                & " Where Left(PERIODO, 2) = '" & Format((i + 1) / 2, "00") & "'" _
                & " And Right(PERIODO, 4) = '" & AñoLiquidado & "'" _
                & " And PUESTOLABORAL = '" & PuestoLaboral & "'"
                If SQLNoMatch(SQL) Then
                    dblImporteSLAVE = 0
                Else
                    rstListadoSlave.Open SQL, dbSlave, adOpenDynamic, adLockOptimistic
                    dblImporteSLAVE = rstListadoSlave!ImporteMensual
                    rstListadoSlave.Close
                End If
                .TextMatrix(15, i) = FormatNumber(dblImporteSLAVE, 2)
                dblImporteAcumulado = dblImporteAcumulado + dblImporteSLAVE
            Else
                .TextMatrix(15, i) = FormatNumber(dblImporteAcumulado, 2)
            End If
        Next i
        dblImporteAcumulado = 0
        'Completamos el Servicio Domestico liquidado por SLAVE
        For i = 1 To (.Cols - 1)
            If i Mod 2 <> 0 Then
                SQL = "Select Sum(ServicioDomestico) As ImporteMensual From" _
                & " LIQUIDACIONGANANCIAS4TACATEGORIA Inner Join CODIGOLIQUIDACIONES On" _
                & " LIQUIDACIONGANANCIAS4TACATEGORIA.CodigoLiquidacion = CODIGOLIQUIDACIONES.Codigo" _
                & " Where Left(PERIODO, 2) = '" & Format((i + 1) / 2, "00") & "'" _
                & " And Right(PERIODO, 4) = '" & AñoLiquidado & "'" _
                & " And PUESTOLABORAL = '" & PuestoLaboral & "'"
                If SQLNoMatch(SQL) Then
                    dblImporteSLAVE = 0
                Else
                    rstListadoSlave.Open SQL, dbSlave, adOpenDynamic, adLockOptimistic
                    dblImporteSLAVE = rstListadoSlave!ImporteMensual
                    rstListadoSlave.Close
                End If
                .TextMatrix(16, i) = FormatNumber(dblImporteSLAVE, 2)
                dblImporteAcumulado = dblImporteAcumulado + dblImporteSLAVE
            Else
                .TextMatrix(16, i) = FormatNumber(dblImporteAcumulado, 2)
            End If
        Next i
        dblImporteAcumulado = 0
        'Completamos el Deducción por Alquiler por SLAVE
        For i = 1 To (.Cols - 1)
            If i Mod 2 <> 0 Then
                SQL = "Select Sum(Alquileres) As ImporteMensual From" _
                & " LIQUIDACIONGANANCIAS4TACATEGORIA Inner Join CODIGOLIQUIDACIONES On" _
                & " LIQUIDACIONGANANCIAS4TACATEGORIA.CodigoLiquidacion = CODIGOLIQUIDACIONES.Codigo" _
                & " Where Left(PERIODO, 2) = '" & Format((i + 1) / 2, "00") & "'" _
                & " And Right(PERIODO, 4) = '" & AñoLiquidado & "'" _
                & " And PUESTOLABORAL = '" & PuestoLaboral & "'"
                If SQLNoMatch(SQL) Then
                    dblImporteSLAVE = 0
                Else
                    rstListadoSlave.Open SQL, dbSlave, adOpenDynamic, adLockOptimistic
                    dblImporteSLAVE = rstListadoSlave!ImporteMensual
                    rstListadoSlave.Close
                End If
                .TextMatrix(17, i) = FormatNumber(dblImporteSLAVE, 2)
                dblImporteAcumulado = dblImporteAcumulado + dblImporteSLAVE
            Else
                .TextMatrix(17, i) = FormatNumber(dblImporteAcumulado, 2)
            End If
        Next i
        dblImporteAcumulado = 0
        'Completamos el Cuota Médico Asistencial liquidado por SLAVE
        For i = 1 To (.Cols - 1)
            If i Mod 2 <> 0 Then
                SQL = "Select Sum(CuotaMedicoAsistencial) As ImporteMensual From" _
                & " LIQUIDACIONGANANCIAS4TACATEGORIA Inner Join CODIGOLIQUIDACIONES On" _
                & " LIQUIDACIONGANANCIAS4TACATEGORIA.CodigoLiquidacion = CODIGOLIQUIDACIONES.Codigo" _
                & " Where Left(PERIODO, 2) = '" & Format((i + 1) / 2, "00") & "'" _
                & " And Right(PERIODO, 4) = '" & AñoLiquidado & "'" _
                & " And PUESTOLABORAL = '" & PuestoLaboral & "'"
                If SQLNoMatch(SQL) Then
                    dblImporteSLAVE = 0
                Else
                    rstListadoSlave.Open SQL, dbSlave, adOpenDynamic, adLockOptimistic
                    dblImporteSLAVE = rstListadoSlave!ImporteMensual
                    rstListadoSlave.Close
                End If
                .TextMatrix(18, i) = FormatNumber(dblImporteSLAVE, 2)
                dblImporteAcumulado = dblImporteAcumulado + dblImporteSLAVE
            Else
                .TextMatrix(18, i) = FormatNumber(dblImporteAcumulado, 2)
            End If
        Next i
        dblImporteAcumulado = 0
        'Completamos el Donaciones liquidado por SLAVE
        For i = 1 To (.Cols - 1)
            If i Mod 2 <> 0 Then
                SQL = "Select Sum(Donaciones) As ImporteMensual From" _
                & " LIQUIDACIONGANANCIAS4TACATEGORIA Inner Join CODIGOLIQUIDACIONES On" _
                & " LIQUIDACIONGANANCIAS4TACATEGORIA.CodigoLiquidacion = CODIGOLIQUIDACIONES.Codigo" _
                & " Where Left(PERIODO, 2) = '" & Format((i + 1) / 2, "00") & "'" _
                & " And Right(PERIODO, 4) = '" & AñoLiquidado & "'" _
                & " And PUESTOLABORAL = '" & PuestoLaboral & "'"
                If SQLNoMatch(SQL) Then
                    dblImporteSLAVE = 0
                Else
                    rstListadoSlave.Open SQL, dbSlave, adOpenDynamic, adLockOptimistic
                    dblImporteSLAVE = rstListadoSlave!ImporteMensual
                    rstListadoSlave.Close
                End If
                .TextMatrix(19, i) = FormatNumber(dblImporteSLAVE, 2)
                dblImporteAcumulado = dblImporteAcumulado + dblImporteSLAVE
            Else
                .TextMatrix(19, i) = FormatNumber(dblImporteAcumulado, 2)
            End If
        Next i
        dblImporteAcumulado = 0
        'Completamos el Honorarios Médicos liquidado por SLAVE
        For i = 1 To (.Cols - 1)
            If i Mod 2 <> 0 Then
                SQL = "Select Sum(HonorariosMedicos) As ImporteMensual From" _
                & " LIQUIDACIONGANANCIAS4TACATEGORIA Inner Join CODIGOLIQUIDACIONES On" _
                & " LIQUIDACIONGANANCIAS4TACATEGORIA.CodigoLiquidacion = CODIGOLIQUIDACIONES.Codigo" _
                & " Where Left(PERIODO, 2) = '" & Format((i + 1) / 2, "00") & "'" _
                & " And Right(PERIODO, 4) = '" & AñoLiquidado & "'" _
                & " And PUESTOLABORAL = '" & PuestoLaboral & "'"
                If SQLNoMatch(SQL) Then
                    dblImporteSLAVE = 0
                Else
                    rstListadoSlave.Open SQL, dbSlave, adOpenDynamic, adLockOptimistic
                    dblImporteSLAVE = rstListadoSlave!ImporteMensual
                    rstListadoSlave.Close
                End If
                .TextMatrix(20, i) = FormatNumber(dblImporteSLAVE, 2)
                dblImporteAcumulado = dblImporteAcumulado + dblImporteSLAVE
            Else
                .TextMatrix(20, i) = FormatNumber(dblImporteAcumulado, 2)
            End If
        Next i
        dblImporteAcumulado = 0
        'Completamos el Ganancia Mínima No Imponible liquidado por SLAVE
        For i = 1 To (.Cols - 1)
            If i Mod 2 <> 0 Then
                SQL = "Select Sum(MinimoNoImponible) As ImporteMensual From" _
                & " LIQUIDACIONGANANCIAS4TACATEGORIA Inner Join CODIGOLIQUIDACIONES On" _
                & " LIQUIDACIONGANANCIAS4TACATEGORIA.CodigoLiquidacion = CODIGOLIQUIDACIONES.Codigo" _
                & " Where Left(PERIODO, 2) = '" & Format((i + 1) / 2, "00") & "'" _
                & " And Right(PERIODO, 4) = '" & AñoLiquidado & "'" _
                & " And PUESTOLABORAL = '" & PuestoLaboral & "'"
                If SQLNoMatch(SQL) Then
                    dblImporteSLAVE = 0
                Else
                    rstListadoSlave.Open SQL, dbSlave, adOpenDynamic, adLockOptimistic
                    dblImporteSLAVE = rstListadoSlave!ImporteMensual
                    rstListadoSlave.Close
                End If
                .TextMatrix(21, i) = FormatNumber(dblImporteSLAVE, 2)
                dblImporteAcumulado = dblImporteAcumulado + dblImporteSLAVE
            Else
                .TextMatrix(21, i) = FormatNumber(dblImporteAcumulado, 2)
            End If
        Next i
        dblImporteAcumulado = 0
        'Completamos la Deducción por Conyuge liquidado por SLAVE
        For i = 1 To (.Cols - 1)
            If i Mod 2 <> 0 Then
                SQL = "Select Sum(Conyuge) As ImporteMensual From" _
                & " LIQUIDACIONGANANCIAS4TACATEGORIA Inner Join CODIGOLIQUIDACIONES On" _
                & " LIQUIDACIONGANANCIAS4TACATEGORIA.CodigoLiquidacion = CODIGOLIQUIDACIONES.Codigo" _
                & " Where Left(PERIODO, 2) = '" & Format((i + 1) / 2, "00") & "'" _
                & " And Right(PERIODO, 4) = '" & AñoLiquidado & "'" _
                & " And PUESTOLABORAL = '" & PuestoLaboral & "'"
                If SQLNoMatch(SQL) Then
                    dblImporteSLAVE = 0
                Else
                    rstListadoSlave.Open SQL, dbSlave, adOpenDynamic, adLockOptimistic
                    dblImporteSLAVE = rstListadoSlave!ImporteMensual
                    rstListadoSlave.Close
                End If
                .TextMatrix(22, i) = FormatNumber(dblImporteSLAVE, 2)
                dblImporteAcumulado = dblImporteAcumulado + dblImporteSLAVE
            Else
                .TextMatrix(22, i) = FormatNumber(dblImporteAcumulado, 2)
            End If
        Next i
        dblImporteAcumulado = 0
        'Completamos la Deducción por Hijos liquidado por SLAVE
        For i = 1 To (.Cols - 1)
            If i Mod 2 <> 0 Then
                SQL = "Select Sum(Hijo) As ImporteMensual From" _
                & " LIQUIDACIONGANANCIAS4TACATEGORIA Inner Join CODIGOLIQUIDACIONES On" _
                & " LIQUIDACIONGANANCIAS4TACATEGORIA.CodigoLiquidacion = CODIGOLIQUIDACIONES.Codigo" _
                & " Where Left(PERIODO, 2) = '" & Format((i + 1) / 2, "00") & "'" _
                & " And Right(PERIODO, 4) = '" & AñoLiquidado & "'" _
                & " And PUESTOLABORAL = '" & PuestoLaboral & "'"
                If SQLNoMatch(SQL) Then
                    dblImporteSLAVE = 0
                Else
                    rstListadoSlave.Open SQL, dbSlave, adOpenDynamic, adLockOptimistic
                    dblImporteSLAVE = rstListadoSlave!ImporteMensual
                    rstListadoSlave.Close
                End If
                .TextMatrix(23, i) = FormatNumber(dblImporteSLAVE, 2)
                dblImporteAcumulado = dblImporteAcumulado + dblImporteSLAVE
            Else
                .TextMatrix(23, i) = FormatNumber(dblImporteAcumulado, 2)
            End If
        Next i
        dblImporteAcumulado = 0
        'Completamos la Deducción por Otras Cargas de Familia liquidado por SLAVE
        For i = 1 To (.Cols - 1)
            If i Mod 2 <> 0 Then
                SQL = "Select Sum(OtrasCargasDeFamilia) As ImporteMensual From" _
                & " LIQUIDACIONGANANCIAS4TACATEGORIA Inner Join CODIGOLIQUIDACIONES On" _
                & " LIQUIDACIONGANANCIAS4TACATEGORIA.CodigoLiquidacion = CODIGOLIQUIDACIONES.Codigo" _
                & " Where Left(PERIODO, 2) = '" & Format((i + 1) / 2, "00") & "'" _
                & " And Right(PERIODO, 4) = '" & AñoLiquidado & "'" _
                & " And PUESTOLABORAL = '" & PuestoLaboral & "'"
                If SQLNoMatch(SQL) Then
                    dblImporteSLAVE = 0
                Else
                    rstListadoSlave.Open SQL, dbSlave, adOpenDynamic, adLockOptimistic
                    dblImporteSLAVE = rstListadoSlave!ImporteMensual
                    rstListadoSlave.Close
                End If
                .TextMatrix(24, i) = FormatNumber(dblImporteSLAVE, 2)
                dblImporteAcumulado = dblImporteAcumulado + dblImporteSLAVE
            Else
                .TextMatrix(24, i) = FormatNumber(dblImporteAcumulado, 2)
            End If
        Next i
        dblImporteAcumulado = 0
        'Completamos la Deducción Especial liquidado por SLAVE
        For i = 1 To (.Cols - 1)
            If i Mod 2 <> 0 Then
                SQL = "Select Sum(DeduccionEspecial) As ImporteMensual From" _
                & " LIQUIDACIONGANANCIAS4TACATEGORIA Inner Join CODIGOLIQUIDACIONES On" _
                & " LIQUIDACIONGANANCIAS4TACATEGORIA.CodigoLiquidacion = CODIGOLIQUIDACIONES.Codigo" _
                & " Where Left(PERIODO, 2) = '" & Format((i + 1) / 2, "00") & "'" _
                & " And Right(PERIODO, 4) = '" & AñoLiquidado & "'" _
                & " And PUESTOLABORAL = '" & PuestoLaboral & "'"
                If SQLNoMatch(SQL) Then
                    dblImporteSLAVE = 0
                Else
                    rstListadoSlave.Open SQL, dbSlave, adOpenDynamic, adLockOptimistic
                    dblImporteSLAVE = rstListadoSlave!ImporteMensual
                    rstListadoSlave.Close
                End If
                .TextMatrix(25, i) = FormatNumber(dblImporteSLAVE, 2)
                dblImporteAcumulado = dblImporteAcumulado + dblImporteSLAVE
            Else
                .TextMatrix(25, i) = FormatNumber(dblImporteAcumulado, 2)
            End If
        Next i
        dblImporteAcumulado = 0
        'Completamos la Ganancia Imponible liquidada por SLAVE
        For i = 1 To (.Cols - 1)
            If i Mod 2 <> 0 Then
                SQL = "Select (Sum(HaberOptimo) + Sum(Ajuste) + Sum(Pluriempleo)" _
                & " - Sum(Jubilacion) - Sum(ObraSocial) - Sum(AdherenteObraSocial)" _
                & " - Sum(SeguroDeVidaObligatorio) - Sum(CuotaSindical)" _
                & " - Sum(SeguroDeVidaOptativo) - Sum(ServicioDomestico) - Sum(Alquileres)" _
                & " - Sum(CuotaMedicoAsistencial) - Sum(Donaciones) - Sum(HonorariosMedicos) - Sum(MinimoNoImponible)" _
                & " - Sum(Conyuge) - Sum(Hijo) - Sum(OtrasCargasDeFamilia) - Sum(DeduccionEspecial)) As ImporteMensual From" _
                & " LIQUIDACIONGANANCIAS4TACATEGORIA Inner Join CODIGOLIQUIDACIONES On" _
                & " LIQUIDACIONGANANCIAS4TACATEGORIA.CodigoLiquidacion = CODIGOLIQUIDACIONES.Codigo" _
                & " Where Left(PERIODO, 2) = '" & Format((i + 1) / 2, "00") & "'" _
                & " And Right(PERIODO, 4) = '" & AñoLiquidado & "'" _
                & " And PUESTOLABORAL = '" & PuestoLaboral & "'"
                If SQLNoMatch(SQL) Then
                    dblImporteSLAVE = 0
                Else
                    rstListadoSlave.Open SQL, dbSlave, adOpenDynamic, adLockOptimistic
                    dblImporteSLAVE = rstListadoSlave!ImporteMensual
                    rstListadoSlave.Close
                End If
                .TextMatrix(28, i) = FormatNumber(dblImporteSLAVE, 2)
                dblImporteAcumulado = dblImporteAcumulado + dblImporteSLAVE
            Else
                .TextMatrix(28, i) = FormatNumber(dblImporteAcumulado, 2)
            End If
        Next i
        dblImporteAcumulado = 0
        'Completamos la Alícuota Aplicable, Importe Variable, Importe Fijo e Impuesto Determinado SLAVE
        For i = 1 To (.Cols - 1)
            If i Mod 2 <> 0 Then
                strCLSlave = Format((i + 1) / 2, "00") & "/" & AñoLiquidado
                strCLSlave = BuscarCodigoLiquidacion(strCLSlave)
                SQL = "Select (Sum(HaberOptimo) + Sum(Ajuste) + Sum(Pluriempleo)" _
                & " - Sum(Jubilacion) - Sum(ObraSocial) - Sum(AdherenteObraSocial)" _
                & " - Sum(SeguroDeVidaObligatorio) - Sum(CuotaSindical)" _
                & " - Sum(SeguroDeVidaOptativo) - Sum(ServicioDomestico) - Sum(Alquileres)" _
                & " - Sum(CuotaMedicoAsistencial) - Sum(Donaciones) - Sum(HonorariosMedicos) - Sum(MinimoNoImponible)" _
                & " - Sum(Conyuge) - Sum(Hijo) - Sum(OtrasCargasDeFamilia) - Sum(DeduccionEspecial)) As ImporteMensual From" _
                & " LIQUIDACIONGANANCIAS4TACATEGORIA Inner Join CODIGOLIQUIDACIONES On" _
                & " LIQUIDACIONGANANCIAS4TACATEGORIA.CodigoLiquidacion = CODIGOLIQUIDACIONES.Codigo" _
                & " Where Left(PERIODO, 2) = '" & Format((i + 1) / 2, "00") & "'" _
                & " And Right(PERIODO, 4) = '" & AñoLiquidado & "'" _
                & " And PUESTOLABORAL = '" & PuestoLaboral & "'"
                If SQLNoMatch(SQL) Then
                    dblImporteSLAVE = 0
                Else
                    rstListadoSlave.Open SQL, dbSlave, adOpenDynamic, adLockOptimistic
                    dblImporteSLAVE = rstListadoSlave!ImporteMensual
                    rstListadoSlave.Close
                End If
                dblImporteAcumulado = dblImporteAcumulado + dblImporteSLAVE
            Else
                dblImporteSLAVE = CalcularAlicuotaAplicable(PuestoLaboral, strCLSlave, dblImporteAcumulado)
                .TextMatrix(29, i) = dblImporteSLAVE * 100 & " %"
                dblImporteSLAVE = CalcularImporteVariable(PuestoLaboral, strCLSlave, dblImporteSLAVE, dblImporteAcumulado)
                .TextMatrix(30, i) = FormatNumber(dblImporteSLAVE, 2)
                dblImporteSLAVE = dblImporteSLAVE + CalcularImporteFijo(PuestoLaboral, strCLSlave, dblImporteAcumulado)
                .TextMatrix(31, i) = FormatNumber(CalcularImporteFijo(PuestoLaboral, strCLSlave, dblImporteAcumulado), 2)
                .TextMatrix(32, i) = FormatNumber(dblImporteSLAVE, 2)
            End If
        Next i
        dblImporteAcumulado = 0

        'Completamos el Imporet a Retener/Devolver por SLAVE
        For i = 1 To (.Cols - 1)
            If i Mod 2 <> 0 Then
                SQL = "Select Sum(Retencion) As ImporteMensual From" _
                & " LIQUIDACIONGANANCIAS4TACATEGORIA Inner Join CODIGOLIQUIDACIONES On" _
                & " LIQUIDACIONGANANCIAS4TACATEGORIA.CodigoLiquidacion = CODIGOLIQUIDACIONES.Codigo" _
                & " Where Left(PERIODO, 2) = '" & Format((i + 1) / 2, "00") & "'" _
                & " And Right(PERIODO, 4) = '" & AñoLiquidado & "'" _
                & " And PUESTOLABORAL = '" & PuestoLaboral & "'"
                If SQLNoMatch(SQL) Then
                    dblImporteSLAVE = 0
                Else
                    rstListadoSlave.Open SQL, dbSlave, adOpenDynamic, adLockOptimistic
                    dblImporteSLAVE = rstListadoSlave!ImporteMensual
                    rstListadoSlave.Close
                End If
                dblImporteAcumulado = dblImporteAcumulado + dblImporteSLAVE
            Else
                .TextMatrix(33, i) = FormatNumber(dblImporteAcumulado - dblImporteSLAVE, 2)
'                .TextMatrix(33, i) = FormatNumber(rstListadoSlave!AjusteRetencion, 2)
                .TextMatrix(36, i) = FormatNumber(dblImporteSLAVE, 2)
            End If
        Next i
        dblImporteAcumulado = 0
        
        
    
    End With
    
    Set rstListadoSlave = Nothing
    i = 0
    dblImporteSLAVE = 0
    dblImporteAcumulado = 0
    SQL = ""
         
End Sub

Public Sub ConfigurardgHaberesLiquidados()

    With ReciboDeSueldo.dgHaberesLiquidados
        .Clear
        .Cols = 3
        .Rows = 2
        .TextMatrix(0, 0) = "Código"
        .TextMatrix(0, 1) = "Haberes/Conceptos"
        .TextMatrix(0, 2) = "Importe"
        .ColWidth(0) = 1000
        .ColWidth(1) = 3000
        .ColWidth(2) = 1500
        .FixedCols = 0
        .FixedRows = 1
        .FocusRect = flexFocusHeavy
        .HighLight = flexHighlightWithFocus
        .SelectionMode = flexSelectionByRow
        .ColAlignment(0) = 1
        .ColAlignment(1) = 1
        .ColAlignment(2) = 7
    End With
    
End Sub

Public Sub CargardgHaberesLiquidados(PuestoLaboral As String, CodigoLiquidacion As String)

    Dim i As Integer
    Dim SQL As String
    Dim SumaTotal As Double
    
    SumaTotal = 0
    i = 0
    ReciboDeSueldo.dgHaberesLiquidados.Rows = 2
    
    SQL = "Select *" _
    & " From LIQUIDACIONSUELDOS Inner Join CONCEPTOS" _
    & " On LIQUIDACIONSUELDOS.CodigoConcepto = CONCEPTOS.Codigo" _
    & " Where (PUESTOLABORAL = " & "'" & PuestoLaboral & "'" _
    & " And CodigoLiquidacion = " & "'" & CodigoLiquidacion & "'" _
    & " And CodigoConcepto < " & "'" & "0200" & "'" _
    & " And CodigoConcepto <> " & "'" & "0115" & "')" _
    & " Or (PUESTOLABORAL = " & "'" & PuestoLaboral & "'" _
    & " And CodigoLiquidacion = " & "'" & CodigoLiquidacion & "'" _
    & " And CodigoConcepto = " & "'" & "0603" & "')" _
    & " Order By CodigoConcepto Asc"
    If SQLNoMatch(SQL) = False Then
        Set rstListadoSlave = New ADODB.Recordset
        rstListadoSlave.Open SQL, dbSlave, adOpenDynamic, adLockOptimistic
        With ReciboDeSueldo.dgHaberesLiquidados
            If rstListadoSlave.EOF = False Then
                rstListadoSlave.MoveFirst
                While rstListadoSlave.EOF = False
                    i = i + 1
                    .RowHeight(i) = 300
                    .TextMatrix(i, 0) = rstListadoSlave!CodigoConcepto
                    .TextMatrix(i, 1) = rstListadoSlave!Denominacion
                    .TextMatrix(i, 2) = FormatNumber(rstListadoSlave!Importe, 2)
                    SumaTotal = SumaTotal + rstListadoSlave!Importe
                    rstListadoSlave.MoveNext
                    .Rows = .Rows + 1
                Wend
                ReciboDeSueldo.txtTotalHaberes.Text = FormatNumber(SumaTotal, 2)
                ReciboDeSueldo.txtSueldoNeto.Text = FormatNumber(SumaTotal, 2)
            End If
            .Rows = .Rows - 1
        End With
        rstListadoSlave.Close
        Set rstListadoSlave = Nothing
    Else
        ConfigurardgHaberesLiquidados
        ReciboDeSueldo.txtTotalHaberes.Text = ""
        ReciboDeSueldo.txtSueldoNeto.Text = ""
    End If
    
End Sub

Public Sub ConfigurardgDescuentosLiquidados()

    With ReciboDeSueldo.dgDescuentosLiquidados
        .Clear
        .Cols = 3
        .Rows = 2
        .TextMatrix(0, 0) = "Código"
        .TextMatrix(0, 1) = "Descuentos/Conceptos"
        .TextMatrix(0, 2) = "Importe"
        .ColWidth(0) = 1000
        .ColWidth(1) = 3000
        .ColWidth(2) = 1500
        .FixedCols = 0
        .FixedRows = 1
        .FocusRect = flexFocusHeavy
        .HighLight = flexHighlightWithFocus
        .SelectionMode = flexSelectionByRow
        .ColAlignment(0) = 1
        .ColAlignment(1) = 1
        .ColAlignment(2) = 7
    End With
    
End Sub

Public Sub CargardgDescuentosLiquidados(PuestoLaboral As String, CodigoLiquidacion As String)

    Dim i As Integer
    Dim SQL As String
    Dim SumaTotal As Double
    
    SumaTotal = 0
    i = 0
    ReciboDeSueldo.dgDescuentosLiquidados.Rows = 2
    
    SQL = "Select *" _
    & " From LIQUIDACIONSUELDOS Inner Join CONCEPTOS" _
    & " On LIQUIDACIONSUELDOS.CodigoConcepto = CONCEPTOS.Codigo" _
    & " Where PUESTOLABORAL = " & "'" & PuestoLaboral & "'" _
    & " And CodigoLiquidacion = " & "'" & CodigoLiquidacion & "'" _
    & " And CodigoConcepto > " & "'" & "0200" & "'" _
    & " And CodigoConcepto < " & "'" & "0600" & "'" _
    & " And CodigoConcepto <> " & "'" & "0209" & "'" _
    & " And CodigoConcepto <> " & "'" & "0213" & "'" _
    & " And CodigoConcepto <> " & "'" & "0381" & "'" _
    & " Order By CodigoConcepto Asc"
    If SQLNoMatch(SQL) = False Then
        Set rstListadoSlave = New ADODB.Recordset
        rstListadoSlave.Open SQL, dbSlave, adOpenDynamic, adLockOptimistic
        With ReciboDeSueldo.dgDescuentosLiquidados
            If rstListadoSlave.EOF = False Then
                rstListadoSlave.MoveFirst
                While rstListadoSlave.EOF = False
                    i = i + 1
                    .RowHeight(i) = 300
                    .TextMatrix(i, 0) = rstListadoSlave!CodigoConcepto
                    .TextMatrix(i, 1) = rstListadoSlave!Denominacion
                    .TextMatrix(i, 2) = FormatNumber(rstListadoSlave!Importe, 2)
                    SumaTotal = SumaTotal + rstListadoSlave!Importe
                    rstListadoSlave.MoveNext
                    .Rows = .Rows + 1
                Wend
                ReciboDeSueldo.txtTotalDescuentos.Text = FormatNumber(SumaTotal, 2)
                If ReciboDeSueldo.txtSueldoNeto.Text <> "" Then
                    ReciboDeSueldo.txtSueldoNeto.Text = FormatNumber(ReciboDeSueldo.txtSueldoNeto.Text _
                    - FormatNumber(SumaTotal, 2), 2)
                Else
                    ReciboDeSueldo.txtTotalHaberes.Text = FormatNumber("0", 2)
                    ReciboDeSueldo.txtSueldoNeto.Text = FormatNumber("0", 2)
                End If
            End If
            .Rows = .Rows - 1
        End With
        rstListadoSlave.Close
        Set rstListadoSlave = Nothing
    Else
        ConfigurardgDescuentosLiquidados
        ReciboDeSueldo.txtTotalDescuentos.Text = ""
    End If
    
End Sub

Public Sub ConfigurardgListadoHonorariosImputados()
    
    With ListadoHonorariosImputados.dgListadoHonorariosImputados
        .Clear
        .Cols = 3
        .Rows = 2
        .TextMatrix(0, 0) = "Prog-Proy-Act-Part"
        .ColWidth(0) = 1500
        .TextMatrix(0, 1) = "Nombre Completo"
        .ColWidth(1) = 3000
        .TextMatrix(0, 2) = "Monto"
        .ColWidth(2) = 1000
        .FixedCols = 0
        .FocusRect = flexFocusHeavy
        .HighLight = flexHighlightWithFocus
        .SelectionMode = flexSelectionByRow
        .AllowUserResizing = flexResizeColumns
        .ColAlignment(0) = 4
        .ColAlignment(1) = 1
        .ColAlignment(2) = 7
    End With
    
End Sub

Public Sub CargardgListadoHonorariosImputados(ComprobanteSIIF As String)
    
    Dim i As Integer
    Dim strEstructuraPresupuestaria As String
    Dim SQL As String
    Dim SumaTotal As Double
    
    i = 0
    ListadoHonorariosImputados.dgListadoHonorariosImputados.Rows = 2
    Set rstListadoSlave = New ADODB.Recordset
    SQL = "Select PROVEEDOR, MONTOBRUTO, ACTIVIDAD, PARTIDA From LIQUIDACIONHONORARIOS " _
    & "Where COMPROBANTE = '" & ComprobanteSIIF & "' " _
    & "Order By ACTIVIDAD Asc, PARTIDA Asc, PROVEEDOR Asc"
    rstListadoSlave.Open SQL, dbSlave, adOpenDynamic, adLockOptimistic
    With ListadoHonorariosImputados.dgListadoHonorariosImputados
        If rstListadoSlave.BOF = False Then
            rstListadoSlave.MoveFirst
            While rstListadoSlave.EOF = False
                i = i + 1
                .RowHeight(i) = 300
                strEstructuraPresupuestaria = rstListadoSlave!ACTIVIDAD & "-" & rstListadoSlave!PARTIDA
                .TextMatrix(i, 0) = strEstructuraPresupuestaria
                .TextMatrix(i, 1) = rstListadoSlave!Proveedor
                .TextMatrix(i, 2) = FormatNumber(rstListadoSlave!MontoBruto, 2)
                SumaTotal = SumaTotal + rstListadoSlave!MontoBruto
                rstListadoSlave.MoveNext
                .Rows = .Rows + 1
            Wend
            ListadoHonorariosImputados.txtTotalImputacion.Text = FormatNumber(SumaTotal, 2)
        End If
        .Rows = .Rows - 1
        .SetFocus
    End With
    rstListadoSlave.Close
    Set rstListadoSlave = Nothing
    
End Sub

Public Sub ConfigurardgCodigosSIRADIG()
    
    With ListadoCodigosSIRADIG.dgCodigosSIRADIG
        .Clear
        .Cols = 2
        .Rows = 2
        .TextMatrix(0, 0) = "Código"
        .TextMatrix(0, 1) = "Denominación"
        .ColWidth(0) = 600
        .ColWidth(1) = 3400
        .FixedCols = 0
        .FocusRect = flexFocusHeavy
        .HighLight = flexHighlightWithFocus
        .SelectionMode = flexSelectionByRow
        .AllowUserResizing = flexResizeColumns
        .ColAlignment(0) = 4
        .ColAlignment(1) = 1
    End With
    
End Sub

Public Sub CargardgCodigosSIRADIG(Tabla As String)
    
    Dim i As Integer
    i = 0
    ListadoCodigosSIRADIG.dgCodigosSIRADIG.Rows = 2
    Set rstListadoSlave = New ADODB.Recordset
    Select Case Tabla
    Case "Parentesco"
        rstListadoSlave.Open "Select * From ParentescoSIRADIG Order By CODIGO", dbSlave, adOpenDynamic, adLockOptimistic
    Case "Deducciones"
        rstListadoSlave.Open "Select * From DeduccionesSIRADIG Order By CODIGO", dbSlave, adOpenDynamic, adLockOptimistic
    Case "OtrasDeducciones"
        rstListadoSlave.Open "Select * From OtrasDeduccionesSIRADIG Order By CODIGO", dbSlave, adOpenDynamic, adLockOptimistic
    End Select
    With ListadoCodigosSIRADIG.dgCodigosSIRADIG
        If rstListadoSlave.BOF = False Then
            rstListadoSlave.MoveFirst
            While rstListadoSlave.EOF = False
                i = i + 1
                .RowHeight(i) = 300
                .TextMatrix(i, 0) = rstListadoSlave!Codigo
                .TextMatrix(i, 1) = rstListadoSlave!Denominacion
                rstListadoSlave.MoveNext
                .Rows = .Rows + 1
            Wend
        End If
        .Rows = .Rows - 1
        .SetFocus
    End With
    rstListadoSlave.Close
    Set rstListadoSlave = Nothing
    
End Sub

Public Sub ConfigurardgPresentacionesF572()
    
    With ListadoF572.dgPresentacionesF572
        .Clear
        .Cols = 4
        .Rows = 2
        .TextMatrix(0, 0) = "ID"
        .TextMatrix(0, 1) = "Período"
        .TextMatrix(0, 2) = "Nro"
        .TextMatrix(0, 3) = "Fecha"
        .ColWidth(0) = 1
        .ColWidth(1) = 750
        .ColWidth(2) = 500
        .ColWidth(3) = 1000
        .FixedCols = 0
        .FocusRect = flexFocusHeavy
        .HighLight = flexHighlightWithFocus
        .SelectionMode = flexSelectionByRow
        .AllowUserResizing = flexResizeColumns
        .ColAlignment(0) = 1
        .ColAlignment(1) = 4
        .ColAlignment(2) = 4
        .ColAlignment(3) = 4
    End With
    
End Sub

Public Sub CargardgPresentacionesF572(CUIL As String)
    
    Dim SQL As String
    Dim i As Integer
    i = 0
    
    ListadoF572.dgPresentacionesF572.Rows = 2
    SQL = "Select * From PresentacionSIRADIG " _
        & "Where CUIL = '" & CUIL & "' " _
        & "Order by Right(ID,2) Desc, NroPresentacion Desc"
    Set rstListadoSlave = New ADODB.Recordset
    rstListadoSlave.Open SQL, dbSlave, adOpenDynamic, adLockOptimistic
    With ListadoF572.dgPresentacionesF572
        If rstListadoSlave.BOF = False Then
            rstListadoSlave.MoveFirst
            While rstListadoSlave.EOF = False
                i = i + 1
                .RowHeight(i) = 300
                .TextMatrix(i, 0) = rstListadoSlave!ID
                .TextMatrix(i, 1) = "20" & Right(rstListadoSlave!ID, 2)
                .TextMatrix(i, 2) = rstListadoSlave!NroPresentacion
                .TextMatrix(i, 3) = rstListadoSlave!Fecha
                rstListadoSlave.MoveNext
                .Rows = .Rows + 1
            Wend
        End If
        .Rows = .Rows - 1
        .SetFocus
    End With
    rstListadoSlave.Close
    Set rstListadoSlave = Nothing
    
End Sub

Public Sub ConfigurardgCargasDeFamiliaF572()
    
    With ListadoF572.dgCargasDeFamilia
        .Clear
        .Cols = 7
        .Rows = 2
        .TextMatrix(0, 0) = "CUIL"
        .TextMatrix(0, 1) = "Nombre Completo"
        .TextMatrix(0, 2) = "Parentesco"
        .TextMatrix(0, 3) = "Fecha Nac."
        .TextMatrix(0, 4) = "Desde"
        .TextMatrix(0, 5) = "Hasta"
        .TextMatrix(0, 6) = "¿Próximo?"
        .ColWidth(0) = 1100
        .ColWidth(1) = 2500
        .ColWidth(2) = 2000
        .ColWidth(3) = 1000
        .ColWidth(4) = 600
        .ColWidth(5) = 600
        .ColWidth(6) = 1000
        .FixedCols = 0
        .FocusRect = flexFocusHeavy
        .HighLight = flexHighlightWithFocus
        .SelectionMode = flexSelectionByRow
        .AllowUserResizing = flexResizeColumns
        .ColAlignment(0) = 1
        .ColAlignment(1) = 1
        .ColAlignment(2) = 7
        .ColAlignment(3) = 4
        .ColAlignment(4) = 7
        .ColAlignment(5) = 7
        .ColAlignment(6) = 7
    End With
    
End Sub

Public Sub CargardgCargasDeFamiliaF572(ID As String)
    
    Dim SQL As String
    Dim i As Integer
    i = 0
    
    ListadoF572.dgCargasDeFamilia.Rows = 2
    SQL = "Select * From CargasFamiliaSIRADIG " _
        & "Where ID = '" & ID & "' " _
        & "Order by CodigoParentesco Desc"
    Set rstListadoSlave = New ADODB.Recordset
    rstListadoSlave.Open SQL, dbSlave, adOpenDynamic, adLockOptimistic
    With ListadoF572.dgCargasDeFamilia
        If rstListadoSlave.BOF = False Then
            rstListadoSlave.MoveFirst
            While rstListadoSlave.EOF = False
                i = i + 1
                .RowHeight(i) = 300
                .TextMatrix(i, 0) = rstListadoSlave!CUIL
                SQL = "Select CargasDeFamilia.NombreCompleto From " _
                    & "(Agentes INNER JOIN CargasDeFamilia " _
                    & "ON Agentes.PuestoLaboral = CargasDeFamilia.PuestoLaboral) " _
                    & "INNER JOIN PresentacionSIRADIG " _
                    & "ON Agentes.CUIL = PresentacionSIRADIG.CUIL " _
                    & "Where PresentacionSIRADIG.ID = '" & ID & "' " _
                    & "And CargasDeFamilia.DNI = '" & Mid(rstListadoSlave!CUIL, 3, 8) & "'"
                If SQLNoMatch(SQL) = False Then
                    Set rstBuscarSlave = New ADODB.Recordset
                    rstBuscarSlave.Open SQL, dbSlave, adOpenDynamic, adLockReadOnly
                    .TextMatrix(i, 1) = rstBuscarSlave!NombreCompleto
                    rstBuscarSlave.Close
                    Set rstBuscarSlave = Nothing
                Else
                    .TextMatrix(i, 1) = "No Existe Familiar DB"
                End If
                SQL = "Select Denominacion From ParentescoSIRADIG " _
                    & "Where Codigo = '" & rstListadoSlave!CodigoParentesco & "'"
                If SQLNoMatch(SQL) = False Then
                    Set rstBuscarSlave = New ADODB.Recordset
                    rstBuscarSlave.Open SQL, dbSlave, adOpenDynamic, adLockReadOnly
                    .TextMatrix(i, 2) = rstBuscarSlave!Denominacion
                    rstBuscarSlave.Close
                    Set rstBuscarSlave = Nothing
                Else
                    .TextMatrix(i, 2) = "Código (" & rstListadoSlave!CodigoParentesco & ") no registrado"
                End If
                .TextMatrix(i, 3) = rstListadoSlave!FechaNacimiento
                .TextMatrix(i, 4) = Format(Format("01/" & Format(rstListadoSlave!MesDesde, "00") _
                & "/2009", "mmm"), ">")
                .TextMatrix(i, 5) = Format(Format("01/" & Format(rstListadoSlave!MesHasta, "00") _
                & "/2009", "mmm"), ">")
                .TextMatrix(i, 6) = rstListadoSlave!ProximoPeriodo
                rstListadoSlave.MoveNext
                .Rows = .Rows + 1
            Wend
        End If
        .Rows = .Rows - 1
        .SetFocus
    End With
    rstListadoSlave.Close
    Set rstListadoSlave = Nothing
    
End Sub

Public Sub ConfigurardgDeduccionesGeneralesF572()
    
    With ListadoF572.dgDeduccionesGenerales
        .Clear
        .Cols = 14
        .Rows = 2
        .TextMatrix(0, 0) = "Concepto"
        .ColWidth(0) = 2600
        .TextMatrix(0, 1) = "Total Anual"
        .ColWidth(1) = 1000
        For x = 1 To 12
            .TextMatrix(0, x + 1) = Format(Format("01/" & Format(x, "00") & "/2009", "mmmm"), ">")
            .ColWidth(x + 1) = 1000
            .ColAlignment(x + 1) = 7
        Next x
        .FixedCols = 0
        .FocusRect = flexFocusHeavy
        .HighLight = flexHighlightWithFocus
        .SelectionMode = flexSelectionByRow
        .AllowUserResizing = flexResizeColumns
        .ColAlignment(0) = 1
        .ColAlignment(1) = 7
    End With
    
End Sub

Public Sub CargardgDeduccionesGeneralesF572(ID As String)
    
    Dim SQL As String
    Dim i As Integer
    Dim dblSumaTotal As Double
    i = 0
    
    ListadoF572.dgDeduccionesGenerales.Rows = 2
    SQL = "Select * From DeduccionesGeneralesSIRADIG " _
        & "Where ID = '" & ID & "' " _
        & "Order by CodigoDeduccion Desc"
    Set rstListadoSlave = New ADODB.Recordset
    rstListadoSlave.Open SQL, dbSlave, adOpenDynamic, adLockOptimistic
    With ListadoF572.dgDeduccionesGenerales
        If rstListadoSlave.BOF = False Then
            rstListadoSlave.MoveFirst
            While rstListadoSlave.EOF = False
                i = i + 1
                dblSumaTotal = 0
                .RowHeight(i) = 300
                SQL = "Select Denominacion From DeduccionesSIRADIG " _
                    & "Where Codigo = '" & rstListadoSlave!CodigoDeduccion & "'"
                If SQLNoMatch(SQL) = False Then
                    Set rstBuscarSlave = New ADODB.Recordset
                    rstBuscarSlave.Open SQL, dbSlave, adOpenDynamic, adLockReadOnly
                    .TextMatrix(i, 0) = rstBuscarSlave!Denominacion
                    rstBuscarSlave.Close
                    Set rstBuscarSlave = Nothing
                Else
                    .TextMatrix(i, 0) = "Código (" & rstListadoSlave!CodigoDeduccion & ") no registrado"
                End If
                For x = 1 To 12
                    .TextMatrix(i, x + 1) = FormatNumber(rstListadoSlave.Fields(x + 1), 2)
                    dblSumaTotal = dblSumaTotal + rstListadoSlave.Fields(x + 1)
                Next x
                .TextMatrix(i, 1) = FormatNumber(CStr(dblSumaTotal), 2)
                rstListadoSlave.MoveNext
                .Rows = .Rows + 1
            Wend
        End If
        .Rows = .Rows - 1
        .SetFocus
    End With
    rstListadoSlave.Close
    Set rstListadoSlave = Nothing
    
End Sub

Public Sub ConfigurardgDeduccionesGeneralesLG4taSIRADIG()
    
    With LiquidacionGanancia4taSIRADIG.dgDeduccionesGenerales
        .Clear
        .Cols = 4
        .Rows = 2
        .TextMatrix(0, 0) = "ID"
        .TextMatrix(0, 1) = "Concepto"
        .TextMatrix(0, 2) = "Mensual"
        .TextMatrix(0, 3) = "Acumulado"
        .ColWidth(0) = 1
        .ColWidth(1) = 2600
        .ColWidth(2) = 1000
        .ColWidth(3) = 1000
        .FixedCols = 0
        .FocusRect = flexFocusHeavy
        .HighLight = flexHighlightWithFocus
        .SelectionMode = flexSelectionByRow
        .AllowUserResizing = flexResizeColumns
        .ColAlignment(0) = 1
        .ColAlignment(1) = 1
        .ColAlignment(2) = 7
        .ColAlignment(3) = 7
    End With
    
End Sub

Public Sub CargardgDeduccionesGeneralesLG4taSIRADIG(ID As String, Concepto As String, ImporteMensual As Double, _
ImporteAcumulado As Double)
    
    Dim i As Integer
    Dim SQL As String
    Dim strPeriodo As String
    Dim strCodigoSIRADIG As String
    
    With LiquidacionGanancia4taSIRADIG.dgDeduccionesGenerales
        'Determinamos el número de filas
        i = .Rows
        'Verificamos si tiene más de dos filas
        If i > 2 Then
            'Agregamos una fila
            .Rows = i + 1
            i = .Rows
        Else
            'Verificamos si la segunda fila tiene datos
            If .TextMatrix(1, 0) <> "" Then
                'Si ya tiene datos, agregamos una fila
                .Rows = i + 1
                i = .Rows
            End If
        End If
        
        'Le restamos 1 a i porque el indíce de filas arranca de 0
        i = i - 1
        
        'A partir de aquí inicia la carga de datos
        .TextMatrix(i, 0) = ID
        .TextMatrix(i, 1) = Concepto
        .TextMatrix(i, 2) = De_Num_a_Tx_01(Round(ImporteMensual, 2))
        .TextMatrix(i, 3) = De_Num_a_Tx_01(Round(ImporteAcumulado, 2))
    
    End With
    
    i = 0
    
End Sub

Public Sub ConfigurardgDeduccionesPersonalesLG4taSIRADIG()
    
    With LiquidacionGanancia4taSIRADIG.dgDeduccionesPersonales
        .Clear
        .Cols = 4
        .Rows = 2
        .TextMatrix(0, 0) = "ID"
        .TextMatrix(0, 1) = "Concepto"
        .TextMatrix(0, 2) = "Mensual"
        .TextMatrix(0, 3) = "Acumulado"
        .ColWidth(0) = 1
        .ColWidth(1) = 2600
        .ColWidth(2) = 1000
        .ColWidth(3) = 1000
        .FixedCols = 0
        .FocusRect = flexFocusHeavy
        .HighLight = flexHighlightWithFocus
        .SelectionMode = flexSelectionByRow
        .AllowUserResizing = flexResizeColumns
        .ColAlignment(0) = 1
        .ColAlignment(1) = 1
        .ColAlignment(2) = 7
        .ColAlignment(3) = 7
    End With
    
End Sub

Public Sub CargardgDeduccionesPersonalesLG4taSIRADIG(ID As String, Concepto As String, ImporteMensual As Double, _
ImporteAcumulado As Double)
    
    Dim i As Integer
    Dim SQL As String
    Dim strPeriodo As String
    Dim strCodigoSIRADIG As String
    
    With LiquidacionGanancia4taSIRADIG.dgDeduccionesPersonales
        'Determinamos el número de filas
        i = .Rows
        'Verificamos si tiene más de dos filas
        If i > 2 Then
            'Agregamos una fila
            .Rows = i + 1
            i = .Rows
        Else
            'Verificamos si la segunda fila tiene datos
            If .TextMatrix(1, 0) <> "" Then
                'Si ya tiene datos, agregamos una fila
                .Rows = i + 1
                i = .Rows
            End If
        End If
        
        'Le restamos 1 a i porque el indíce de filas arranca de 0
        i = i - 1
        
        'A partir de aquí inicia la carga de datos
        .TextMatrix(i, 0) = ID
        .TextMatrix(i, 1) = Concepto
        .TextMatrix(i, 2) = De_Num_a_Tx_01(Round(ImporteMensual, 2))
        .TextMatrix(i, 3) = De_Num_a_Tx_01(Round(ImporteAcumulado, 2))
    
    End With
    
    i = 0
    
End Sub

