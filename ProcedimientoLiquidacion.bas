Attribute VB_Name = "ProcedimientoLiquidacion"
Public Function CalcularConceptoSISPERAcumulado(PuestoLaboral As String, _
CodigoLiquidacion As String, CodigoConcepto As String, _
Optional IncluirLiquidacionActual As Boolean = True) As Double

    Dim SQL As String
    Dim dblConceptoCalculado As Double
    Dim strPL As String
    Dim strSimboloSQL As String
    
    strPL = BuscarPeriodoLiquidacion(CodigoLiquidacion)
    
    If IncluirLiquidacionActual = True Then
        strSimboloSQL = "<="
    Else
        strSimboloSQL = "<"
    End If
    
    'Buscamos el Concepto Acumulado
    Set rstProcedimientoSlave = New ADODB.Recordset
    SQL = "Select Sum(Importe) As TotalImporte" _
    & " From LIQUIDACIONSUELDOS Inner Join CODIGOLIQUIDACIONES" _
    & " ON LIQUIDACIONSUELDOS.CODIGOLIQUIDACION = CODIGOLIQUIDACIONES.CODIGO" _
    & " Where PUESTOLABORAL = '" & PuestoLaboral _
    & "' And CODIGOLIQUIDACION " & strSimboloSQL & " '" & CodigoLiquidacion _
    & "' And Right(PERIODO,4) = '" & Right(strPL, 4) _
    & "' And CODIGOCONCEPTO = '" & CodigoConcepto & "'"
    If SQLNoMatch(SQL) = False Then
        rstProcedimientoSlave.Open SQL, dbSlave, adOpenForwardOnly, adLockOptimistic
        dblConceptoCalculado = rstProcedimientoSlave!TotalImporte
        rstProcedimientoSlave.Close
    Else
        dblConceptoCalculado = 0
    End If

    
    CalcularConceptoSISPERAcumulado = Round(dblConceptoCalculado, 2)
    
    SQL = ""
    strPL = ""
    dblConceptoCalculado = 0
    If rstProcedimientoSlave.State = adStateOpen Then
        rstProcedimientoSlave.Close
    End If
    Set rstProcedimientoSlave = Nothing
   
End Function

Public Function CalcularHaberBruto(PuestoLaboral As String, CodigoLiquidacion As String) As Double

    Dim SQL As String
    Dim dblHBCalculado As Double
    Dim strPL As String
        
    strPL = BuscarPeriodoLiquidacion(CodigoLiquidacion)
    
    Set rstProcedimientoSlave = New ADODB.Recordset
    SQL = "Select * from LIQUIDACIONSUELDOS" _
    & " Where PUESTOLABORAL = '" & PuestoLaboral _
    & "' And CODIGOLIQUIDACION = '" & CodigoLiquidacion _
    & "' And CODIGOCONCEPTO = '9998'"
    If SQLNoMatch(SQL) = False Then
        rstProcedimientoSlave.Open SQL, dbSlave, adOpenForwardOnly, adLockOptimistic
        dblHBCalculado = rstProcedimientoSlave!Importe
        rstProcedimientoSlave.Close
        'Verificamos si la liquidación continene asignaciones familiares y las deducimos
        SQL = "Select * from LIQUIDACIONSUELDOS" _
        & " Where PUESTOLABORAL = '" & PuestoLaboral _
        & "' And CODIGOLIQUIDACION = '" & CodigoLiquidacion _
        & "' And CODIGOCONCEPTO = '0003'"
        If SQLNoMatch(SQL) = False Then
            'Si existe, buscamos el importe y lo restamos al Total Bruto
            rstProcedimientoSlave.Open SQL, dbSlave, adOpenForwardOnly, adLockOptimistic
            dblHBCalculado = dblHBCalculado - rstProcedimientoSlave!Importe
            rstProcedimientoSlave.Close
        End If
        'Verificamos si la liquidación continene discapacitado y las deducimos
        SQL = "Select * from LIQUIDACIONSUELDOS" _
        & " Where PUESTOLABORAL = '" & PuestoLaboral _
        & "' And CODIGOLIQUIDACION = '" & CodigoLiquidacion _
        & "' And CODIGOCONCEPTO = '0058'"
        If SQLNoMatch(SQL) = False Then
            'Si existe, buscamos el importe y lo restamos al Total Bruto
            rstProcedimientoSlave.Open SQL, dbSlave, adOpenForwardOnly, adLockOptimistic
            dblHBCalculado = dblHBCalculado - rstProcedimientoSlave!Importe
            rstProcedimientoSlave.Close
        End If
        'Verificamos si la liquidación continene DIF. SALARIO FLIAR.
        SQL = "Select * from LIQUIDACIONSUELDOS" _
        & " Where PUESTOLABORAL = '" & PuestoLaboral _
        & "' And CODIGOLIQUIDACION = '" & CodigoLiquidacion _
        & "' And CODIGOCONCEPTO = '0104'"
        If SQLNoMatch(SQL) = False Then
            'Si existe, buscamos el importe y lo restamos al Total Bruto
            rstProcedimientoSlave.Open SQL, dbSlave, adOpenForwardOnly, adLockOptimistic
            dblHBCalculado = dblHBCalculado - rstProcedimientoSlave!Importe
            rstProcedimientoSlave.Close
        End If
        'Incremento de 1/12 del Haber Bruto según Ley 27346/16 (Macri)
        If CInt(Right(strPL, 4)) >= 2017 Then
            dblHBCalculado = dblHBCalculado * (13 / 12)
        End If
    Else
        dblHBCalculado = 0
    End If
    
    CalcularHaberBruto = Round(dblHBCalculado, 2)
    
    SQL = ""
    strPL = ""
    dblHBCalculado = 0
    If rstProcedimientoSlave.State = adStateOpen Then
        rstProcedimientoSlave.Close
    End If
    Set rstProcedimientoSlave = Nothing
   
End Function

Public Function CalcularHaberBrutoAcumulado(PuestoLaboral As String, _
CodigoLiquidacion As String, Optional IncluirLiquidacionActual As Boolean = True, _
Optional IncluirSAC As Boolean = False) As Double

    Dim SQL As String
    Dim dblHBCalculado As Double
    Dim strPL As String
    Dim strSimboloSQL As String
        
    strPL = BuscarPeriodoLiquidacion(CodigoLiquidacion)
    
    If IncluirLiquidacionActual = True Then
        strSimboloSQL = "<="
    Else
        strSimboloSQL = "<"
    End If
    
    If IncluirSAC = True Then
        SQL = ""
    Else 'VERIFICAR PORQUE PUEDE TRAER PROBLEMAS CUANDO HAY AJUSTES RETROACTIVOS DE SAC
        SQL = " And CODIGOLIQUIDACION Not In" _
        & " (Select CODIGOLIQUIDACION From" _
        & " LIQUIDACIONSUELDOS Inner Join CODIGOLIQUIDACIONES" _
        & " ON LIQUIDACIONSUELDOS.CODIGOLIQUIDACION = CODIGOLIQUIDACIONES.CODIGO" _
        & " Where PUESTOLABORAL = '" & PuestoLaboral _
        & "' And Right(PERIODO,4) = '" & Right(strPL, 4) _
        & "' And CODIGOCONCEPTO = '0150')"
    End If
    
    Set rstProcedimientoSlave = New ADODB.Recordset
    SQL = "Select Sum(Importe) As TotalImporte" _
    & " From LIQUIDACIONSUELDOS Inner Join CODIGOLIQUIDACIONES" _
    & " ON LIQUIDACIONSUELDOS.CODIGOLIQUIDACION = CODIGOLIQUIDACIONES.CODIGO" _
    & " Where PUESTOLABORAL = '" & PuestoLaboral _
    & "' And CODIGOLIQUIDACION " & strSimboloSQL & " '" & CodigoLiquidacion _
    & "' And Right(PERIODO,4) = '" & Right(strPL, 4) _
    & "' And CODIGOCONCEPTO = '9998'" _
    & SQL 'Para el caso que no se desee incluir el SAC
    If SQLNoMatch(SQL) = False Then
        rstProcedimientoSlave.Open SQL, dbSlave, adOpenForwardOnly, adLockOptimistic
        dblHBCalculado = rstProcedimientoSlave!TotalImporte
        rstProcedimientoSlave.Close
        'Verificamos si la liquidación continene asignaciones familiares y las deducimos
        SQL = "Select Sum(Importe) As TotalImporte" _
        & " From LIQUIDACIONSUELDOS Inner Join CODIGOLIQUIDACIONES" _
        & " ON LIQUIDACIONSUELDOS.CODIGOLIQUIDACION = CODIGOLIQUIDACIONES.CODIGO" _
        & " Where PUESTOLABORAL = '" & PuestoLaboral _
        & "' And CODIGOLIQUIDACION " & strSimboloSQL & " '" & CodigoLiquidacion _
        & "' And Right(PERIODO,4) = '" & Right(strPL, 4) _
        & "' And CODIGOCONCEPTO = '0003'"
        If SQLNoMatch(SQL) = False Then
            'Si existe, buscamos el importe y lo restamos al Total Bruto
            rstProcedimientoSlave.Open SQL, dbSlave, adOpenForwardOnly, adLockOptimistic
            dblHBCalculado = dblHBCalculado - rstProcedimientoSlave!TotalImporte
            rstProcedimientoSlave.Close
        End If
        'Verificamos si la liquidación continene ajuste asignaciones familiares y las deducimos
        SQL = "Select Sum(Importe) As TotalImporte" _
        & " From LIQUIDACIONSUELDOS Inner Join CODIGOLIQUIDACIONES" _
        & " ON LIQUIDACIONSUELDOS.CODIGOLIQUIDACION = CODIGOLIQUIDACIONES.CODIGO" _
        & " Where PUESTOLABORAL = '" & PuestoLaboral _
        & "' And CODIGOLIQUIDACION " & strSimboloSQL & " '" & CodigoLiquidacion _
        & "' And Right(PERIODO,4) = '" & Right(strPL, 4) _
        & "' And CODIGOCONCEPTO = '0104'"
        If SQLNoMatch(SQL) = False Then
            'Si existe, buscamos el importe y lo restamos al Total Bruto
            rstProcedimientoSlave.Open SQL, dbSlave, adOpenForwardOnly, adLockOptimistic
            dblHBCalculado = dblHBCalculado - rstProcedimientoSlave!TotalImporte
            rstProcedimientoSlave.Close
        End If
        'Verificamos si la liquidación continene discapacitado y las deducimos
        SQL = "Select Sum(Importe) As TotalImporte" _
        & " From LIQUIDACIONSUELDOS Inner Join CODIGOLIQUIDACIONES" _
        & " ON LIQUIDACIONSUELDOS.CODIGOLIQUIDACION = CODIGOLIQUIDACIONES.CODIGO" _
        & " Where PUESTOLABORAL = '" & PuestoLaboral _
        & "' And CODIGOLIQUIDACION " & strSimboloSQL & " '" & CodigoLiquidacion _
        & "' And Right(PERIODO,4) = '" & Right(strPL, 4) _
        & "' And CODIGOCONCEPTO = '0058'"
        If SQLNoMatch(SQL) = False Then
            'Si existe, buscamos el importe y lo restamos al Total Bruto
            rstProcedimientoSlave.Open SQL, dbSlave, adOpenForwardOnly, adLockOptimistic
            dblHBCalculado = dblHBCalculado - rstProcedimientoSlave!TotalImporte
            rstProcedimientoSlave.Close
        End If
        'Incremento de 1/12 del Haber Bruto según Ley 27346/16 (Macri)
'        If CInt(Right(strPL, 4)) >= 2017 Then
'            dblHBCalculado = dblHBCalculado * (13 / 12)
'        End If
    Else
        dblHBCalculado = 0
    End If
    
    CalcularHaberBrutoAcumulado = Round(dblHBCalculado, 2)
    
    SQL = ""
    strPL = ""
    dblHBCalculado = 0
    If rstProcedimientoSlave.State = adStateOpen Then
        rstProcedimientoSlave.Close
    End If
    Set rstProcedimientoSlave = Nothing
   
End Function

Public Function CalcularRetribucionNoHabitualAcumulada(PuestoLaboral As String, _
CodigoLiquidacion As String, Optional IncluirLiquidacionActual As Boolean = True) As Double

    Dim dblRetribucionCalculada As Double
    
    dblRetribucionCalculada = CalcularConceptoSISPERAcumulado(PuestoLaboral, _
    CodigoLiquidacion, "0102", IncluirLiquidacionActual)
    dblRetribucionCalculada = dblRetribucionCalculada + CalcularConceptoSISPERAcumulado(PuestoLaboral, _
    CodigoLiquidacion, "0103", IncluirLiquidacionActual)
    dblRetribucionCalculada = dblRetribucionCalculada + CalcularConceptoSISPERAcumulado(PuestoLaboral, _
    CodigoLiquidacion, "0107", IncluirLiquidacionActual)
    
    CalcularRetribucionNoHabitualAcumulada = dblRetribucionCalculada
    
    
    dblRetribucionCalculada = 0

   
End Function

Public Function CalcularRemuneracionNoAlcanzadaAcumulada(PuestoLaboral As String, _
CodigoLiquidacion As String, Optional IncluirLiquidacionActual As Boolean = True) As Double

    Dim dblRetribucionCalculada As Double
    
    dblRetribucionCalculada = CalcularConceptoSISPERAcumulado(PuestoLaboral, _
    CodigoLiquidacion, "0003", IncluirLiquidacionActual)
    dblRetribucionCalculada = dblRetribucionCalculada + CalcularConceptoSISPERAcumulado(PuestoLaboral, _
    CodigoLiquidacion, "0058", IncluirLiquidacionActual)
    dblRetribucionCalculada = dblRetribucionCalculada + CalcularConceptoSISPERAcumulado(PuestoLaboral, _
    CodigoLiquidacion, "0104", IncluirLiquidacionActual)
    
    CalcularRemuneracionNoAlcanzadaAcumulada = dblRetribucionCalculada
    
    
    dblRetribucionCalculada = 0

   
End Function

Public Function CalcularPluriempleo(PuestoLaboral As String, CodigoLiquidacion As String) As Double

    Dim SQL As String
    Dim dblPluriempleoCalculado As Double
    
    'Buscamos Sueldo Pluriempleo Neto en Liquidación Anterior
    Set rstProcedimientoSlave = New ADODB.Recordset
    SQL = "Select * from LIQUIDACIONGANANCIAS4TACATEGORIA" _
    & " Where PUESTOLABORAL = '" & PuestoLaboral _
    & "' And CODIGOLIQUIDACION <= '" & CodigoLiquidacion _
    & "' Order By CODIGOLIQUIDACION Desc"
    If SQLNoMatch(SQL) = False Then
        rstProcedimientoSlave.Open SQL, dbSlave, adOpenForwardOnly, adLockOptimistic
        rstProcedimientoSlave.MoveFirst
        dblPluriempleoCalculado = rstProcedimientoSlave!Pluriempleo
        rstProcedimientoSlave.Close
    Else
        dblPluriempleoCalculado = 0
    End If
    
    CalcularPluriempleo = Round(dblPluriempleoCalculado, 2)
    
    SQL = ""
    dblPluriempleoCalculado = 0
    If rstProcedimientoSlave.State = adStateOpen Then
        rstProcedimientoSlave.Close
    End If
    Set rstProcedimientoSlave = Nothing
   
End Function

Public Function CalcularPluriempleoAcumulado(PuestoLaboral As String, _
CodigoLiquidacion As String, Optional IncluirLiquidacionActual As Boolean = True) As Double

    Dim SQL As String
    Dim dblPluriempleoCalculado As Double
    Dim strPL As String
    Dim strSimboloSQL As String
    
    strPL = BuscarPeriodoLiquidacion(CodigoLiquidacion)
    
    If IncluirLiquidacionActual = True Then
        strSimboloSQL = "<="
    Else
        strSimboloSQL = "<"
    End If
    
    'Buscamos Sueldo Pluriempleo Acumulado
    Set rstProcedimientoSlave = New ADODB.Recordset
    SQL = "Select Sum(Pluriempleo) As TotalImporte" _
    & " From LIQUIDACIONGANANCIAS4TACATEGORIA Inner Join CODIGOLIQUIDACIONES" _
    & " ON LIQUIDACIONGANANCIAS4TACATEGORIA.CODIGOLIQUIDACION = CODIGOLIQUIDACIONES.CODIGO" _
    & " Where PUESTOLABORAL = '" & PuestoLaboral _
    & "' And CODIGOLIQUIDACION " & strSimboloSQL & " '" & CodigoLiquidacion _
    & "' And Right(PERIODO,4) = '" & Right(strPL, 4) & "'"

    If SQLNoMatch(SQL) = False Then
        rstProcedimientoSlave.Open SQL, dbSlave, adOpenForwardOnly, adLockOptimistic
        dblPluriempleoCalculado = rstProcedimientoSlave!TotalImporte
        rstProcedimientoSlave.Close
    Else
        dblPluriempleoCalculado = 0
    End If
    
    CalcularPluriempleoAcumulado = Round(dblPluriempleoCalculado, 2)
    
    SQL = ""
    strPL = ""
    dblPluriempleoCalculado = 0
    If rstProcedimientoSlave.State = adStateOpen Then
        rstProcedimientoSlave.Close
    End If
    Set rstProcedimientoSlave = Nothing
   
End Function

Public Function CalcularAjuste(PuestoLaboral As String, CodigoLiquidacion As String) As String

    'Ver en qué situación conviene usar
    CalcularAjuste = Round(0, 2)

End Function

Public Function CalcularSAC(PuestoLaboral As String, CodigoLiquidacion As String, _
SACaIncluir As String, Optional IncluirLiquidacionActual As Boolean = True) As Double

    Dim dblSACPrimerSemestre As Double
    Dim dblSACSegundoSemestre As Double
    Dim SQL As String
    Dim strPL As String
    Dim strSimboloSQL As String
    
    dblSACPrimerSemestre = 0
    dblSACSegundoSemestre = 0
    strPL = BuscarPeriodoLiquidacion(CodigoLiquidacion)
    
    If IncluirLiquidacionActual = True Then
        strSimboloSQL = "<="
    Else
        strSimboloSQL = "<"
    End If
    
    'Buscamos los SAC liquidados en el año
    Set rstProcedimientoSlave = New ADODB.Recordset
    SQL = "Select Periodo, Importe" _
    & " From LIQUIDACIONSUELDOS Inner Join CODIGOLIQUIDACIONES" _
    & " ON LIQUIDACIONSUELDOS.CODIGOLIQUIDACION = CODIGOLIQUIDACIONES.CODIGO" _
    & " Where PUESTOLABORAL = '" & PuestoLaboral _
    & "' And CODIGOLIQUIDACION " & strSimboloSQL & " '" & CodigoLiquidacion _
    & "' And Right(PERIODO,4) = '" & Right(strPL, 4) _
    & "' And CODIGOCONCEPTO = '0150'" _
    & " Order by Periodo"
    If SQLNoMatch(SQL) = False Then
        rstProcedimientoSlave.Open SQL, dbSlave, adOpenForwardOnly, adLockOptimistic
        rstProcedimientoSlave.MoveFirst
        While rstProcedimientoSlave.EOF = False
            strPL = Left(rstProcedimientoSlave!Periodo, 2)
            If strPL <= 7 Then
                If dblSACPrimerSemestre < rstProcedimientoSlave!Importe Then
                    dblSACPrimerSemestre = rstProcedimientoSlave!Importe
                End If
            Else
                If dblSACSegundoSemestre < rstProcedimientoSlave!Importe Then
                    dblSACSegundoSemestre = rstProcedimientoSlave!Importe
                End If
            End If
            rstProcedimientoSlave.MoveNext
        Wend
        rstProcedimientoSlave.Close
    End If
    
    CalcularSAC = 0
    
    Select Case SACaIncluir
    Case "PrimerSAC"
        CalcularSAC = dblSACPrimerSemestre
    Case "SegundoSAC"
        CalcularSAC = dblSACSegundoSemestre
    Case "Ambos"
        CalcularSAC = dblSACPrimerSemestre + dblSACSegundoSemestre
    End Select
    
    SQL = ""
    strPL = ""
    dblSACPrimerSemestre = 0
    dblSACSegundoSemestre = 0
    If rstProcedimientoSlave.State = adStateOpen Then
        rstProcedimientoSlave.Close
    End If
    Set rstProcedimientoSlave = Nothing
   
End Function

Public Function CalcularJubilacion(PuestoLaboral As String, CodigoLiquidacion As String) As Double

    Dim SQL As String
    Dim dblJubilacionCalculada As Double
    Dim strPL As String
    
    strPL = BuscarPeriodoLiquidacion(CodigoLiquidacion)
    
    'Buscamos el Descuento por Jubilación
    Set rstProcedimientoSlave = New ADODB.Recordset
    SQL = "Select * from LIQUIDACIONSUELDOS" _
    & " Where PUESTOLABORAL = '" & PuestoLaboral _
    & "' And CODIGOLIQUIDACION = '" & CodigoLiquidacion _
    & "' And CODIGOCONCEPTO = '0208'"
    If SQLNoMatch(SQL) = False Then
        rstProcedimientoSlave.Open SQL, dbSlave, adOpenForwardOnly, adLockOptimistic
        dblJubilacionCalculada = rstProcedimientoSlave!Importe
        rstProcedimientoSlave.Close
        'Incremento de 1/12 del Haber Bruto según Ley 27346/16 (Macri) _
        - La ley solo indica incremntar el Haber Bruto pero no tiene sentido no incremantar esto también -
        If CInt(Right(strPL, 4)) >= 2017 Then
            dblJubilacionCalculada = dblJubilacionCalculada * (13 / 12)
        End If
    Else
        dblJubilacionCalculada = 0
    End If

    
    CalcularJubilacion = Round(dblJubilacionCalculada, 2)
    
    SQL = ""
    strPL = ""
    dblJubilacionCalculada = 0
    If rstProcedimientoSlave.State = adStateOpen Then
        rstProcedimientoSlave.Close
    End If
    Set rstProcedimientoSlave = Nothing
   
End Function

Public Function CalcularJubilacionAcumulada(PuestoLaboral As String, _
CodigoLiquidacion As String, Optional IncluirLiquidacionActual As Boolean = True) As Double

    Dim SQL As String
    Dim dblJubilacionCalculada As Double
    Dim strPL As String
    Dim strSimboloSQL As String
    
    strPL = BuscarPeriodoLiquidacion(CodigoLiquidacion)
    
    If IncluirLiquidacionActual = True Then
        strSimboloSQL = "<="
    Else
        strSimboloSQL = "<"
    End If
    
    'Buscamos el Descuento por Jubilación Acumulado
    Set rstProcedimientoSlave = New ADODB.Recordset
    SQL = "Select Sum(Importe) As TotalImporte" _
    & " From LIQUIDACIONSUELDOS Inner Join CODIGOLIQUIDACIONES" _
    & " ON LIQUIDACIONSUELDOS.CODIGOLIQUIDACION = CODIGOLIQUIDACIONES.CODIGO" _
    & " Where PUESTOLABORAL = '" & PuestoLaboral _
    & "' And CODIGOLIQUIDACION " & strSimboloSQL & " '" & CodigoLiquidacion _
    & "' And Right(PERIODO,4) = '" & Right(strPL, 4) _
    & "' And CODIGOCONCEPTO = '0208'"
    If SQLNoMatch(SQL) = False Then
        rstProcedimientoSlave.Open SQL, dbSlave, adOpenForwardOnly, adLockOptimistic
        dblJubilacionCalculada = rstProcedimientoSlave!TotalImporte
        rstProcedimientoSlave.Close
    Else
        dblJubilacionCalculada = 0
    End If

    
    CalcularJubilacionAcumulada = Round(dblJubilacionCalculada, 2)
    
    SQL = ""
    strPL = ""
    dblJubilacionCalculada = 0
    If rstProcedimientoSlave.State = adStateOpen Then
        rstProcedimientoSlave.Close
    End If
    Set rstProcedimientoSlave = Nothing
   
End Function

Public Function CalcularObraSocial(PuestoLaboral As String, CodigoLiquidacion As String) As Double

    Dim SQL As String
    Dim dblObraSocialCalculada As Double
    Dim strPL As String
    
    strPL = BuscarPeriodoLiquidacion(CodigoLiquidacion)
    
    'Buscamos el Descuento por ObraSocial
    Set rstProcedimientoSlave = New ADODB.Recordset
    SQL = "Select * from LIQUIDACIONSUELDOS" _
    & " Where PUESTOLABORAL = '" & PuestoLaboral _
    & "' And CODIGOLIQUIDACION = '" & CodigoLiquidacion _
    & "' And CODIGOCONCEPTO = '0212'"
    If SQLNoMatch(SQL) = False Then
        rstProcedimientoSlave.Open SQL, dbSlave, adOpenForwardOnly, adLockOptimistic
        dblObraSocialCalculada = rstProcedimientoSlave!Importe
        rstProcedimientoSlave.Close
        'Incremento de 1/12 del Haber Bruto según Ley 27346/16 (Macri) _
        - La ley solo indica incremntar el Haber Bruto pero no tiene sentido no incremantar esto también -
        If CInt(Right(strPL, 4)) >= 2017 Then
            dblObraSocialCalculada = dblObraSocialCalculada * (13 / 12)
        End If
    Else
        dblObraSocialCalculada = 0
    End If
    CalcularObraSocial = Round(dblObraSocialCalculada, 2)
    
    SQL = ""
    strPL = ""
    dblObraSocialCalculada = 0
    If rstProcedimientoSlave.State = adStateOpen Then
        rstProcedimientoSlave.Close
    End If
    Set rstProcedimientoSlave = Nothing
   
End Function

Public Function CalcularObraSocialAcumulada(PuestoLaboral As String, _
CodigoLiquidacion As String, Optional IncluirLiquidacionActual As Boolean = True) As Double

    Dim SQL As String
    Dim dblObraSocialCalculada As Double
    Dim strPL As String
    Dim strSimboloSQL As String
    
    strPL = BuscarPeriodoLiquidacion(CodigoLiquidacion)
    
    If IncluirLiquidacionActual = True Then
        strSimboloSQL = "<="
    Else
        strSimboloSQL = "<"
    End If
    
    'Buscamos el Descuento por Jubilación Acumulado
    Set rstProcedimientoSlave = New ADODB.Recordset
    SQL = "Select Sum(Importe) As TotalImporte" _
    & " From LIQUIDACIONSUELDOS Inner Join CODIGOLIQUIDACIONES" _
    & " ON LIQUIDACIONSUELDOS.CODIGOLIQUIDACION = CODIGOLIQUIDACIONES.CODIGO" _
    & " Where PUESTOLABORAL = '" & PuestoLaboral _
    & "' And CODIGOLIQUIDACION " & strSimboloSQL & " '" & CodigoLiquidacion _
    & "' And Right(PERIODO,4) = '" & Right(strPL, 4) _
    & "' And CODIGOCONCEPTO = '0212'"
    If SQLNoMatch(SQL) = False Then
        rstProcedimientoSlave.Open SQL, dbSlave, adOpenForwardOnly, adLockOptimistic
        dblObraSocialCalculada = rstProcedimientoSlave!TotalImporte
        rstProcedimientoSlave.Close
    Else
        dblObraSocialCalculada = 0
    End If
    CalcularObraSocialAcumulada = Round(dblObraSocialCalculada, 2)
    
    SQL = ""
    strPL = ""
    dblObraSocialCalculada = 0
    If rstProcedimientoSlave.State = adStateOpen Then
        rstProcedimientoSlave.Close
    End If
    Set rstProcedimientoSlave = Nothing
   
End Function

Public Function CalcularAdherenteObraSocial(PuestoLaboral As String, CodigoLiquidacion As String) As Double

    Dim SQL As String
    Dim dblAOSCalculado As Double
    
    'Buscamos si existe Adherente Obra Social
    Set rstProcedimientoSlave = New ADODB.Recordset
    SQL = "Select * from LIQUIDACIONSUELDOS" _
    & " Where PUESTOLABORAL = '" & PuestoLaboral _
    & "' And CODIGOLIQUIDACION = '" & CodigoLiquidacion _
    & "' And CODIGOCONCEPTO = '0234'"
    If SQLNoMatch(SQL) = True Then
        'Si no existe Adherente Obra Social
        dblAOSCalculado = 0
    Else
        'Si existe Adherente Obra Social, buscamos el valor de la presente liquidacion
        rstProcedimientoSlave.Open SQL, dbSlave, adOpenForwardOnly, adLockOptimistic
        dblAOSCalculado = rstProcedimientoSlave!Importe
        rstProcedimientoSlave.Close
        'Luego, buscamos lo que se liquidó en la anterior liquidación (VERIFICAR PROCEDIMIENTO)
        SQL = "Select * from LIQUIDACIONGANANCIAS4TACATEGORIA" _
        & " Where PUESTOLABORAL = '" & PuestoLaboral _
        & "' And CODIGOLIQUIDACION <= '" & CodigoLiquidacion _
        & "' Order by CODIGOLIQUIDACION Desc"
        If SQLNoMatch(SQL) = False Then
            rstProcedimientoSlave.Open SQL, dbSlave, adOpenForwardOnly, adLockOptimistic
            rstProcedimientoSlave.MoveFirst
            'Si hay diferencia entre la liquidación previa y la actual, escojemos la liquidación previa
            If rstProcedimientoSlave!AdherenteObraSocial <> dblAOSCalculado Then
                dblAOSCalculado = rstProcedimientoSlave!AdherenteObraSocial
            End If
            rstProcedimientoSlave.Close
        End If
    End If
    
    CalcularAdherenteObraSocial = Round(dblAOSCalculado, 2)
    
    SQL = ""
    dblAOSCalculado = 0
    If rstProcedimientoSlave.State = adStateOpen Then
        rstProcedimientoSlave.Close
    End If
    Set rstProcedimientoSlave = Nothing
   
End Function

Public Function CalcularDescuentoSeguroDeVidaObligatorio(PuestoLaboral As String, CodigoLiquidacion As String) As Double

    Dim SQL As String
    Dim dblSVObligatorioCalculado As Double
    
    'Buscamos el Descuento Seguro Obligatorio
    Set rstProcedimientoSlave = New ADODB.Recordset
    SQL = "Select * from LIQUIDACIONSUELDOS" _
    & " Where ((PUESTOLABORAL = '" & PuestoLaboral _
    & "' And CODIGOLIQUIDACION = '" & CodigoLiquidacion _
    & "' And CODIGOCONCEPTO = '0360') Or " _
    & "(PUESTOLABORAL = '" & PuestoLaboral _
    & "' And CODIGOLIQUIDACION = '" & CodigoLiquidacion _
    & "' And CODIGOCONCEPTO = '0366') Or " _
    & "(PUESTOLABORAL = '" & PuestoLaboral _
    & "' And CODIGOLIQUIDACION = '" & CodigoLiquidacion _
    & "' And CODIGOCONCEPTO = '0369'))"
    If SQLNoMatch(SQL) = True Then
        'Si no existe descuento de Seguro Obligatorio
        dblSVObligatorioCalculado = 0
    Else
        'Si existe descuento de Seguro Obligatorio
        rstProcedimientoSlave.Open SQL, dbSlave, adOpenForwardOnly, adLockOptimistic
        dblSVObligatorioCalculado = rstProcedimientoSlave!Importe
        rstProcedimientoSlave.Close
    End If
    
    CalcularDescuentoSeguroDeVidaObligatorio = Round(dblSVObligatorioCalculado, 2)
    
    SQL = ""
    dblSVObligatorioCalculado = 0
    If rstProcedimientoSlave.State = adStateOpen Then
        rstProcedimientoSlave.Close
    End If
    Set rstProcedimientoSlave = Nothing
   
End Function

Public Function CalcularDescuentoSeguroDeVidaObligatorioAcumulado(PuestoLaboral As String, _
CodigoLiquidacion As String, Optional IncluirLiquidacionActual As Boolean = True) As Double

    Dim SQL As String
    Dim dblSVObligatorioCalculado As Double
    Dim strPL As String
    Dim strSimboloSQL As String
    
    strPL = BuscarPeriodoLiquidacion(CodigoLiquidacion)
    
    If IncluirLiquidacionActual = True Then
        strSimboloSQL = "<="
    Else
        strSimboloSQL = "<"
    End If
    
    'Buscamos el Descuento por Seguro Obligatorio Acumulado
    Set rstProcedimientoSlave = New ADODB.Recordset
    SQL = "Select Sum(Importe) As TotalImporte" _
    & " From LIQUIDACIONSUELDOS Inner Join CODIGOLIQUIDACIONES" _
    & " ON LIQUIDACIONSUELDOS.CODIGOLIQUIDACION = CODIGOLIQUIDACIONES.CODIGO" _
    & " Where ((PUESTOLABORAL = '" & PuestoLaboral _
    & "' And CODIGOLIQUIDACION " & strSimboloSQL & " '" & CodigoLiquidacion _
    & "' And Right(PERIODO,4) = '" & Right(strPL, 4) _
    & "' And CODIGOCONCEPTO = '0360') Or " _
    & "(PUESTOLABORAL = '" & PuestoLaboral _
    & "' And CODIGOLIQUIDACION " & strSimboloSQL & " '" & CodigoLiquidacion _
    & "' And Right(PERIODO,4) = '" & Right(strPL, 4) _
    & "' And CODIGOCONCEPTO = '0366') Or " _
    & "(PUESTOLABORAL = '" & PuestoLaboral _
    & "' And CODIGOLIQUIDACION " & strSimboloSQL & " '" & CodigoLiquidacion _
    & "' And Right(PERIODO,4) = '" & Right(strPL, 4) _
    & "' And CODIGOCONCEPTO = '0369'))"

    If SQLNoMatch(SQL) = False Then
        rstProcedimientoSlave.Open SQL, dbSlave, adOpenForwardOnly, adLockOptimistic
        dblSVObligatorioCalculado = rstProcedimientoSlave!TotalImporte
        rstProcedimientoSlave.Close
    Else
        dblSVObligatorioCalculado = 0
    End If

    
    CalcularDescuentoSeguroDeVidaObligatorioAcumulado = Round(dblSVObligatorioCalculado, 2)
    
    SQL = ""
    strPL = ""
    dblSVObligatorioCalculado = 0
    If rstProcedimientoSlave.State = adStateOpen Then
        rstProcedimientoSlave.Close
    End If
    Set rstProcedimientoSlave = Nothing
   
End Function

Public Function CalcularDescuentoCuotaSindical(PuestoLaboral As String, CodigoLiquidacion As String) As Double

    Dim SQL As String
    Dim dblCuotaSindicalCalculada As Double
    Dim strPL As String
    
    strPL = BuscarPeriodoLiquidacion(CodigoLiquidacion)
    
    'Buscamos el Descuento Cuota Sindical
    Set rstProcedimientoSlave = New ADODB.Recordset
    SQL = "Select * from LIQUIDACIONSUELDOS" _
    & " Where ((PUESTOLABORAL = '" & PuestoLaboral _
    & "' And CODIGOLIQUIDACION = '" & CodigoLiquidacion _
    & "' And CODIGOCONCEPTO = '0219') Or " _
    & "(PUESTOLABORAL = '" & PuestoLaboral _
    & "' And CODIGOLIQUIDACION = '" & CodigoLiquidacion _
    & "' And CODIGOCONCEPTO = '0227'))"
    If SQLNoMatch(SQL) = True Then
        'Si no existe descuento Cuota Sindical
        dblCuotaSindicalCalculada = 0
    Else
        'Si existe descuento Cuota Sindical
        rstProcedimientoSlave.Open SQL, dbSlave, adOpenForwardOnly, adLockOptimistic
        dblCuotaSindicalCalculada = rstProcedimientoSlave!Importe
        rstProcedimientoSlave.Close
        'Incremento de 1/12 del Haber Bruto según Ley 27346/16 (Macri) _
        - La ley solo indica incremntar el Haber Bruto pero no tiene sentido no incremantar esto también -
        If CInt(Right(strPL, 4)) >= 2017 Then
            dblCuotaSindicalCalculada = dblCuotaSindicalCalculada * (13 / 12)
        End If
    End If

    
    CalcularDescuentoCuotaSindical = Round(dblCuotaSindicalCalculada, 2)
    
    SQL = ""
    strPL = ""
    dblCuotaSindicalCalculada = 0
    If rstProcedimientoSlave.State = adStateOpen Then
        rstProcedimientoSlave.Close
    End If
    Set rstProcedimientoSlave = Nothing
   
End Function

Public Function CalcularDescuentoCuotaSindicalAcumulado(PuestoLaboral As String, _
CodigoLiquidacion As String, Optional IncluirLiquidacionActual As Boolean = True, _
Optional IncluirSAC As Boolean = False) As Double

    Dim SQL As String
    Dim dblCuotaSindicalCalculada As Double
    Dim strPL As String
    Dim strSimboloSQL As String
        
    strPL = BuscarPeriodoLiquidacion(CodigoLiquidacion)
    
    If IncluirLiquidacionActual = True Then
        strSimboloSQL = "<="
    Else
        strSimboloSQL = "<"
    End If

    If IncluirSAC = True Then
        SQL = ""
    Else
        SQL = " And CODIGOLIQUIDACION Not In" _
        & " (Select CODIGOLIQUIDACION From" _
        & " LIQUIDACIONSUELDOS Inner Join CODIGOLIQUIDACIONES" _
        & " ON LIQUIDACIONSUELDOS.CODIGOLIQUIDACION = CODIGOLIQUIDACIONES.CODIGO" _
        & " Where PUESTOLABORAL = '" & PuestoLaboral _
        & "' And Right(PERIODO,4) = '" & Right(strPL, 4) _
        & "' And CODIGOCONCEPTO = '0150')"
    End If

    strPL = BuscarPeriodoLiquidacion(CodigoLiquidacion)
    
    'Buscamos el Descuento Cuota Sindical
    Set rstProcedimientoSlave = New ADODB.Recordset
    SQL = "Select Sum(Importe) As TotalImporte" _
    & " From LIQUIDACIONSUELDOS Inner Join CODIGOLIQUIDACIONES" _
    & " ON LIQUIDACIONSUELDOS.CODIGOLIQUIDACION = CODIGOLIQUIDACIONES.CODIGO" _
    & " Where ((PUESTOLABORAL = '" & PuestoLaboral _
    & "' And CODIGOLIQUIDACION " & strSimboloSQL & " '" & CodigoLiquidacion _
    & "' And Right(PERIODO,4) = '" & Right(strPL, 4) _
    & "' And CODIGOCONCEPTO = '0219') Or " _
    & "(PUESTOLABORAL = '" & PuestoLaboral _
    & "' And CODIGOLIQUIDACION " & strSimboloSQL & " '" & CodigoLiquidacion _
    & "' And Right(PERIODO,4) = '" & Right(strPL, 4) _
    & "' And CODIGOCONCEPTO = '0227'))" _
    & SQL
    
    If SQLNoMatch(SQL) = True Then
        'Si no existe descuento Cuota Sindical
        dblCuotaSindicalCalculada = 0
    Else
        'Si existe descuento Cuota Sindical
        rstProcedimientoSlave.Open SQL, dbSlave, adOpenForwardOnly, adLockOptimistic
        dblCuotaSindicalCalculada = rstProcedimientoSlave!TotalImporte
        rstProcedimientoSlave.Close
        'Incremento de 1/12 del Haber Bruto según Ley 27346/16 (Macri) _
        - La ley solo indica incremntar el Haber Bruto pero no tiene sentido no incremantar esto también -
        If IncluirSAC = False And CInt(Right(strPL, 4)) >= 2017 Then
            dblCuotaSindicalCalculada = dblCuotaSindicalCalculada * (13 / 12)
        End If
    End If

    
    CalcularDescuentoCuotaSindicalAcumulado = Round(dblCuotaSindicalCalculada, 2)
    
    SQL = ""
    strPL = ""
    dblCuotaSindicalCalculada = 0
    If rstProcedimientoSlave.State = adStateOpen Then
        rstProcedimientoSlave.Close
    End If
    Set rstProcedimientoSlave = Nothing
   
End Function

Public Function CalcularDescuentoSeguroDeVidaOptativo(PuestoLaboral As String, CodigoLiquidacion As String) As Double

    Dim SQL As String
    Dim dblSVOptativoCalculado As Double
    Dim strPeriodo As String
    Dim datFecha As Date
     
    'Buscamos el Descuento Seguro Optativo
    Set rstProcedimientoSlave = New ADODB.Recordset
    SQL = "Select * from LIQUIDACIONSUELDOS" _
    & " Where ((PUESTOLABORAL = '" & PuestoLaboral _
    & "' And CODIGOLIQUIDACION = '" & CodigoLiquidacion _
    & "' And CODIGOCONCEPTO = '0317') Or " _
    & "(PUESTOLABORAL = '" & PuestoLaboral _
    & "' And CODIGOLIQUIDACION = '" & CodigoLiquidacion _
    & "' And CODIGOCONCEPTO = '0361') Or " _
    & "(PUESTOLABORAL = '" & PuestoLaboral _
    & "' And CODIGOLIQUIDACION = '" & CodigoLiquidacion _
    & "' And CODIGOCONCEPTO = '0367') Or " _
    & "(PUESTOLABORAL = '" & PuestoLaboral _
    & "' And CODIGOLIQUIDACION = '" & CodigoLiquidacion _
    & "' And CODIGOCONCEPTO = '0370') Or " _
    & "(PUESTOLABORAL = '" & PuestoLaboral _
    & "' And CODIGOLIQUIDACION = '" & CodigoLiquidacion _
    & "' And CODIGOCONCEPTO = '0373') Or " _
    & "(PUESTOLABORAL = '" & PuestoLaboral _
    & "' And CODIGOLIQUIDACION = '" & CodigoLiquidacion _
    & "' And CODIGOCONCEPTO = '0374'))"
    If SQLNoMatch(SQL) = True Then
        'Si no existe descuento Seguro Optativo
        dblSVOptativoCalculado = 0
    Else
        'Si existe descuento Seguro Optativo
        strPeriodo = BuscarPeriodoLiquidacion(CodigoLiquidacion)
        datFecha = BuscarUltimoDiaDelPeriodo(strPeriodo)
        rstProcedimientoSlave.Open SQL, dbSlave, adOpenForwardOnly, adLockOptimistic
        Set rstBuscarSlave = New ADODB.Recordset
        SQL = "Select SEGURODEVIDA From DEDUCCIONES4TACATEGORIA" _
        & " Where FECHA <= #" & Format(datFecha, "MM/DD/YYYY") _
        & "# Order by FECHA Desc"
        rstBuscarSlave.Open SQL, dbSlave, adOpenForwardOnly, adLockOptimistic
        If rstProcedimientoSlave!Importe < (rstBuscarSlave!SeguroDeVida / 12) Then
            dblSVOptativoCalculado = rstProcedimientoSlave!Importe
        Else
            dblSVOptativoCalculado = (rstBuscarSlave!SeguroDeVida / 12)
        End If
        rstProcedimientoSlave.Close
        rstBuscarSlave.Close
        Set rstBuscarSlave = Nothing
    End If
    
    CalcularDescuentoSeguroDeVidaOptativo = Round(dblSVOptativoCalculado, 2)
    
    SQL = ""
    dblSVOptativoCalculado = 0
    If rstProcedimientoSlave.State = adStateOpen Then
        rstProcedimientoSlave.Close
    End If
    Set rstProcedimientoSlave = Nothing
    strPeriodo = ""
    datFecha = 0
   
End Function

Public Function CalcularDescuentoSeguroDeVidaOptativoAcumulado(PuestoLaboral As String, _
CodigoLiquidacion As String, Optional IncluirLiquidacionActual As Boolean = True) As Double

    Dim SQL As String
    Dim strSimboloSQL As String
    Dim dblSVOptativoCalculado As Double
    Dim strPL As String
    Dim datFecha As Date
    Dim intCantidadMeses As Integer
     
    strPL = BuscarPeriodoLiquidacion(CodigoLiquidacion)
    datFecha = BuscarUltimoDiaDelPeriodo(strPL)
     
    If IncluirLiquidacionActual = True Then
        strSimboloSQL = "<="
        intCantidadMeses = Month(datFecha)
    Else
        strSimboloSQL = "<"
        intCantidadMeses = Month(datFecha) - 1
    End If
     
    'Buscamos el Descuento Seguro Optativo Acumulado
    Set rstProcedimientoSlave = New ADODB.Recordset
    SQL = "Select Sum(Importe) As TotalImporte" _
    & " From LIQUIDACIONSUELDOS Inner Join CODIGOLIQUIDACIONES" _
    & " ON LIQUIDACIONSUELDOS.CODIGOLIQUIDACION = CODIGOLIQUIDACIONES.CODIGO" _
    & " Where ((PUESTOLABORAL = '" & PuestoLaboral _
    & "' And CODIGOLIQUIDACION " & strSimboloSQL & " '" & CodigoLiquidacion _
    & "' And Right(PERIODO,4) = '" & Right(strPL, 4) & "' And CODIGOCONCEPTO = '0317') Or " _
    & "(PUESTOLABORAL = '" & PuestoLaboral _
    & "' And CODIGOLIQUIDACION " & strSimboloSQL & " '" & CodigoLiquidacion _
    & "' And Right(PERIODO,4) = '" & Right(strPL, 4) & "' And CODIGOCONCEPTO = '0361') Or " _
    & "(PUESTOLABORAL = '" & PuestoLaboral _
    & "' And CODIGOLIQUIDACION " & strSimboloSQL & " '" & CodigoLiquidacion _
    & "' And Right(PERIODO,4) = '" & Right(strPL, 4) & "' And CODIGOCONCEPTO = '0367') Or " _
    & "(PUESTOLABORAL = '" & PuestoLaboral _
    & "' And CODIGOLIQUIDACION " & strSimboloSQL & " '" & CodigoLiquidacion _
    & "' And Right(PERIODO,4) = '" & Right(strPL, 4) & "' And CODIGOCONCEPTO = '0370') Or " _
    & "(PUESTOLABORAL = '" & PuestoLaboral _
    & "' And CODIGOLIQUIDACION " & strSimboloSQL & " '" & CodigoLiquidacion _
    & "' And Right(PERIODO,4) = '" & Right(strPL, 4) & "' And CODIGOCONCEPTO = '0373') Or " _
    & "(PUESTOLABORAL = '" & PuestoLaboral _
    & "' And CODIGOLIQUIDACION " & strSimboloSQL & " '" & CodigoLiquidacion _
    & "' And Right(PERIODO,4) = '" & Right(strPL, 4) & "' And CODIGOCONCEPTO = '0374'))"
    If SQLNoMatch(SQL) = True Then
        'Si no existe descuento Seguro Optativo
        dblSVOptativoCalculado = 0
    Else
        'Si existe descuento Seguro Optativo
        rstProcedimientoSlave.Open SQL, dbSlave, adOpenForwardOnly, adLockOptimistic
        Set rstBuscarSlave = New ADODB.Recordset
        SQL = "Select SEGURODEVIDA From DEDUCCIONES4TACATEGORIA" _
        & " Where FECHA <= #" & Format(datFecha, "MM/DD/YYYY") _
        & "# Order by FECHA Desc"
        rstBuscarSlave.Open SQL, dbSlave, adOpenForwardOnly, adLockOptimistic
        If rstProcedimientoSlave!TotalImporte < (rstBuscarSlave!SeguroDeVida / 12 * intCantidadMeses) Then
            dblSVOptativoCalculado = rstProcedimientoSlave!TotalImporte
        Else
            dblSVOptativoCalculado = (rstBuscarSlave!SeguroDeVida / 12 * intCantidadMeses)
        End If
        rstProcedimientoSlave.Close
        rstBuscarSlave.Close
        Set rstBuscarSlave = Nothing
    End If
    
    CalcularDescuentoSeguroDeVidaOptativoAcumulado = Round(dblSVOptativoCalculado, 2)
    
    SQL = ""
    dblSVOptativoCalculado = 0
    If rstProcedimientoSlave.State = adStateOpen Then
        rstProcedimientoSlave.Close
    End If
    Set rstProcedimientoSlave = Nothing
    strPeriodo = ""
    datFecha = 0
   
End Function


Public Function CalcularOtrosDescuentos(PuestoLaboral As String, CodigoLiquidacion As String) As String

    'Ver en qué situación conviene usar
    CalcularOtrosDescuentos = Round(0, 2)

End Function

Public Function CalcularRentaAcumulada(PuestoLaboral As String, _
CodigoLiquidacion As String, Optional IncluirLiquidacionActual As Boolean = True) As Double

    Dim SQL As String
    Dim dblImporteCalculado As Double
    Dim strPeriodo As String
    Dim strSimboloSQL As String
    
    If IncluirLiquidacionActual = True Then
        strSimboloSQL = "<="
    Else
        strSimboloSQL = "<"
    End If
    
    'TENER CUIDADO con este PROCEDIMENTO!!! Renta Acumulada INCLUYENDO la presente liquidación
    strPeriodo = BuscarPeriodoLiquidacion(CodigoLiquidacion)
    Set rstProcedimientoSlave = New ADODB.Recordset
    'Cargamos la Renta Acumulada
    SQL = "Select Sum(HABEROPTIMO) as HO," _
    & " Sum(AJUSTE) as A, Sum(PLURIEMPLEO) as P From CODIGOLIQUIDACIONES" _
    & " Inner Join LIQUIDACIONGANANCIAS4TACATEGORIA" _
    & " On CODIGOLIQUIDACIONES.Codigo = LIQUIDACIONGANANCIAS4TACATEGORIA.CodigoLiquidacion" _
    & " Where PUESTOLABORAL = '" & PuestoLaboral _
    & "' And Right(PERIODO,4) = '" & Right(strPeriodo, 4) _
    & "' And CODIGO " & strSimboloSQL & " '" & CodigoLiquidacion & "'" 'Menor y/o IGUAL a la presente liquidación
    If SQLNoMatch(SQL) = True Then
        dblImporteCalculado = 0
    Else
        rstProcedimientoSlave.Open SQL, dbSlave, adOpenDynamic, adLockOptimistic
        dblImporteCalculado = rstProcedimientoSlave!HO + rstProcedimientoSlave!A + rstProcedimientoSlave!P
        rstProcedimientoSlave.Close
    End If
    
    CalcularRentaAcumulada = Round(dblImporteCalculado, 2)
    
    SQL = ""
    strSimboloSQL = ""
    dblImporteCalculado = 0
    strPeriodo = ""
    If rstProcedimientoSlave.State = adStateOpen Then
        rstProcedimientoSlave.Close
    End If
    Set rstProcedimientoSlave = Nothing
   
End Function

Public Function CalcularDescuentoAcumulado(PuestoLaboral As String, _
CodigoLiquidacion As String, Optional IncluirLiquidacionActual As Boolean = True) As Double

    Dim SQL As String
    Dim dblDescuentoCalculado As Double
    Dim strPeriodo As String
    Dim strSimboloSQL As String
    
    If IncluirLiquidacionActual = True Then
        strSimboloSQL = "<="
    Else
        strSimboloSQL = "<"
    End If
    
    
    'TENER CUIDADO con este PROCEDIMENTO!!! Descuento Acumulado INCLUYENDO la presente liquidación
    strPeriodo = BuscarPeriodoLiquidacion(CodigoLiquidacion)
    Set rstProcedimientoSlave = New ADODB.Recordset
    'Cargamos el Descuento Acumulada (No incluimos seguro de vida optativo y cuota sindical)
    SQL = "Select Sum(JUBILACION) as JUB, Sum(OBRASOCIAL) as OS," _
    & " Sum(ADHERENTEOBRASOCIAL) as AOS, Sum(SEGURODEVIDAOBLIGATORIO) as SVO," _
    & " Sum(CUOTASINDICAL) as CS, Sum(SEGURODEVIDAOPTATIVO) as SVOPT" _
    & " From CODIGOLIQUIDACIONES Inner Join LIQUIDACIONGANANCIAS4TACATEGORIA" _
    & " On CODIGOLIQUIDACIONES.Codigo = LIQUIDACIONGANANCIAS4TACATEGORIA.CodigoLiquidacion" _
    & " Where PUESTOLABORAL = '" & PuestoLaboral _
    & "' And Right(PERIODO,4) = '" & Right(strPeriodo, 4) _
    & "' And CODIGO " & strSimboloSQL & " '" & CodigoLiquidacion & "'" 'Menor e IGUAL a la presente liquidación
    If SQLNoMatch(SQL) = True Then
        dblDescuentoCalculado = 0
    Else
        rstProcedimientoSlave.Open SQL, dbSlave, adOpenDynamic, adLockOptimistic
        dblDescuentoCalculado = rstProcedimientoSlave!Jub + rstProcedimientoSlave!OS + rstProcedimientoSlave!AOS _
        + rstProcedimientoSlave!SVO '+ rstProcedimientoSlave!CS + rstProcedimientoSlave!SVOPT (Dejamos afuera el seguro de vida optativo y la Cuota SINDICAL)
        rstProcedimientoSlave.Close
    End If
    
    CalcularDescuentoAcumulado = Round(dblDescuentoCalculado, 2)
    
    SQL = ""
    strSimboloSQL = ""
    dblDescuentoCalculado = 0
    strPeriodo = ""
    If rstProcedimientoSlave.State = adStateOpen Then
        rstProcedimientoSlave.Close
    End If
    Set rstProcedimientoSlave = Nothing
   
End Function

Public Function CalcularSueldoNetoAcumulado(PuestoLaboral As String, _
CodigoLiquidacion As String, Optional IncluirLiquidacionActual As Boolean = True) As Double

    Dim dblSueldoNetoCalculado As Double
    
    'TENER CUIDADO con este PROCEDIMENTO!!! Sueldo Neto Acumulado INCLUYENDO la presente liquidación
    'Empezamos calculando la Renta Acumulada. _
    & TENER CUIDADO con este PROCEDIMENTO!!! Renta Acumulada INCLUYENDO la presente liquidación
    dblSueldoNetoCalculado = CalcularRentaAcumulada(PuestoLaboral, CodigoLiquidacion, IncluirLiquidacionActual)
    'Le restamos el descuento acumulado sin incluir Seguro De Vida Optativo. _
    & TENER CUIDADO con este PROCEDIMENTO!!! Descuento Acumulado INCLUYENDO la presente liquidación
    dblSueldoNetoCalculado = dblSueldoNetoCalculado - CalcularDescuentoAcumulado(PuestoLaboral, CodigoLiquidacion, IncluirLiquidacionActual)
       
    CalcularSueldoNetoAcumulado = Round(dblSueldoNetoCalculado, 2)
       
End Function

Public Function CalcularDeduccionMinimoNoImponible(PuestoLaboral As String, CodigoLiquidacion As String) As Double

    Dim SQL As String
    Dim dblImporteMensual As Double
    Dim dblImporteCalculado As Double
    Dim strPeriodo As String
    Dim strCodigoLiquidacion As String
    Dim datFecha As Date
    Dim datFechaControl As Date
    
    'Nos posicionamos en el último día del período a liquidar
    strPeriodo = BuscarPeriodoLiquidacion(CodigoLiquidacion)
    datFecha = BuscarUltimoDiaDelPeriodo(strPeriodo)
    'Controlamos lo que debió liquidarse en concepto de Mínimo no Imponible
    dblImporteMensual = ImporteDeduccionPersonal("MINIMONOIMPONIBLE", DateSerial(Year(datFecha) - 1, 12, 31))
    SQL = "Select MINIMONOIMPONIBLE From DEDUCCIONES4TACATEGORIA Where FECHA <= #" & Format(datFecha, "MM/DD/YYYY") & "# " _
    & "And Year(FECHA) = '" & Year(datFecha) & "' Order by FECHA Desc"
    If SQLNoMatch(SQL) = True Then
        dblImporteCalculado = dblImporteMensual * Month(datFecha)
    Else
        For i = 1 To Month(datFecha)
            datFechaControl = DateSerial(Year(datFecha), i, 1)
            datFechaControl = DateAdd("m", 1, datFechaControl)
            datFechaControl = DateAdd("d", -1, datFechaControl)
            dblImporteMensual = ImporteDeduccionPersonal("MINIMONOIMPONIBLE", datFechaControl)
            dblImporteCalculado = dblImporteCalculado + dblImporteMensual
        Next i
        datFechaControl = 0
    End If
    'Controlamos lo que realmente se liquidó en concepto de Mínimo no Imponible y _
    & cargamos el Importe de la diferencia entre lo que debió liquidarse y lo que se liquidó.
    strCodigoLiquidacion = BuscarCodigoLiquidacionAnterior(CodigoLiquidacion)
    dblImporteCalculado = dblImporteCalculado - ImporteRegistradoAcumuladoDeduccionEspecifica("MinimoNoImponible", PuestoLaboral, strPeriodo, strCodigoLiquidacion)

    CalcularDeduccionMinimoNoImponible = Round(dblImporteCalculado, 2)
    
    dblImporteMensual = 0
    dblImporteCalculado = 0
    datFecha = 0
    datFechaControl = 0
    strPeriodo = ""
    strCodigoLiquidacion = ""
   
End Function

Public Function CalcularDeduccionConyuge(PuestoLaboral As String, CodigoLiquidacion As String) As Double

    Dim SQL As String
    Dim dblImporteMensual As Double
    Dim dblImporteCalculado As Double
    Dim strPeriodo As String
    Dim strCodigoLiquidacion As String
    Dim datFecha As Date
    Dim datFechaControl As Date
        
    'Tener en cuenta que no está previsto el caso de ALTA/BAJA durante el año en curso
    If TieneConyugeDeducible(PuestoLaboral) = False Then
        dblImporteCalculado = 0
    Else
        'Nos posicionamos en el último día del período a liquidar
        strPeriodo = BuscarPeriodoLiquidacion(CodigoLiquidacion)
        datFecha = BuscarUltimoDiaDelPeriodo(strPeriodo)
        'Controlamos lo que debió liquidarse en concepto de Conyuge
        dblImporteMensual = ImporteDeduccionPersonal("CONYUGE", DateSerial(Year(datFecha) - 1, 12, 31))
        SQL = "Select CONYUGE From DEDUCCIONES4TACATEGORIA Where FECHA <= #" & Format(datFecha, "MM/DD/YYYY") & "# " _
        & "And Year(FECHA) = '" & Year(datFecha) & "' Order by FECHA Desc"
        If SQLNoMatch(SQL) = True Then
            dblImporteCalculado = dblImporteMensual * Month(datFecha)
        Else
            For i = 1 To Month(datFecha)
                datFechaControl = DateSerial(Year(datFecha), i, 1)
                datFechaControl = DateAdd("m", 1, datFechaControl)
                datFechaControl = DateAdd("d", -1, datFechaControl)
                dblImporteMensual = ImporteDeduccionPersonal("CONYUGE", datFechaControl)
                dblImporteCalculado = dblImporteCalculado + dblImporteMensual
            Next i
            datFechaControl = 0
        End If
        'Controlamos lo que realmente se liquidó en concepto de Conyuge y _
        & cargamos el Importe de la diferencia entre lo que debió liquidarse y lo que se liquidó.
        strCodigoLiquidacion = BuscarCodigoLiquidacionAnterior(CodigoLiquidacion)
        dblImporteCalculado = dblImporteCalculado - ImporteRegistradoAcumuladoDeduccionEspecifica("Conyuge", PuestoLaboral, strPeriodo, strCodigoLiquidacion)
    End If
    
    CalcularDeduccionConyuge = Round(dblImporteCalculado, 2)
    
    dblImporteMensual = 0
    dblImporteCalculado = 0
    datFecha = 0
    datFechaControl = 0
    strPeriodo = ""
    strCodigoLiquidacion = ""
   
End Function

Public Function CalcularDeduccionHijo(PuestoLaboral As String, CodigoLiquidacion As String) As Double

    Dim SQL As String
    Dim intContar As String
    Dim dblImporteMensual As Double
    Dim dblImporteCalculado As Double
    Dim strPeriodo As String
    Dim strCodigoLiquidacion As String
    Dim datFecha As Date
    Dim datFechaControl As Date
        
    'Tener en cuenta que no está previsto el caso de ALTA/BAJA durante el año en curso
    'Nos posicionamos en el último día del período a liquidar
    strPeriodo = BuscarPeriodoLiquidacion(CodigoLiquidacion)
    datFecha = BuscarUltimoDiaDelPeriodo(strPeriodo)
    intContar = CantidadHijosDeducibles(PuestoLaboral, datFecha)
    If intContar = 0 Then
        dblImporteCalculado = 0
    Else
        'Controlamos lo que debió liquidarse en concepto de Hijo
        dblImporteMensual = ImporteDeduccionPersonal("HIJO", DateSerial(Year(datFecha) - 1, 12, 31))
        SQL = "Select HIJO From DEDUCCIONES4TACATEGORIA Where FECHA <= #" & Format(datFecha, "MM/DD/YYYY") & "# " _
        & "And Year(FECHA) = '" & Year(datFecha) & "' Order by FECHA Desc"
        If SQLNoMatch(SQL) = True Then
            For i = 1 To Month(datFecha)
                intContar = CantidadHijosDeducibles(PuestoLaboral, DateSerial(Year(datFecha), i, 1))
                dblImporteCalculado = dblImporteCalculado + (dblImporteMensual * intContar)
            Next i
        Else
            For i = 1 To Month(datFecha)
                datFechaControl = DateSerial(Year(datFecha), i, 1)
                datFechaControl = DateAdd("m", 1, datFechaControl)
                datFechaControl = DateAdd("d", -1, datFechaControl)
                dblImporteMensual = ImporteDeduccionPersonal("HIJO", datFechaControl)
                intContar = CantidadHijosDeducibles(PuestoLaboral, DateSerial(Year(datFecha), i, 1))
                dblImporteCalculado = dblImporteCalculado + (dblImporteMensual * intContar)
            Next i
        End If
        'Controlamos lo que realmente se liquidó en concepto de Hijo y _
        & cargamos el Importe de la diferencia entre lo que debió liquidarse y lo que se liquidó.
        strCodigoLiquidacion = BuscarCodigoLiquidacionAnterior(CodigoLiquidacion)
        dblImporteCalculado = dblImporteCalculado - ImporteRegistradoAcumuladoDeduccionEspecifica("Hijo", PuestoLaboral, strPeriodo, strCodigoLiquidacion)
    End If
    
    CalcularDeduccionHijo = Round(dblImporteCalculado, 2)
    
    dblImporteMensual = 0
    dblImporteCalculado = 0
    datFecha = 0
    datFechaControl = 0
    strPeriodo = ""
    strCodigoLiquidacion = ""
   
End Function

Public Function CalcularDeduccionOtrasCargasDeFamilia(PuestoLaboral As String, CodigoLiquidacion As String) As Double

    Dim SQL As String
    Dim intContar As String
    Dim dblImporteMensual As Double
    Dim dblImporteCalculado As Double
    Dim strPeriodo As String
    Dim strCodigoLiquidacion As String
    Dim datFecha As Date
    Dim datFechaControl As Date
        
    'Tener en cuenta que no está previsto el caso de ALTA/BAJA durante el año en curso
    'Nos posicionamos en el último día del período a liquidar
    strPeriodo = BuscarPeriodoLiquidacion(CodigoLiquidacion)
    datFecha = BuscarUltimoDiaDelPeriodo(strPeriodo)
    intContar = CantidadOtrasCargasFamiliaDeducibles(PuestoLaboral)
    If intContar = 0 Then
        dblImporteCalculado = 0
    Else
        'Controlamos lo que debió liquidarse en concepto de Otras Cargas de Familia
        dblImporteMensual = ImporteDeduccionPersonal("OTRASCARGASDEFAMILIA", DateSerial(Year(datFecha) - 1, 12, 31))
        SQL = "Select OTRASCARGASDEFAMILIA From DEDUCCIONES4TACATEGORIA Where FECHA <= #" & Format(datFecha, "MM/DD/YYYY") & "# " _
        & "And Year(FECHA) = '" & Year(datFecha) & "' Order by FECHA Desc"
        If SQLNoMatch(SQL) = True Then
            dblImporteCalculado = dblImporteMensual * Month(datFecha) * intContar
        Else
            For i = 1 To Month(datFecha)
                datFechaControl = DateSerial(Year(datFecha), i, 1)
                datFechaControl = DateAdd("m", 1, datFechaControl)
                datFechaControl = DateAdd("d", -1, datFechaControl)
                dblImporteMensual = ImporteDeduccionPersonal("OTRASCARGASDEFAMILIA", datFechaControl)
                dblImporteCalculado = dblImporteCalculado + dblImporteMensual
            Next i
            dblImporteCalculado = dblImporteCalculado * intContar
        End If
        'Controlamos lo que realmente se liquidó en concepto de Otras Cargas de Familia y _
        & cargamos el Importe de la diferencia entre lo que debió liquidarse y lo que se liquidó.
        strCodigoLiquidacion = BuscarCodigoLiquidacionAnterior(CodigoLiquidacion)
        dblImporteCalculado = dblImporteCalculado - ImporteRegistradoAcumuladoDeduccionEspecifica("OtrasCargasDeFamilia", PuestoLaboral, strPeriodo, strCodigoLiquidacion)
    End If
    
    CalcularDeduccionOtrasCargasDeFamilia = Round(dblImporteCalculado, 2)
    
    dblImporteMensual = 0
    dblImporteCalculado = 0
    datFecha = 0
    datFechaControl = 0
    strPeriodo = ""
    strCodigoLiquidacion = ""
   
End Function

Public Function CalcularDeduccionEspecial(PuestoLaboral As String, _
CodigoLiquidacion As String, Optional IncluirLiquidacionActual As Boolean = True, _
Optional GciaNetaAntesDeDeduccionEspecial As Variant = "NaN") As Double

    Dim SQL As String
    Dim dblImporteMensual As Double
    Dim dblIncrementoDeduccionEspecial As Double
    Dim dblImporteCalculado As Double
    Dim dblGciaNetaAntesDeDeduccionEspecial As Double
    Dim strPeriodo As String
    Dim strCodigoLiquidacion As String
    Dim datFecha As Date
    Dim datFechaControl As Date
    
    'Nos posicionamos en el último día del período a liquidar
    strPeriodo = BuscarPeriodoLiquidacion(CodigoLiquidacion)
    datFecha = BuscarUltimoDiaDelPeriodo(strPeriodo)
    'Controlamos lo que debió liquidarse en concepto de Deducción Especial
    dblImporteMensual = ImporteDeduccionPersonal("DEDUCCIONESPECIAL", DateSerial(Year(datFecha) - 1, 12, 31))
    SQL = "Select DEDUCCIONESPECIAL From DEDUCCIONES4TACATEGORIA Where FECHA <= #" & Format(datFecha, "MM/DD/YYYY") & "# " _
    & "And Year(FECHA) = '" & Year(datFecha) & "' Order by FECHA Desc"
    If SQLNoMatch(SQL) = True Then
        dblImporteCalculado = dblImporteMensual * Month(datFecha)
    Else
        For i = 1 To Month(datFecha)
            datFechaControl = DateSerial(Year(datFecha), i, 1)
            datFechaControl = DateAdd("m", 1, datFechaControl)
            datFechaControl = DateAdd("d", -1, datFechaControl)
            dblImporteMensual = ImporteDeduccionPersonal("DEDUCCIONESPECIAL", datFechaControl)
            dblImporteCalculado = dblImporteCalculado + dblImporteMensual
        Next i
        datFechaControl = 0
    End If
    
    'Ver ViejoProcedimientoIncrementoDeduccionEspecial
    
    'Procedimiento General cuando una liquidación está exenta (Ejemplo: Exenciones en SAC)
    dblImporteMensual = 0
    SQL = "Select * From CODIGOLIQUIDACIONES" _
    & " Where MONTOEXENTO > 0" _
    & " And Right(PERIODO,4) = '" & Year(datFecha) _
    & "' And Codigo <= '" & CodigoLiquidacion & "'"
    If SQLNoMatch(SQL) = False Then
        Set rstListadoSlave = New ADODB.Recordset
        With rstListadoSlave
            .Open SQL, dbSlave, adOpenDynamic, adLockOptimistic
            .MoveFirst
            While rstListadoSlave.EOF = False
                dblImporteMensual = dblImporteMensual _
                + CalcularHaberBruto(PuestoLaboral, rstListadoSlave!Codigo)
                dblImporteMensual = dblImporteMensual _
                - CalcularJubilacion(PuestoLaboral, rstListadoSlave!Codigo)
                dblImporteMensual = dblImporteMensual _
                - CalcularObraSocial(PuestoLaboral, rstListadoSlave!Codigo)
                dblImporteMensual = dblImporteMensual _
                - CalcularAdherenteObraSocial(PuestoLaboral, rstListadoSlave!Codigo)
                dblImporteMensual = dblImporteMensual _
                - CalcularDescuentoSeguroDeVidaObligatorio(PuestoLaboral, rstListadoSlave!Codigo)
                dblImporteMensual = dblImporteMensual _
                - CalcularDescuentoCuotaSindical(PuestoLaboral, rstListadoSlave!Codigo)
                'Que hacer con Deducciones Personales y Generales
                If rstListadoSlave!MontoExento < dblImporteMensual Then
                    dblIncrementoDeduccionEspecial = dblIncrementoDeduccionEspecial _
                    + rstListadoSlave!MontoExento
                Else
                    dblIncrementoDeduccionEspecial = dblIncrementoDeduccionEspecial _
                    + dblImporteMensual
                End If
                .MoveNext
            Wend
            .Close
        End With
        dblImporteCalculado = dblImporteCalculado + dblIncrementoDeduccionEspecial
        Set rstListadoSlave = Nothing
    End If
    'Limite Deduccion Especial
    If IsNumeric(GciaNetaAntesDeDeduccionEspecial) Then
        dblGciaNetaAntesDeDeduccionEspecial = GciaNetaAntesDeDeduccionEspecial
    Else 'Revisar para el caso que no traiga la Gcia. Neta Antes de Deducción Especial
        dblGciaNetaAntesDeDeduccionEspecial = CalcularSueldoNetoAcumulado(PuestoLaboral, CodigoLiquidacion, IncluirLiquidacionActual) _
        - CalcularDeduccionesGeneralesAcumuladas(PuestoLaboral, CodigoLiquidacion, IncluirLiquidacionActual, False) _
        - CalcularDeduccionesPersonalesAcumuladas(PuestoLaboral, CodigoLiquidacion, IncluirLiquidacionActual, True)
    End If
    If dblGciaNetaAntesDeDeduccionEspecial < dblImporteCalculado Then
        dblImporteCalculado = dblGciaNetaAntesDeDeduccionEspecial
    End If
    'Controlamos lo que realmente se liquidó en concepto de Deducción Especial y _
    & cargamos el Importe de la diferencia entre lo que debió liquidarse y lo que se liquidó.
    strCodigoLiquidacion = BuscarCodigoLiquidacionAnterior(CodigoLiquidacion)
    dblImporteCalculado = dblImporteCalculado - ImporteRegistradoAcumuladoDeduccionEspecifica("DeduccionEspecial", PuestoLaboral, strPeriodo, strCodigoLiquidacion)

    CalcularDeduccionEspecial = Round(dblImporteCalculado, 2)
    
    SQL = ""
    dblImporteMensual = 0
    dblImporteCalculado = 0
    dblIncrementoDeduccionEspecial = 0
    dblGciaNetaAntesDeDeduccionEspecial = 0
    datFecha = 0
    datFechaControl = 0
    strPeriodo = ""
    strCodigoLiquidacion = ""
    Set rstProcedimientoSlave = Nothing
   
End Function

Public Function CalcularDeduccionServicioDomestico(PuestoLaboral As String, CodigoLiquidacion As String) As Double

    Dim SQL As String
    Dim dblImporteMensual As Double
    Dim dblImporteCalculado As Double
    Dim strPeriodo As String
    Dim strCodigoLiquidacion As String
    Dim datFecha As Date
    
    'Nos posicionamos en el último día del período a liquidar
    strPeriodo = BuscarPeriodoLiquidacion(CodigoLiquidacion)
    datFecha = BuscarUltimoDiaDelPeriodo(strPeriodo)
    'Controlamos lo que debió liquidarse en concepto de Servicio Doméstico
    dblImporteMensual = ImporteDeduccionGeneral(PuestoLaboral, "SERVICIODOMESTICO", datFecha)
    dblImporteCalculado = dblImporteMensual * Month(datFecha)
    'Controlamos lo que realmente se liquidó en concepto de Servicio Doméstico y _
    & cargamos el Importe de la diferencia entre lo que debió liquidarse y lo que se liquidó.
    strCodigoLiquidacion = BuscarCodigoLiquidacionAnterior(CodigoLiquidacion)
    dblImporteCalculado = dblImporteCalculado - ImporteRegistradoAcumuladoDeduccionEspecifica("ServicioDomestico", PuestoLaboral, strPeriodo, strCodigoLiquidacion)

    CalcularDeduccionServicioDomestico = Round(dblImporteCalculado, 2)
    
    dblImporteMensual = 0
    dblImporteCalculado = 0
    datFecha = 0
    strPeriodo = ""
    strCodigoLiquidacion = ""
   
End Function

Public Function CalcularDeduccionSeguroDeVida(PuestoLaboral As String, _
CodigoLiquidacion As String, Optional SeguroOptativoRecibo As Variant = "NaN") As Double

    Dim SQL As String
    Dim dblImporteMensual As Double
    Dim dblImporteCalculado As Double
    Dim strPeriodo As String
    Dim strCodigoLiquidacion As String
    Dim datFecha As Date
    
    'Buscamos el monto de Seguro de Vida Optativo del recibo de Sueldo INVICO
    If IsNumeric(SeguroOptativoRecibo) Then
        dblImporteCalculado = SeguroOptativoRecibo
    Else
        dblImporteCalculado = CalcularDescuentoSeguroDeVidaOptativo(PuestoLaboral, CodigoLiquidacion)
    End If
    'Nos posicionamos en el último día del período a liquidar
    strPeriodo = BuscarPeriodoLiquidacion(CodigoLiquidacion)
    datFecha = BuscarUltimoDiaDelPeriodo(strPeriodo)
    'Controlamos lo que debió liquidarse en concepto de Seguro de Vida
    dblImporteMensual = ImporteDeduccionGeneral(PuestoLaboral, "SEGURODEVIDA", datFecha, , dblImporteCalculado)
    dblImporteCalculado = dblImporteMensual * Month(datFecha)
    'Controlamos lo que realmente se liquidó en concepto de Seguro de Vida (Verificar)
    If dblImporteCalculado > 0 Then
        strCodigoLiquidacion = BuscarCodigoLiquidacionAnterior(CodigoLiquidacion)
        dblImporteCalculado = dblImporteCalculado - ImporteRegistradoAcumuladoDeduccionEspecifica("SeguroDeVidaOptativo", PuestoLaboral, strPeriodo, strCodigoLiquidacion)
    Else
        'dblImporteCalculado = 0
    End If
    CalcularDeduccionSeguroDeVida = Round(dblImporteCalculado, 2)
    
    dblImporteMensual = 0
    dblImporteCalculado = 0
    datFecha = 0
    strPeriodo = ""
    strCodigoLiquidacion = ""
   
End Function


Public Function CalcularDeduccionAlquiler(PuestoLaboral As String, CodigoLiquidacion As String) As Double

    Dim SQL As String
    Dim dblImporteMensual As Double
    Dim dblImporteCalculado As Double
    Dim strPeriodo As String
    Dim strCodigoLiquidacion As String
    Dim datFecha As Date
    
    'Nos posicionamos en el último día del período a liquidar
    strPeriodo = BuscarPeriodoLiquidacion(CodigoLiquidacion)
    datFecha = BuscarUltimoDiaDelPeriodo(strPeriodo)
    'Controlamos lo que debió liquidarse en concepto de Servicio Doméstico
    dblImporteMensual = ImporteDeduccionGeneral(PuestoLaboral, "ALQUILERES", datFecha)
    dblImporteCalculado = dblImporteMensual * Month(datFecha)
    'Controlamos lo que realmente se liquidó en concepto de Servicio Doméstico y _
    & cargamos el Importe de la diferencia entre lo que debió liquidarse y lo que se liquidó.
    strCodigoLiquidacion = BuscarCodigoLiquidacionAnterior(CodigoLiquidacion)
    dblImporteCalculado = dblImporteCalculado - ImporteRegistradoAcumuladoDeduccionEspecifica("Alquileres", PuestoLaboral, strPeriodo, strCodigoLiquidacion)

    CalcularDeduccionAlquiler = Round(dblImporteCalculado, 2)
    
    dblImporteMensual = 0
    dblImporteCalculado = 0
    datFecha = 0
    strPeriodo = ""
    strCodigoLiquidacion = ""
   
End Function


Public Function CalcularDeduccionCuotaMedicoAsistencial(PuestoLaboral As String, _
CodigoLiquidacion As String, Optional IncluirLiquidacionActual As Boolean = True, _
Optional RdoNetaAntesDeCMAyDyHM As Variant = "NaN") As Double

    Dim SQL As String
    Dim dblImporteCalculado As Double
    Dim strPeriodo As String
    Dim strCodigoLiquidacion As String
    Dim datFecha As Date
    
    'Buscamos el monto de la Ganancia Neta Impositiva
    If IsNumeric(RdoNetaAntesDeCMAyDyHM) Then
        dblImporteCalculado = RdoNetaAntesDeCMAyDyHM
    Else
        dblImporteCalculado = CalcularSueldoNetoAcumulado(PuestoLaboral, CodigoLiquidacion, IncluirLiquidacionActual) _
        - CalcularDeduccionesGeneralesAcumuladas(PuestoLaboral, CodigoLiquidacion, IncluirLiquidacionActual, True)
    End If
    'Nos posicionamos en el último día del período a liquidar
    strPeriodo = BuscarPeriodoLiquidacion(CodigoLiquidacion)
    datFecha = BuscarUltimoDiaDelPeriodo(strPeriodo)
    'Controlamos lo que debió liquidarse en concepto de Cuota Médico Asistencial
    dblImporteCalculado = ImporteDeduccionGeneral(PuestoLaboral, "CUOTAMEDICOASISTENCIAL", datFecha, dblImporteCalculado)
    'Controlamos lo que realmente se liquidó en concepto de Servicio Doméstico y _
    & cargamos el Importe de la diferencia entre lo que debió liquidarse y lo que se liquidó.
    strCodigoLiquidacion = BuscarCodigoLiquidacionAnterior(CodigoLiquidacion)
    dblImporteCalculado = dblImporteCalculado - ImporteRegistradoAcumuladoDeduccionEspecifica("CuotaMedicoAsistencial", PuestoLaboral, strPeriodo, strCodigoLiquidacion)

    CalcularDeduccionCuotaMedicoAsistencial = Round(dblImporteCalculado, 2)
    
    dblImporteMensual = 0
    dblImporteCalculado = 0
    datFecha = 0
    strPeriodo = ""
    strCodigoLiquidacion = ""
   
End Function

Public Function CalcularDeduccionCuotaSindical(PuestoLaboral As String, CodigoLiquidacion As String, _
Optional CuotaSindicalRecibo As Variant = "NaN") As Double

    'No esta previsto esta deduccion por fuera del Recibo de Haberes INVICO
    If IsNumeric(CuotaSindicalRecibo) Then
        CalcularDeduccionCuotaSindical = CuotaSindicalRecibo
    Else
        CalcularDeduccionCuotaSindical = CalcularDescuentoCuotaSindical(PuestoLaboral, CodigoLiquidacion)
    End If

End Function

Public Function CalcularDeduccionDonaciones(PuestoLaboral As String, _
CodigoLiquidacion As String, Optional IncluirLiquidacionActual As Boolean = True, _
Optional RdoNetaAntesDeCMAyDyHM As Variant = "NaN") As Double

    Dim SQL As String
    Dim dblImporteCalculado As Double
    Dim strPeriodo As String
    Dim strCodigoLiquidacion As String
    Dim datFecha As Date
    
    'Buscamos el monto de la Ganancia Neta Impositiva
    If IsNumeric(RdoNetaAntesDeCMAyDyHM) Then
        dblImporteCalculado = RdoNetaAntesDeCMAyDyHM
    Else
        dblImporteCalculado = CalcularSueldoNetoAcumulado(PuestoLaboral, CodigoLiquidacion, IncluirLiquidacionActual) _
        - CalcularDeduccionesGeneralesAcumuladas(PuestoLaboral, CodigoLiquidacion, IncluirLiquidacionActual, True)
    End If
    'dblImporteCalculado = CalcularDescuentoSeguroDeVidaOptativo(PuestoLaboral, CodigoLiquidacion)
    'Nos posicionamos en el último día del período a liquidar
    strPeriodo = BuscarPeriodoLiquidacion(CodigoLiquidacion)
    datFecha = BuscarUltimoDiaDelPeriodo(strPeriodo)
    'Controlamos lo que debió liquidarse en concepto de Cuota Médico Asistencial
    dblImporteCalculado = ImporteDeduccionGeneral(PuestoLaboral, "DONACIONES", datFecha, dblImporteCalculado)
    'Controlamos lo que realmente se liquidó en concepto de Servicio Doméstico y _
    & cargamos el Importe de la diferencia entre lo que debió liquidarse y lo que se liquidó.
    strCodigoLiquidacion = BuscarCodigoLiquidacionAnterior(CodigoLiquidacion)
    dblImporteCalculado = dblImporteCalculado - ImporteRegistradoAcumuladoDeduccionEspecifica("Donaciones", PuestoLaboral, strPeriodo, strCodigoLiquidacion)

    CalcularDeduccionDonaciones = Round(dblImporteCalculado, 2)
    
    dblImporteMensual = 0
    dblImporteCalculado = 0
    datFecha = 0
    strPeriodo = ""
    strCodigoLiquidacion = ""
   
End Function

Public Function CalcularDeduccionHonorariosMedicos(PuestoLaboral As String, _
CodigoLiquidacion As String, Optional IncluirLiquidacionActual As Boolean = True, _
Optional RdoNetaAntesDeCMAyDyHM As Variant = "NaN") As Double

    Dim SQL As String
    Dim dblImporteCalculado As Double
    Dim strPeriodo As String
    Dim strCodigoLiquidacion As String
    Dim datFecha As Date
    
    'Buscamos el monto de la Ganancia Neta Impositiva
    If IsNumeric(RdoNetaAntesDeCMAyDyHM) Then
        dblImporteCalculado = RdoNetaAntesDeCMAyDyHM
    Else
        dblImporteCalculado = CalcularSueldoNetoAcumulado(PuestoLaboral, CodigoLiquidacion, IncluirLiquidacionActual) _
        - CalcularDeduccionesGeneralesAcumuladas(PuestoLaboral, CodigoLiquidacion, IncluirLiquidacionActual, True)
    End If
    'Nos posicionamos en el último día del período a liquidar
    strPeriodo = BuscarPeriodoLiquidacion(CodigoLiquidacion)
    datFecha = BuscarUltimoDiaDelPeriodo(strPeriodo)
    'Controlamos lo que debió liquidarse en concepto de Cuota Médico Asistencial
    dblImporteCalculado = ImporteDeduccionGeneral(PuestoLaboral, "HONORARIOSMEDICOS", datFecha, dblImporteCalculado)
    'Controlamos lo que realmente se liquidó en concepto de Servicio Doméstico y _
    & cargamos el Importe de la diferencia entre lo que debió liquidarse y lo que se liquidó.
    strCodigoLiquidacion = BuscarCodigoLiquidacionAnterior(CodigoLiquidacion)
    dblImporteCalculado = dblImporteCalculado - ImporteRegistradoAcumuladoDeduccionEspecifica("HonorariosMedicos", PuestoLaboral, strPeriodo, strCodigoLiquidacion)

    CalcularDeduccionHonorariosMedicos = Round(dblImporteCalculado, 2)
    
    dblImporteMensual = 0
    dblImporteCalculado = 0
    datFecha = 0
    strPeriodo = ""
    strCodigoLiquidacion = ""
   
End Function

Public Function BuscarPeriodoLiquidacion(CodigoLiquidacion As String) As String

    Dim SQL As String
    
    SQL = "Select * From CODIGOLIQUIDACIONES Where CODIGO = '" & CodigoLiquidacion & "'"
    Set rstBuscarSlave = New ADODB.Recordset
    rstBuscarSlave.Open SQL, dbSlave, adOpenDynamic, adLockReadOnly
    BuscarPeriodoLiquidacion = rstBuscarSlave!Periodo
    rstBuscarSlave.Close
    Set rstBuscarSlave = Nothing

End Function

Public Function BuscarUltimoDiaDelPeriodo(MesBarraAño As String) As Date

    Dim datFecha As Date

    'Nos posicionamos en el último día del período a liquidar
    datFecha = DateTime.DateSerial(Right(MesBarraAño, 4), Left(MesBarraAño, 2), 1)
    datFecha = DateAdd("m", 1, datFecha)
    datFecha = DateAdd("d", -1, datFecha)
    
    BuscarUltimoDiaDelPeriodo = datFecha

End Function

Public Function BuscarDenominacionDeduccionSIRADIG(CodigoSIRADIG As String) As String

    Dim SQL As String
    
    SQL = "Select Denominacion From DeduccionesSIRADIG Where CODIGO = '" & CodigoSIRADIG & "'"
    Set rstBuscarSlave = New ADODB.Recordset
    rstBuscarSlave.Open SQL, dbSlave, adOpenDynamic, adLockReadOnly
    BuscarDenominacionDeduccionSIRADIG = rstBuscarSlave!Denominacion
    rstBuscarSlave.Close
    Set rstBuscarSlave = Nothing

End Function

Public Function EquipararConceptoDeduccionSISPERconCodigoSIRADIG(DeduccionSLAVE As String) As String

    Select Case DeduccionSLAVE
    Case "Jubilacion"
        EquipararConceptoDeduccionSISPERconCodigoSIRADIG = "01"
    Case "ObraSocial"
        EquipararConceptoDeduccionSISPERconCodigoSIRADIG = "02"
    Case "SeguroDeVidaObligatorio"
        EquipararConceptoDeduccionSISPERconCodigoSIRADIG = "03"
    Case "SeguroDeVidaOptativo"
        EquipararConceptoDeduccionSISPERconCodigoSIRADIG = "2"
    Case "CuotaSindical"
        EquipararConceptoDeduccionSISPERconCodigoSIRADIG = "04"
    Case "AdherenteObraSocial"
        EquipararConceptoDeduccionSISPERconCodigoSIRADIG = "05"
    Case "ServicioDomestico"
        EquipararConceptoDeduccionSISPERconCodigoSIRADIG = "8"
    Case "Alquileres"
        EquipararConceptoDeduccionSISPERconCodigoSIRADIG = "22"
    Case "CuotaMedicoAsistencial"
        EquipararConceptoDeduccionSISPERconCodigoSIRADIG = "1"
    Case "Donaciones"
        EquipararConceptoDeduccionSISPERconCodigoSIRADIG = "3"
    Case "HonorariosMedicos"
        EquipararConceptoDeduccionSISPERconCodigoSIRADIG = "7"
    Case Else
        EquipararConceptoDeduccionSISPERconCodigoSIRADIG = "00"
    End Select

End Function

Public Function EquipararCodigoDeduccionSIRADIGconConceptoSISPER(CodigoSIRADIG As String) As String

    Select Case CodigoSIRADIG
    Case "01"
        EquipararCodigoDeduccionSIRADIGconConceptoSISPER = "Jubilacion"
    Case "02"
        EquipararCodigoDeduccionSIRADIGconConceptoSISPER = "ObraSocial"
    Case "03"
        EquipararCodigoDeduccionSIRADIGconConceptoSISPER = "SeguroDeVidaObligatorio"
    Case "2"
        EquipararCodigoDeduccionSIRADIGconConceptoSISPER = "SeguroDeVidaOptativo"
    Case "04"
        EquipararCodigoDeduccionSIRADIGconConceptoSISPER = "CuotaSindical"
    Case "05"
        EquipararCodigoDeduccionSIRADIGconConceptoSISPER = "AdherenteObraSocial"
    Case "8"
        EquipararCodigoDeduccionSIRADIGconConceptoSISPER = "ServicioDomestico"
    Case "22"
        EquipararCodigoDeduccionSIRADIGconConceptoSISPER = "Alquileres"
    Case "1"
        EquipararCodigoDeduccionSIRADIGconConceptoSISPER = "CuotaMedicoAsistencial"
    Case "3"
        EquipararCodigoDeduccionSIRADIGconConceptoSISPER = "Donaciones"
    Case "7"
        EquipararCodigoDeduccionSIRADIGconConceptoSISPER = "HonorariosMedicos"
    Case Else
        EquipararCodigoDeduccionSIRADIGconConceptoSISPER = "NoRegistrado"
    End Select

End Function

Public Function BuscarDenominacionDeduccionPersonalSLAVE(CodigoDataGridSLAVE As String) As String

    Select Case CodigoDataGridSLAVE
    Case "MNI"
        BuscarDenominacionDeduccionPersonalSLAVE = "MinimoNoImponible"
    Case "C"
        BuscarDenominacionDeduccionPersonalSLAVE = "Conyuge"
    Case "H"
        BuscarDenominacionDeduccionPersonalSLAVE = "Hijo"
    Case "OCF"
        BuscarDenominacionDeduccionPersonalSLAVE = "OtrasCargasDeFamilia"
    Case "DE"
        BuscarDenominacionDeduccionPersonalSLAVE = "DeduccionEspecial"
    Case Else
        BuscarDenominacionDeduccionPersonalSLAVE = "00"
    End Select

End Function


Public Function EquipararCodigoParentescoSISPERconSIRADIG(CodigoParentescoSISPER As String, _
Optional EsDiscapacitado As Boolean = False) As String

    Select Case CodigoParentescoSISPER
    Case "1", "3"
        If EsDiscapacitado = False Then
            EquipararCodigoParentescoSISPERconSIRADIG = "3"
        Else
            EquipararCodigoParentescoSISPERconSIRADIG = "31"
        End If
    Case "2", "50"
        EquipararCodigoParentescoSISPERconSIRADIG = "1"
    Case "4"
        EquipararCodigoParentescoSISPERconSIRADIG = "35"
    Case Else
        EquipararCodigoParentescoSISPERconSIRADIG = "00"
    End Select

End Function


Public Function ImporteRegistradoDeduccionEspecifica(Deduccion As String, PuestoLaboral As String, _
CodigoLiquidacion As String) As Double

    Dim SQL As String
    Dim dblImporteCalculado As Double
    
    SQL = "Select " & Deduccion & " AS Deduccion From LIQUIDACIONGANANCIAS4TACATEGORIA" _
    & " Where PUESTOLABORAL = '" & PuestoLaboral _
    & "' And CODIGOLIQUIDACION = '" & CodigoLiquidacion & "'"
    If SQLNoMatch(SQL) = True Then
        dblImporteCalculado = 0
    Else
        Set rstBuscarSlave = New ADODB.Recordset
        rstBuscarSlave.Open SQL, dbSlave, adOpenDynamic, adLockOptimistic
        dblImporteCalculado = rstBuscarSlave!Deduccion
        rstBuscarSlave.Close
        Set rstBuscarSlave = Nothing
    End If
    
    ImporteRegistradoDeduccionEspecifica = Round(dblImporteCalculado, 2)
    
    SQL = ""
    dblImporteCalculado = 0
    
End Function

Public Function CalcularDeduccionGeneralEspecificaSIRADIG(ID As String, CodigoDeduccion As String, _
PuestoLaboral As String, CodigoLiquidacion As String, Optional GciaNeta As Double = 0) As Double

    Dim SQL As String
    Dim strPeriodo As String
    Dim strDenominacionSISPER As String
    Dim dblImporteInformado As Double
    Dim dblImporteLimite As Double
    Dim dblImporteAcumulado As Double
    Dim dblImporteCalculado As Double
    
    dblImporteInformado = 0
    dblImporteLimite = 0
    dblImporteAcumulado = 0
    dblImporteCalculado = 0
    
    strPeriodo = BuscarPeriodoLiquidacion(CodigoLiquidacion)
    '1) Debemos calcular el importe declarado por SIRADIG acumulado a la fecha de liquidacion
    If ID = "Ninguno" Then
        dblImporteInformado = 0
    Else
        SQL = "Select "
        If Val(Left(strPeriodo, 2)) > 1 Then
            For i = 1 To (Val(Left(strPeriodo, 2)) - 1)
                SQL = SQL & "Mes" & Format(i, "00") & ", "
            Next i
        End If
        SQL = SQL & "Mes" & Left(strPeriodo, 2) & " From DeduccionesGeneralesSIRADIG " _
            & "Where ID = '" & ID & "' " _
            & "And CodigoDeduccion = '" & CodigoDeduccion & "'"
        If SQLNoMatch(SQL) = True Then
            dblImporteInformado = 0
        Else
            Set rstBuscarSlave = New ADODB.Recordset
            rstBuscarSlave.Open SQL, dbSlave, adOpenDynamic, adLockOptimistic
            For i = 1 To rstBuscarSlave.Fields.Count
                dblImporteInformado = dblImporteInformado + rstBuscarSlave.Fields(i - 1)
            Next i
            rstBuscarSlave.Close
            Set rstBuscarSlave = Nothing
        End If
    End If
    '2) Debemos buscar el límite legal para la deducción y comparar con 1) (elegir el menor)
    dblImporteLimite = ImporteLimiteDeduccionGeneralSIRADIG(CodigoDeduccion, strPeriodo, GciaNeta)
    If dblImporteLimite > dblImporteInformado Then
        dblImporteLimite = dblImporteInformado
    End If
    
    '3) debemos determinar el importe ya registrado hasta la liquidación previa y comparar con 2) (sacar diferencia)
    strDenominacionSISPER = EquipararCodigoDeduccionSIRADIGconConceptoSISPER(CodigoDeduccion)
    dblImporteAcumulado = ImporteRegistradoAcumuladoDeduccionEspecifica(strDenominacionSISPER, _
    PuestoLaboral, strPeriodo, CodigoLiquidacion, False)
    dblImporteCalculado = dblImporteLimite - dblImporteAcumulado
    
    CalcularDeduccionGeneralEspecificaSIRADIG = Round(dblImporteCalculado, 2)
    
    SQL = ""
    dblImporteCalculado = 0
    
End Function

Public Function CalcularDeduccionPersonalEspecificaSIRADIG(ID As String, DenominacionSLAVE As String, _
PuestoLaboral As String, CodigoLiquidacion As String, Optional GciaNeta As Double = 0) As Double

    Dim SQL As String
    Dim strPeriodo As String
    Dim datFecha As Date
    Dim datFechaControl As Date
    Dim inPrincipio As Integer
    Dim inFin As Integer
    Dim strDenominacionSISPER As String
    Dim dblImporteMensual As Double
    Dim dblImporteLimite As Double
    Dim dblImporteAcumulado As Double
    Dim dblImporteCalculado As Double
    
    dblImporteMensual = 0
    dblImporteLimite = 0
    dblImporteAcumulado = 0
    dblImporteCalculado = 0
    
    strPeriodo = BuscarPeriodoLiquidacion(CodigoLiquidacion)
    datFecha = BuscarUltimoDiaDelPeriodo(strPeriodo)
    
    Select Case DenominacionSLAVE
    Case "MinimoNoImponible", "DeduccionEspecial"
        '1) Debemos buscar el límite legal para la deducción
        dblImporteMensual = ImporteDeduccionPersonal(DenominacionSLAVE, DateSerial(Year(datFecha) - 1, 12, 31))
        SQL = "Select " & DenominacionSLAVE & " From DEDUCCIONES4TACATEGORIA Where FECHA <= #" & Format(datFecha, "MM/DD/YYYY") & "# " _
        & "And Year(FECHA) = '" & Year(datFecha) & "' Order by FECHA Desc"
        If SQLNoMatch(SQL) = True Then
            dblImporteLimite = dblImporteMensual * Month(datFecha)
        Else
            For i = 1 To Month(datFecha)
                datFechaControl = DateSerial(Year(datFecha), i, 1)
                datFechaControl = DateAdd("m", 1, datFechaControl)
                datFechaControl = DateAdd("d", -1, datFechaControl)
                dblImporteMensual = ImporteDeduccionPersonal(DenominacionSLAVE, datFechaControl)
                dblImporteLimite = dblImporteLimite + dblImporteMensual
            Next i
            datFechaControl = 0
        End If
        If DenominacionSLAVE = "DeduccionEspecial" Then
            If GciaNeta < dblImporteLimite Then
                dblImporteLimite = GciaNeta
            End If
        End If
    Case Else
        If ID = "Ninguno" Then
            dblImporteLimite = 0
        Else
            '1) Debemos determinar si hay Cargas de Familia declaradas por F572 Web
            SQL = "Select * From CargasFamiliaSIRADIG Where "
            Select Case DenominacionSLAVE
            Case "Conyuge"
            SQL = SQL & "(ID = '" & ID & "' " _
            & "And CodigoParentesco = '1')"
            Case "Hijo"
            SQL = SQL & "(ID = '" & ID & "' " _
            & "And CodigoParentesco = '3') OR " _
            & "(ID = '" & ID & "' " _
            & "And CodigoParentesco = '30') OR " _
            & "(ID = '" & ID & "' " _
            & "And CodigoParentesco = '31') OR " _
            & "(ID = '" & ID & "' " _
            & "And CodigoParentesco = '32')"
            Case "OtrasCargasDeFamilia"
            SQL = SQL & "(ID = '" & ID & "' " _
            & "And CodigoParentesco = '35')"
            End Select
            If SQLNoMatch(SQL) = True Then
                dblImporteLimite = 0
            Else
                '2) En caso de que haya cargas de familia informadas por F 572 Web, _
                debemos buscar el límite legal para la deducción
                Set rstRegistroSlave = New ADODB.Recordset
                rstRegistroSlave.Open SQL, dbSlave, adOpenDynamic, adLockOptimistic
                While rstRegistroSlave.EOF = False
                    'Determinamos el mes de inicio de la carga de familia
                    inPrincipio = rstRegistroSlave!MesDesde
                    'Determinamos el mes de Fin de la carga de familia
                    inFin = rstRegistroSlave!MesHasta
                    If inFin > Month(datFecha) Then
                        inFin = Month(datFecha)
                    End If
                    'Procedemos en la medida que el mes de inicio sea menor al mes que estamos liquidando
                    If inPrincipio <= Month(datFecha) Then '10/01/2018 Antes estaba "<"
                        dblImporteMensual = ImporteDeduccionPersonal(DenominacionSLAVE, DateSerial(Year(datFecha) - 1, 12, 31))
                        SQL = "Select " & DenominacionSLAVE & " From DEDUCCIONES4TACATEGORIA Where FECHA <= #" & Format(datFecha, "MM/DD/YYYY") & "# " _
                        & "And Year(FECHA) = '" & Year(datFecha) & "' Order by FECHA Desc"
                        If SQLNoMatch(SQL) = True Then
                            dblImporteLimite = dblImporteMensual * (inFin - inPrincipio + 1)
                        Else
                            For i = inPrincipio To inFin
                                datFechaControl = DateSerial(Year(datFecha), i, 1)
                                datFechaControl = DateAdd("m", 1, datFechaControl)
                                datFechaControl = DateAdd("d", -1, datFechaControl)
                                dblImporteMensual = ImporteDeduccionPersonal(DenominacionSLAVE, datFechaControl)
                                dblImporteLimite = dblImporteLimite + dblImporteMensual
                            Next i
                        End If
                    End If
                    rstRegistroSlave.MoveNext
                Wend
                rstRegistroSlave.Close
                Set rstRegistroSlave = Nothing
            End If
        End If
    End Select
    '3) debemos determinar el importe ya registrado hasta la liquidación previa y comparar con 2) (sacar diferencia)
    dblImporteAcumulado = ImporteRegistradoAcumuladoDeduccionEspecifica(DenominacionSLAVE, _
    PuestoLaboral, strPeriodo, CodigoLiquidacion, False)
    dblImporteCalculado = dblImporteLimite - dblImporteAcumulado
    
    CalcularDeduccionPersonalEspecificaSIRADIG = Round(dblImporteCalculado, 2)
    
    SQL = ""
    dblImporteCalculado = 0
    
End Function


Public Function ImporteRegistradoAcumuladoDeduccionEspecifica(Deduccion As String, PuestoLaboral As String, _
MesBarraAño As String, CodigoLiquidacion As String, Optional IncluirLiquidacionActual As Boolean = True) As Double

    Dim SQL As String
    Dim strSimboloSQL As String
    Dim dblImporteCalculado As Double
    
    If IncluirLiquidacionActual = True Then
        strSimboloSQL = "<="
    Else
        strSimboloSQL = "<"
    End If
    
    'TENER CUIDADO con este PROCEDIMENTO!!! Deducción Acumulada INCLUYENDO la presente liquidación
    SQL = "Select  Sum(LIQUIDACIONGANANCIAS4TACATEGORIA." & Deduccion & ") AS SumaDeduccion" _
    & " From CODIGOLIQUIDACIONES Inner Join LIQUIDACIONGANANCIAS4TACATEGORIA" _
    & " On CODIGOLIQUIDACIONES.Codigo = LIQUIDACIONGANANCIAS4TACATEGORIA.CodigoLiquidacion" _
    & " Where PUESTOLABORAL = '" & PuestoLaboral _
    & "' And Right(PERIODO,4) = '" & Right(MesBarraAño, 4) _
    & "' And CODIGO " & strSimboloSQL & " '" & CodigoLiquidacion & "'"
    If SQLNoMatch(SQL) = True Then
        dblImporteCalculado = 0
    Else
        Set rstBuscarSlave = New ADODB.Recordset
        rstBuscarSlave.Open SQL, dbSlave, adOpenDynamic, adLockOptimistic
        dblImporteCalculado = rstBuscarSlave!SumaDeduccion
        rstBuscarSlave.Close
        Set rstBuscarSlave = Nothing
    End If
    
    ImporteRegistradoAcumuladoDeduccionEspecifica = Round(dblImporteCalculado, 2)
    
    SQL = ""
    dblImporteCalculado = 0
    
End Function

Public Function CalcularDeduccionesPersonalesAcumuladas(PuestoLaboral As String, _
CodigoLiquidacion As String, Optional IncluirLiquidacionActual As Boolean = True, _
Optional DeduccionesAntesDeDeduccionEspecial As Boolean = False) As Double

    Dim SQL As String
    Dim strPeriodo As String
    Dim dblImporteCalculado As Double
    Dim strSimboloSQL As String
    
    If IncluirLiquidacionActual = True Then
        strSimboloSQL = "<="
    Else
        strSimboloSQL = "<"
    End If
    
    'TENER CUIDADO con este PROCEDIMENTO!!! Deducciones Acumuladas INCLUYENDO la presente liquidación
    'Nos posicionamos en el último día del período a liquidar
    strPeriodo = BuscarPeriodoLiquidacion(CodigoLiquidacion)
    SQL = "Select Sum(MINIMONOIMPONIBLE) as MNO, Sum(CONYUGE) as C," _
    & " Sum(HIJO) as H, Sum(OTRASCARGASDEFAMILIA) as OT, Sum(DEDUCCIONESPECIAL) as DE" _
    & " From CODIGOLIQUIDACIONES Inner Join LIQUIDACIONGANANCIAS4TACATEGORIA" _
    & " On CODIGOLIQUIDACIONES.Codigo = LIQUIDACIONGANANCIAS4TACATEGORIA.CodigoLiquidacion" _
    & " Where PUESTOLABORAL = '" & PuestoLaboral _
    & "' And Right(PERIODO,4) = '" & Right(strPeriodo, 4) _
    & "' And CODIGO " & strSimboloSQL & " '" & CodigoLiquidacion & "'"
    If SQLNoMatch(SQL) = True Then
        dblImporteCalculado = 0
    Else
        Set rstBuscarSlave = New ADODB.Recordset
        rstBuscarSlave.Open SQL, dbSlave, adOpenDynamic, adLockOptimistic
        dblImporteCalculado = rstBuscarSlave!MNO + rstBuscarSlave!C + rstBuscarSlave!H _
        + rstBuscarSlave!OT
        If DeduccionesAntesDeDeduccionEspecial = False Then
            dblImporteCalculado = dblImporteCalculado + rstBuscarSlave!de
        End If
        rstBuscarSlave.Close
        Set rstBuscarSlave = Nothing
    End If
    
    CalcularDeduccionesPersonalesAcumuladas = Round(dblImporteCalculado, 2)
    
    strSimboloSQL = ""
    SQL = ""
    strPeriodo = ""
    dblImporteCalculado = 0
    
End Function

Public Function CalcularDeduccionesGeneralesAcumuladas(PuestoLaboral As String, _
CodigoLiquidacion As String, Optional IncluirLiquidacionActual As Boolean = True, _
Optional DeduccionesAntesDeCMAyDyHM As Boolean = False, _
Optional IncluirDescuntosLegalesObligatorios As Boolean = False) As Double

    Dim SQL As String
    Dim strPeriodo As String
    Dim dblImporteCalculado As Double
    Dim strSimboloSQL As String
    Dim strSimboloSQL2 As String
    
    If IncluirDescuntosLegalesObligatorios = True Then
        strSimboloSQL = ", Sum(JUBILACION) as Jub, Sum(OBRASOCIAL) as OS," _
        & " Sum(ADHERENTEOBRASOCIAL) as AOS, Sum(SEGURODEVIDAOBLIGATORIO) as SVObl"
    Else
        strSimboloSQL = ""
    End If
    
    
    If IncluirLiquidacionActual = True Then
        strSimboloSQL2 = "<="
    Else
        strSimboloSQL2 = "<"
    End If
    
    'TENER CUIDADO con este PROCEDIMENTO!!! Deducciones Acumuladas INCLUYENDO la presente liquidación
    'Nos posicionamos en el último día del período a liquidar
    strPeriodo = BuscarPeriodoLiquidacion(CodigoLiquidacion)
    SQL = "Select Sum(SERVICIODOMESTICO) as SM, Sum(CUOTAMEDICOASISTENCIAL) as CMA," _
    & " Sum(ALQUILERES) as ALQ, Sum(DONACIONES) as Don, Sum(HONORARIOSMEDICOS) as HM," _
    & " Sum(SEGURODEVIDAOPTATIVO) as SVOpt, Sum(CUOTASINDICAL) as CS" & strSimboloSQL _
    & " From CODIGOLIQUIDACIONES Inner Join LIQUIDACIONGANANCIAS4TACATEGORIA" _
    & " On CODIGOLIQUIDACIONES.Codigo = LIQUIDACIONGANANCIAS4TACATEGORIA.CodigoLiquidacion" _
    & " Where PUESTOLABORAL = '" & PuestoLaboral _
    & "' And Right(PERIODO,4) = '" & Right(strPeriodo, 4) _
    & "' And CODIGO " & strSimboloSQL2 & " '" & CodigoLiquidacion & "'"
    If SQLNoMatch(SQL) = True Then
        dblImporteCalculado = 0
    Else
        Set rstBuscarSlave = New ADODB.Recordset
        rstBuscarSlave.Open SQL, dbSlave, adOpenDynamic, adLockOptimistic
        dblImporteCalculado = rstBuscarSlave!SM + rstBuscarSlave!SVOpt + rstBuscarSlave!ALQ _
        + rstBuscarSlave!CS
        If IncluirDescuntosLegalesObligatorios = True Then
            dblImporteCalculado = dblImporteCalculado + rstBuscarSlave!Jub _
            + rstBuscarSlave!OS + rstBuscarSlave!AOS + rstBuscarSlave!SVObl
        End If
        If DeduccionesAntesDeCMAyDyHM = False Then
            dblImporteCalculado = dblImporteCalculado + rstBuscarSlave!CMA _
            + rstBuscarSlave!Don + rstBuscarSlave!HM
        End If
        rstBuscarSlave.Close
        Set rstBuscarSlave = Nothing
    End If
    
    CalcularDeduccionesGeneralesAcumuladas = Round(dblImporteCalculado, 2)
    
    SQL = ""
    strPeriodo = ""
    dblImporteCalculado = 0
    
End Function

Public Function CalcularBaseImponible(PuestoLaboral As String, CodigoLiquidacion As String) As Double

    Dim dblBICalculado As Double
    
    'TENER CUIDADO con este PROCEDIMENTO!!! Base Imponible INCLUYENDO la presente liquidación
    'Calculamos el Sueldo Neto Acumulado a la presente liquidación
    dblBICalculado = CalcularSueldoNetoAcumulado(PuestoLaboral, CodigoLiquidacion)
    'Le restamos las Deducciones Personales. _
    & 'TENER CUIDADO con este PROCEDIMENTO!!! Deducciones Acumuladas INCLUYENDO la presente liquidación
    dblBICalculado = dblBICalculado - CalcularDeduccionesPersonalesAcumuladas(PuestoLaboral, CodigoLiquidacion)
    'Le restamos las Deducciones Generales. _
    & 'TENER CUIDADO con este PROCEDIMENTO!!! Deducciones Acumuladas INCLUYENDO la presente liquidación
    dblBICalculado = dblBICalculado - CalcularDeduccionesGeneralesAcumuladas(PuestoLaboral, CodigoLiquidacion)
       
    CalcularBaseImponible = Round(dblBICalculado, 2)
    
    dblBICalculado = 0
       
End Function

Public Function CalcularAlicuotaAplicable(PuestoLaboral As String, CodigoLiquidacion As String, Optional BaseImponible As Variant = "NaN") As Double

    Dim dblAlicuotaCalculada As Double
    Dim dblBI As Double
    Dim strPeriodo As String
    Dim datFecha As Date
    Dim SQL As String
    
    'Si la función no trae la base imponible
    If IsNumeric(BaseImponible) = False Then
        'TENER CUIDADO con este PROCEDIMENTO!!! Base Imponible INCLUYENDO la presente liquidación
        dblBI = CalcularBaseImponible(PuestoLaboral, CodigoLiquidacion)
    Else
        dblBI = BaseImponible
    End If
    'Ajustamos la BI en terminos anuales
    strPeriodo = BuscarPeriodoLiquidacion(CodigoLiquidacion)
    datFecha = BuscarUltimoDiaDelPeriodo(strPeriodo)
    dblBI = dblBI / CDbl(Left(strPeriodo, 2))
    dblBI = dblBI * 12
    'dblBI = Round(dblBI, 0) 'Tengo problemas en SQL con la ","
    SQL = "Select * From ESCALAAPLICABLEGANANCIAS" _
    & " Where IMPORTEMAXIMO > " & De_Num_a_Tx_01(dblBI, , 2) _
    & " And FECHA <= #" & Format(datFecha, "MM/DD/YYYY") _
    & "# Order by FECHA Desc, IMPORTEMAXIMO Asc"
    Set rstProcedimientoSlave = New ADODB.Recordset
    rstProcedimientoSlave.Open SQL, dbSlave, adOpenDynamic, adLockOptimistic
    dblAlicuotaCalculada = rstProcedimientoSlave!ImporteVariable
    rstProcedimientoSlave.Close

    CalcularAlicuotaAplicable = Round(dblAlicuotaCalculada, 2)
       
    Set rstProcedimientoSlave = Nothing
    dblAlicuotaCalculada = 0
    dblBI = 0
    SQL = ""
    strPeriodo = ""
       
End Function

Public Function CalcularImporteVariable(PuestoLaboral As String, CodigoLiquidacion As String, _
Optional AlicuotaAplicable As Variant = "NaN", Optional BaseImponible As Variant = "NaN") As Double

    Dim dblImporteVariableCalculado As Double
    Dim dblAlicuota As Double
    Dim dblBI As Double
    Dim datFecha As Date
    Dim strPeriodo As String
    Dim SQL As String
    
    'Si la función no trae la base imponible
    If IsNumeric(BaseImponible) = False Then
        'TENER CUIDADO con este PROCEDIMENTO!!! Base Imponible INCLUYENDO la presente liquidación
        dblBI = CalcularBaseImponible(PuestoLaboral, CodigoLiquidacion)
    Else
        dblBI = BaseImponible
    End If
    'Si la función no trae alicuotaaplicable
    If IsNumeric(AlicuotaAplicable) = False Then
        'TENER CUIDADO con este PROCEDIMENTO!!! Base Imponible INCLUYENDO la presente liquidación
        dblAlicuota = CalcularAlicuotaAplicable(PuestoLaboral, CodigoLiquidacion, dblBI)
    Else
        dblAlicuota = AlicuotaAplicable
    End If
    'dblAlicuota = Round(dblAlicuota, 0) 'Tengo problemas en SQL con la ","
    strPeriodo = BuscarPeriodoLiquidacion(CodigoLiquidacion)
    datFecha = BuscarUltimoDiaDelPeriodo(strPeriodo)
    SQL = "Select * From ESCALAAPLICABLEGANANCIAS" _
    & " Where IMPORTEVARIABLE < " & De_Num_a_Tx_01(dblAlicuota, , 2) _
    & " And FECHA <= #" & Format(datFecha, "MM/DD/YYYY") _
    & "# Order by FECHA Desc, IMPORTEMAXIMO Desc"
    If SQLNoMatch(SQL) = True Then
        dblImporteVariableCalculado = 0
    Else
        Set rstProcedimientoSlave = New ADODB.Recordset
        rstProcedimientoSlave.Open SQL, dbSlave, adOpenDynamic, adLockOptimistic
        dblImporteVariableCalculado = rstProcedimientoSlave!ImporteMaximo
        rstProcedimientoSlave.Close
        'Ajustamos el Importe Maximo encontrado al periodo requerido
        dblImporteVariableCalculado = (dblImporteVariableCalculado / 12) * CDbl(Left(strPeriodo, 2))
    End If
    'Procedemos a calcular el importe variable
    dblImporteVariableCalculado = dblBI - dblImporteVariableCalculado
    dblImporteVariableCalculado = dblImporteVariableCalculado * dblAlicuota

    CalcularImporteVariable = Round(dblImporteVariableCalculado, 2)
       
    Set rstProcedimientoSlave = Nothing
    dblImporteVariableCalculado = 0
    dblBI = 0
    dblAlicuota = 0
    SQL = ""
    strPeriodo = ""
       
End Function

Public Function CalcularImporteFijo(PuestoLaboral As String, CodigoLiquidacion As String, Optional BaseImponible As Variant = "NaN") As Double

    Dim dblImporteFijoCalculado As Double
    Dim dblBI As Double
    Dim strPeriodo As String
    Dim datFecha As Date
    Dim SQL As String
    
    'Si la función no trae la base imponible
    If IsNumeric(BaseImponible) = False Then
        'TENER CUIDADO con este PROCEDIMENTO!!! Base Imponible INCLUYENDO la presente liquidación
        dblBI = CalcularBaseImponible(PuestoLaboral, CodigoLiquidacion)
    Else
        dblBI = BaseImponible
    End If
    'Ajustamos la BI en terminos anuales
    strPeriodo = BuscarPeriodoLiquidacion(CodigoLiquidacion)
    datFecha = BuscarUltimoDiaDelPeriodo(strPeriodo)
    dblBI = dblBI / CDbl(Left(strPeriodo, 2))
    dblBI = dblBI * 12
    'dblBI = Round(dblBI, 0) 'Tengo problemas en SQL con la ","
    SQL = "Select * From ESCALAAPLICABLEGANANCIAS" _
    & " Where IMPORTEMAXIMO > " & De_Num_a_Tx_01(dblBI, , 2) _
    & " And FECHA <= #" & Format(datFecha, "MM/DD/YYYY") _
    & "# Order by FECHA Desc, IMPORTEMAXIMO Asc"
    Set rstProcedimientoSlave = New ADODB.Recordset
    rstProcedimientoSlave.Open SQL, dbSlave, adOpenDynamic, adLockOptimistic
    dblImporteFijoCalculado = rstProcedimientoSlave!ImporteFijo
    rstProcedimientoSlave.Close
    'Ajustamos el Importe Fijo al período de liquidación
    dblImporteFijoCalculado = dblImporteFijoCalculado / 12
    dblImporteFijoCalculado = dblImporteFijoCalculado * CDbl(Left(strPeriodo, 2))

    CalcularImporteFijo = Round(dblImporteFijoCalculado, 2)
       
    Set rstProcedimientoSlave = Nothing
    dblImporteFijoCalculado = 0
    dblBI = 0
    SQL = ""
    strPeriodo = ""
       
End Function

Public Function CalcularRetencionAcumulada(PuestoLaboral As String, _
CodigoLiquidacion As String, Optional IncluirLiquidacionActual As Boolean = True) As Double

    Dim SQL As String
    Dim dblRetencionCalculada As Double
    Dim strPeriodo As String
    Dim strSimboloSQL As String
    
    If IncluirLiquidacionActual = True Then
        strSimboloSQL = "<="
    Else
        strSimboloSQL = "<"
    End If
    
    'TENER CUIDADO con este PROCEDIMENTO!!! Retencion Acumulada INCLUYENDO la presente liquidación
    strPeriodo = BuscarPeriodoLiquidacion(CodigoLiquidacion)
    Set rstProcedimientoSlave = New ADODB.Recordset
    'Cargamos la Retencion Acumulada
    SQL = "Select Sum(LIQUIDACIONGANANCIAS4TACATEGORIA.Retencion) AS SumaDeImporte" _
    & " From CODIGOLIQUIDACIONES Inner Join LIQUIDACIONGANANCIAS4TACATEGORIA" _
    & " On CODIGOLIQUIDACIONES.Codigo = LIQUIDACIONGANANCIAS4TACATEGORIA.CodigoLiquidacion" _
    & " Where PUESTOLABORAL = '" & PuestoLaboral _
    & "' And Right(PERIODO,4) = '" & Right(strPeriodo, 4) _
    & "' And CODIGO " & strSimboloSQL & " '" & CodigoLiquidacion & "'" 'Menor y/o IGUAL a la presente liquidación
    If SQLNoMatch(SQL) = True Then
        dblRetencionCalculada = 0
    Else
        rstProcedimientoSlave.Open SQL, dbSlave, adOpenDynamic, adLockOptimistic
        dblRetencionCalculada = rstProcedimientoSlave!SumaDeImporte
        rstProcedimientoSlave.Close
    End If
    
    CalcularRetencionAcumulada = Round(dblRetencionCalculada, 2)
    
    SQL = ""
    strSimboloSQL = ""
    dblRetencionCalculada = 0
    strPeriodo = ""
    If rstProcedimientoSlave.State = adStateOpen Then
        rstProcedimientoSlave.Close
    End If
    Set rstProcedimientoSlave = Nothing
   
End Function

Public Function CalcularAjusteRetencion(PuestoLaboral As String, CodigoLiquidacion As String) As Double

    Dim SQL As String
    Dim dblAjusteCalculado As Double
    Dim strPeriodo As String
    
    strPeriodo = BuscarPeriodoLiquidacion(CodigoLiquidacion)
    'Cargamos Ajuste Retenciones (VERIFICAR PROCEDIMIENTO)
    SQL = "Select * From CODIGOLIQUIDACIONES Inner Join LIQUIDACIONGANANCIAS4TACATEGORIA " _
    & " On CODIGOLIQUIDACIONES.Codigo = LIQUIDACIONGANANCIAS4TACATEGORIA.CodigoLiquidacion" _
    & " Where PUESTOLABORAL = '" & PuestoLaboral _
    & "' And CODIGOLIQUIDACION <= '" & CodigoLiquidacion _
    & "' And Right(PERIODO,4) = '" & Right(strPeriodo, 4) _
    & "' Order By CODIGOLIQUIDACION Desc"
    If SQLNoMatch(SQL) = True Then
        dblAjusteCalculado = 0
    Else
        Set rstProcedimientoSlave = New ADODB.Recordset
        rstProcedimientoSlave.Open SQL, dbSlave, adOpenForwardOnly, adLockOptimistic
        If rstProcedimientoSlave!AjusteRetencion = 0 Then
            dblAjusteCalculado = 0
        Else
            'Si ya tiene un valor previamente cargado, lo volvemos a cargar
            dblAjusteCalculado = rstProcedimientoSlave!AjusteRetencion
        End If
        rstProcedimientoSlave.Close
        Set rstProcedimientoSlave = Nothing
    End If
    
    CalcularAjusteRetencion = Round(dblAjusteCalculado, 2)
    
    SQL = ""
    dblAjusteCalculado = 0
   
End Function

Public Function CalcularRetencionDelPeriodo(PuestoLaboral As String, CodigoLiquidacion As String) As Double

    Dim SQL As String
    Dim dblRetencionDelPeriodoCalculada As Double
    Dim dblBaseImponible As Double
    Dim strCodigoLiquidacion As String
    
    'Buscamos el período anterior
    strCodigoLiquidacion = BuscarCodigoLiquidacionAnterior(CodigoLiquidacion)
    'Calculamos la Base Imponible
    dblBaseImponible = CalcularBaseImponible(PuestoLaboral, CodigoLiquidacion)
    'Calculamos el Importe Variable (VERIFICAR PROCEDIMIENTO)
    dblRetencionDelPeriodoCalculada = CalcularImporteVariable(PuestoLaboral, CodigoLiquidacion, , dblBaseImponible)
    'Calculamos y sumamos el Importe fijo
    dblRetencionDelPeriodoCalculada = dblRetencionDelPeriodoCalculada + _
    CalcularImporteFijo(PuestoLaboral, CodigoLiquidacion, dblBaseImponible)
    'Obtenemos y restamos la Retencion Acumulada
    dblRetencionDelPeriodoCalculada = dblRetencionDelPeriodoCalculada - _
    CalcularRetencionAcumulada(PuestoLaboral, strCodigoLiquidacion)
    'Obtenemos y sumamos el Ajuste de Retención
    dblRetencionDelPeriodoCalculada = dblRetencionDelPeriodoCalculada + _
    CalcularAjusteRetencion(PuestoLaboral, strCodigoLiquidacion)
    
    CalcularRetencionDelPeriodo = Round(dblRetencionDelPeriodoCalculada, 2)
    
    SQL = ""
    dblRetencionDelPeriodoCalculada = 0
    dblBaseImponible = 0
    strCodigoLiquidacion = ""
   
End Function

Public Function BuscarCodigoLiquidacion(MesBarraAno As String, _
Optional SoloPares As Boolean = False) As String

    Dim SQL As String
    Dim SQLPar As String
    
    If SoloPares = True Then
        SQLPar = " And CInt(CODIGO) Mod 2 = 0"
    Else
        SQLPar = ""
    End If
    
    SQL = "Select CODIGO From CODIGOLIQUIDACIONES" _
    & " Where Right(PERIODO, 4) = '" & Right(MesBarraAno, 4) & "'" _
    & " And Left(PERIODO, 2) <= '" & Left(MesBarraAno, 2) & "'" _
    & SQLPar _
    & " Order by CODIGO Desc"
    If SQLNoMatch(SQL) = False Then
        Set rstBuscarSlave = New ADODB.Recordset
        rstBuscarSlave.Open SQL, dbSlave, adOpenForwardOnly, adLockOptimistic
        rstBuscarSlave.MoveFirst
        BuscarCodigoLiquidacion = rstBuscarSlave!Codigo
        rstBuscarSlave.Close
        Set rstBuscarSlave = Nothing
    Else
        BuscarCodigoLiquidacion = 0
    End If

    
    SQL = ""
    
End Function

Public Function BuscarCodigoLiquidacionAnterior(CodigoLiquidacionBase As String, _
Optional SoloPares As Boolean = False) As String

    Dim SQL As String
    Dim SQLPar As String
    
    If SoloPares = True Then
        SQLPar = " And CInt(CODIGO) Mod 2 = 0"
    Else
        SQLPar = ""
    End If
    
    SQL = "Select CODIGO From CODIGOLIQUIDACIONES" _
    & " Where CODIGO < '" & CodigoLiquidacionBase & "'" _
    & SQLPar _
    & " Order by CODIGO Desc"
    Set rstBuscarSlave = New ADODB.Recordset
    rstBuscarSlave.Open SQL, dbSlave, adOpenForwardOnly, adLockOptimistic
    rstBuscarSlave.MoveFirst
    BuscarCodigoLiquidacionAnterior = rstBuscarSlave!Codigo
    rstBuscarSlave.Close
    Set rstBuscarSlave = Nothing
    
    SQL = ""
    
End Function

Public Function BuscarCodigoLiquidacionGananciasAnterior(CodigoLiquidacionBase As String) As String

    Dim SQL As String
    
    SQL = "Select CODIGOLIQUIDACION From LIQUIDACIONGANANCIAS4TACATEGORIA" _
    & " Group by CODIGOLIQUIDACION" _
    & " Having CODIGOLIQUIDACION < '" & CodigoLiquidacionBase _
    & "' Order by CODIGOLIQUIDACION Desc"
    Set rstBuscarSlave = New ADODB.Recordset
    rstBuscarSlave.Open SQL, dbSlave, adOpenForwardOnly, adLockOptimistic
    rstBuscarSlave.MoveFirst
    BuscarCodigoLiquidacionGananciasAnterior = rstBuscarSlave!CodigoLiquidacion
    rstBuscarSlave.Close
    Set rstBuscarSlave = Nothing
    
    SQL = ""
    
End Function

Public Function DiferenciaHaberBrutoAcumuladoSISPERvsSLAVE(PuestoLaboral As String, _
CodigoLiquidacion As String, HaberBrutoAIncorporarSLAVE As Double, _
Optional EsLiquidacionFinal As Boolean = False) As Double

    Dim SQL As String
    Dim dblAcumuladoSLAVE As Double
    Dim dblAcumuladoSISPER As Double
    Dim strPL As String
        
    'Controlamos lo que liquidó el SISPER hasta la presente liquidación (la misma puede ser proyecta y no la real)
    dblAcumuladoSISPER = CalcularHaberBrutoAcumulado(PuestoLaboral, CodigoLiquidacion, , EsLiquidacionFinal)
    strPL = BuscarPeriodoLiquidacion(CodigoLiquidacion)
    If EsLiquidacionFinal = False And CInt(Right(strPL, 4)) >= 2017 Then 'Incremento de 1/12 del Haber Bruto según Ley 27346/16 (Macri)
        dblAcumuladoSISPER = dblAcumuladoSISPER * (13 / 12)
    End If
    'Controlamos lo que tiene cargado SLAVE hasta la liquidacion anterior
    SQL = "Select Sum(HaberOptimo) As HO, Sum(Ajuste) as AJ" _
    & " From CODIGOLIQUIDACIONES Inner Join LIQUIDACIONGANANCIAS4TACATEGORIA" _
    & " On CODIGOLIQUIDACIONES.Codigo = LIQUIDACIONGANANCIAS4TACATEGORIA.CodigoLiquidacion" _
    & " Where PUESTOLABORAL = '" & PuestoLaboral & "'" _
    & " And Right(PERIODO,4) = '" & Right(strPL, 4) & "'" _
    & " And CODIGO < '" & CodigoLiquidacion & "'"
    If SQLNoMatch(SQL) = False Then
        Set rstProcedimientoSlave = New ADODB.Recordset
        rstProcedimientoSlave.Open SQL, dbSlave, adOpenForwardOnly, adLockOptimistic
        dblAcumuladoSLAVE = rstProcedimientoSlave!HO '+ rstProcedimientoSlave!AJ
        dblAcumuladoSLAVE = Round(dblAcumuladoSLAVE, 2)
        rstProcedimientoSlave.Close
    Else
        dblAcumuladoSLAVE = 0
    End If
    'Le sumamos lo que pretendemos incorporar en la presente liquidacion de ganancias
    dblAcumuladoSLAVE = dblAcumuladoSLAVE + HaberBrutoAIncorporarSLAVE
    dblAcumuladoSLAVE = Round(dblAcumuladoSLAVE, 2)
    'Calculamos la diferencia entre los Haberes Brutos SISPER y SLAVE
    DiferenciaHaberBrutoAcumuladoSISPERvsSLAVE = Round(dblAcumuladoSISPER - dblAcumuladoSLAVE, 2)
    
    SQL = ""
    dblAcumuladoSLAVE = 0
    dblAcumuladoSISPER = 0
    strPL = ""

End Function

Public Function DiferenciaCuotaSindicalAcumuladaSISPERvsSLAVE(PuestoLaboral As String, _
CodigoLiquidacion As String, CuotaSindicalAIncorporarSLAVE As Double, _
Optional EsLiquidacionFinal As Boolean = False) As Double

    Dim SQL As String
    Dim dblAcumuladoSLAVE As Double
    Dim dblAcumuladoSISPER As Double
    Dim strPL As String
        
    'Controlamos lo que liquidó el SISPER hasta la presente liquidación (la misma puede ser proyecta y no la real)
    dblAcumuladoSISPER = CalcularDescuentoCuotaSindicalAcumulado(PuestoLaboral, CodigoLiquidacion, , EsLiquidacionFinal)
    strPL = BuscarPeriodoLiquidacion(CodigoLiquidacion)
    If EsLiquidacionFinal = True And CInt(Right(strPL, 4)) >= 2017 Then 'Incremento de 1/12 del Haber Bruto según Ley 27346/16 (Macri)
        dblAcumuladoSISPER = dblAcumuladoSISPER / (13 / 12)
    End If
    'Controlamos lo que tiene cargado SLAVE hasta la liquidacion anterior
    SQL = "Select Sum(CuotaSindical) As CS" _
    & " From CODIGOLIQUIDACIONES Inner Join LIQUIDACIONGANANCIAS4TACATEGORIA" _
    & " On CODIGOLIQUIDACIONES.Codigo = LIQUIDACIONGANANCIAS4TACATEGORIA.CodigoLiquidacion" _
    & " Where PUESTOLABORAL = '" & PuestoLaboral & "'" _
    & " And Right(PERIODO,4) = '" & Right(strPL, 4) & "'" _
    & " And CODIGO < '" & CodigoLiquidacion & "'"
    If SQLNoMatch(SQL) = False Then
        Set rstProcedimientoSlave = New ADODB.Recordset
        rstProcedimientoSlave.Open SQL, dbSlave, adOpenForwardOnly, adLockOptimistic
        dblAcumuladoSLAVE = rstProcedimientoSlave!CS
        dblAcumuladoSLAVE = Round(dblAcumuladoSLAVE, 2)
        rstProcedimientoSlave.Close
    Else
        dblAcumuladoSLAVE = 0
    End If
    'Le sumamos lo que pretendemos incorporar en la presente liquidacion de ganancias
    dblAcumuladoSLAVE = dblAcumuladoSLAVE + CuotaSindicalAIncorporarSLAVE
    dblAcumuladoSLAVE = Round(dblAcumuladoSLAVE, 2)
    'Calculamos la diferencia entre los Haberes Brutos SISPER y SLAVE
    DiferenciaCuotaSindicalAcumuladaSISPERvsSLAVE = Round(dblAcumuladoSISPER - dblAcumuladoSLAVE, 2)
    
    SQL = ""
    dblAcumuladoSLAVE = 0
    dblAcumuladoSISPER = 0
    strPL = ""

End Function

Public Function CantidadHijosDeducibles(PuestoLaboral As String, MesYAñoLiquidacion As Date) As Integer

    Dim dat24Years As Date
    Dim dat18Years As Date
    Dim SQL As String
    
    dat24Years = DateSerial(Year(MesYAñoLiquidacion), Month(MesYAñoLiquidacion), 1)
    dat18Years = dat24Years
    dat24Years = DateAdd("yyyy", -24, dat24Years)
    dat18Years = DateAdd("yyyy", -18, dat18Years)
    
    If Year(MesYAñoLiquidacion) < 2017 Then
        SQL = "Select * From CARGASDEFAMILIA Where ((PUESTOLABORAL = '" & PuestoLaboral & "' And CODIGOPARENTESCO = '1' And DEDUCIBLEGANANCIAS = TRUE And FECHAALTA >= #" & Format(dat24Years, "MM/DD/YYYY") & "#) " _
        & "Or (PUESTOLABORAL = '" & PuestoLaboral & "' And CODIGOPARENTESCO = '3' And DEDUCIBLEGANANCIAS = TRUE And FECHAALTA >= #" & Format(dat24Years, "MM/DD/YYYY") & "#) " _
        & "Or (PUESTOLABORAL = '" & PuestoLaboral & "' And CODIGOPARENTESCO = '5' And DEDUCIBLEGANANCIAS = TRUE And FECHAALTA >= #" & Format(dat24Years, "MM/DD/YYYY") & "#) " _
        & "Or (PUESTOLABORAL = '" & PuestoLaboral & "' And CODIGOPARENTESCO = '31' And DEDUCIBLEGANANCIAS = TRUE And FECHAALTA >= #" & Format(dat24Years, "MM/DD/YYYY") & "#) " _
        & "Or (PUESTOLABORAL = '" & PuestoLaboral & "' And CODIGOPARENTESCO = '1' And DEDUCIBLEGANANCIAS = TRUE And DISCAPACITADO = TRUE) " _
        & "Or (PUESTOLABORAL = '" & PuestoLaboral & "' And CODIGOPARENTESCO = '3' And DEDUCIBLEGANANCIAS = TRUE And DISCAPACITADO = TRUE) " _
        & "Or (PUESTOLABORAL = '" & PuestoLaboral & "' And CODIGOPARENTESCO = '5' And DEDUCIBLEGANANCIAS = TRUE And DISCAPACITADO = TRUE) " _
        & "Or (PUESTOLABORAL = '" & PuestoLaboral & "' And CODIGOPARENTESCO = '31' And DEDUCIBLEGANANCIAS = TRUE And DISCAPACITADO = TRUE))"
    Else
        SQL = "Select * From CARGASDEFAMILIA Where ((PUESTOLABORAL = '" & PuestoLaboral & "' And CODIGOPARENTESCO = '1' And DEDUCIBLEGANANCIAS = TRUE And FECHAALTA >= #" & Format(dat18Years, "MM/DD/YYYY") & "#) " _
        & "Or (PUESTOLABORAL = '" & PuestoLaboral & "' And CODIGOPARENTESCO = '3' And DEDUCIBLEGANANCIAS = TRUE And FECHAALTA >= #" & Format(dat18Years, "MM/DD/YYYY") & "#) " _
        & "Or (PUESTOLABORAL = '" & PuestoLaboral & "' And CODIGOPARENTESCO = '5' And DEDUCIBLEGANANCIAS = TRUE And FECHAALTA >= #" & Format(dat18Years, "MM/DD/YYYY") & "#) " _
        & "Or (PUESTOLABORAL = '" & PuestoLaboral & "' And CODIGOPARENTESCO = '31' And DEDUCIBLEGANANCIAS = TRUE And FECHAALTA >= #" & Format(dat18Years, "MM/DD/YYYY") & "#) " _
        & "Or (PUESTOLABORAL = '" & PuestoLaboral & "' And CODIGOPARENTESCO = '1' And DEDUCIBLEGANANCIAS = TRUE And DISCAPACITADO = TRUE) " _
        & "Or (PUESTOLABORAL = '" & PuestoLaboral & "' And CODIGOPARENTESCO = '3' And DEDUCIBLEGANANCIAS = TRUE And DISCAPACITADO = TRUE) " _
        & "Or (PUESTOLABORAL = '" & PuestoLaboral & "' And CODIGOPARENTESCO = '5' And DEDUCIBLEGANANCIAS = TRUE And DISCAPACITADO = TRUE) " _
        & "Or (PUESTOLABORAL = '" & PuestoLaboral & "' And CODIGOPARENTESCO = '31' And DEDUCIBLEGANANCIAS = TRUE And DISCAPACITADO = TRUE))"
    End If
    Set rstBuscarSlave = New ADODB.Recordset
    rstBuscarSlave.Open SQL, dbSlave, adOpenDynamic, adLockOptimistic
    CantidadHijosDeducibles = rstBuscarSlave.RecordCount
    rstBuscarSlave.Close
    Set rstBuscarSlave = Nothing
    dat24Years = 0
    dat18Years = 0
    SQL = ""
    
End Function

Public Function CantidadOtrasCargasFamiliaDeducibles(PuestoLaboral As String) As Integer

    Dim SQL As String
    
    SQL = "Select * From CARGASDEFAMILIA Where ((PUESTOLABORAL = '" & PuestoLaboral & "' And CODIGOPARENTESCO <> '1')" _
    & "And (PUESTOLABORAL = '" & PuestoLaboral & "' And CODIGOPARENTESCO <> '3')" _
    & "And (PUESTOLABORAL = '" & PuestoLaboral & "' And CODIGOPARENTESCO <> '5')" _
    & "And (PUESTOLABORAL = '" & PuestoLaboral & "' And CODIGOPARENTESCO <> '2')" _
    & "And (PUESTOLABORAL = '" & PuestoLaboral & "' And CODIGOPARENTESCO <> '31')" _
    & "And (PUESTOLABORAL = '" & PuestoLaboral & "' And CODIGOPARENTESCO <> '50')" _
    & "And (PUESTOLABORAL = '" & PuestoLaboral & "' And DEDUCIBLEGANANCIAS = TRUE))"
    Set rstBuscarSlave = New ADODB.Recordset
    rstBuscarSlave.Open SQL, dbSlave, adOpenDynamic, adLockOptimistic
    CantidadOtrasCargasFamiliaDeducibles = rstBuscarSlave.RecordCount
    rstBuscarSlave.Close
    Set rstBuscarSlave = Nothing
    SQL = ""
    
End Function

Public Function TieneConyugeDeducible(PuestoLaboral As String) As Boolean

    Dim SQL As String
    
    SQL = "Select * From CARGASDEFAMILIA Where ((PUESTOLABORAL = '" & PuestoLaboral & "' And CODIGOPARENTESCO = '2' And DEDUCIBLEGANANCIAS = TRUE) " _
    & "Or (PUESTOLABORAL = '" & PuestoLaboral & "' And CODIGOPARENTESCO = '50' And DEDUCIBLEGANANCIAS = TRUE) " _
    & "Or (PUESTOLABORAL = '" & PuestoLaboral & "' And CODIGOPARENTESCO = '51' And DEDUCIBLEGANANCIAS = TRUE))"
    If SQLNoMatch(SQL) = True Then
        TieneConyugeDeducible = False
    Else
        TieneConyugeDeducible = True
    End If
    
    SQL = ""
    
End Function

Public Function EsCargaDeFamiliaDeducibleSIRADIG(CodigoParentesco As String, FechaNacimiento As Date, FechaControl As Date) As Boolean

    Dim dat24Years As Date
    Dim dat18Years As Date
    Dim SQL As String
    
    EsCargaDeFamiliaDeducibleSIRADIG = False
    
    Select Case CodigoParentesco
    Case "1" 'Para Conyuge
        If ImporteDeduccionPersonal("Conyuge", FechaControl) > 0 Then
            EsCargaDeFamiliaDeducibleSIRADIG = True
        Else
            EsCargaDeFamiliaDeducibleSIRADIG = False
        End If
    Case "3", "30"  'Para Hijo/a o Hijastro/a
        If ImporteDeduccionPersonal("Hijo", FechaControl) > 0 Then
            'Hay que verificar la edad
            dat24Years = DateSerial(Year(FechaControl), Month(FechaControl), 1)
            dat18Years = dat24Years
            dat24Years = DateAdd("yyyy", -24, dat24Years)
            dat18Years = DateAdd("yyyy", -18, dat18Years)
            If Year(FechaControl) < 2017 Then
                If FechaNacimiento > dat24Years Then
                    EsCargaDeFamiliaDeducibleSIRADIG = True
                Else
                    EsCargaDeFamiliaDeducibleSIRADIG = False
                End If
            Else
                If FechaNacimiento > dat18Years Then
                    EsCargaDeFamiliaDeducibleSIRADIG = True
                Else
                    EsCargaDeFamiliaDeducibleSIRADIG = False
                End If
            End If
        Else
            EsCargaDeFamiliaDeducibleSIRADIG = False
        End If
    Case "31", "32"  'Para Hijo/a o Hijastro/a incapacitado
        If ImporteDeduccionPersonal("Hijo", FechaControl) > 0 Then
            EsCargaDeFamiliaDeducibleSIRADIG = True
        Else
            EsCargaDeFamiliaDeducibleSIRADIG = False
        End If
    Case Else 'Otras Cargas de Familia (Requiere Mayor Desagregación)
        If ImporteDeduccionPersonal("OtrasCargasDeFamilia", FechaControl) > 0 Then
            EsCargaDeFamiliaDeducibleSIRADIG = True
        Else
            EsCargaDeFamiliaDeducibleSIRADIG = False
        End If
    End Select
    
    dat24Years = 0
    dat18Years = 0
    SQL = ""
    
End Function

Public Function MesHastaCargaDeFamiliaDeducibleSIRADIG(CodigoParentesco As String, FechaNacimiento As Date, FechaControl As Date) As Integer

    Dim dat24Years As Date
    Dim dat18Years As Date
    Dim SQL As String
    
    MesHastaCargaDeFamiliaDeducibleSIRADIG = 0
    
    Select Case CodigoParentesco
    Case "1" 'Para Conyuge
        If ImporteDeduccionPersonal("Conyuge", FechaControl) > 0 Then
            MesHastaCargaDeFamiliaDeducibleSIRADIG = 12
        Else
            MesHastaCargaDeFamiliaDeducibleSIRADIG = 0
        End If
    Case "3", "30"  'Para Hijo/a o Hijastro/a
        If ImporteDeduccionPersonal("Hijo", FechaControl) > 0 Then
            'Hay que verificar la edad
            dat24Years = DateSerial(Year(FechaControl), Month(FechaControl), 1)
            dat18Years = dat24Years
            dat24Years = DateAdd("yyyy", -24, dat24Years)
            dat18Years = DateAdd("yyyy", -18, dat18Years)
            If Year(FechaControl) < 2017 Then
                If FechaNacimiento >= dat24Years Then
                    If (Year(FechaNacimiento) - Year(dat24Years)) > 0 Then
                        MesHastaCargaDeFamiliaDeducibleSIRADIG = 12
                    Else
                        MesHastaCargaDeFamiliaDeducibleSIRADIG = Month(FechaNacimiento)
                    End If
                Else
                    MesHastaCargaDeFamiliaDeducibleSIRADIG = 0
                End If
            Else
                If FechaNacimiento >= dat18Years Then
                    If (Year(FechaNacimiento) - Year(dat18Years)) > 0 Then
                        MesHastaCargaDeFamiliaDeducibleSIRADIG = 12
                    Else
                        MesHastaCargaDeFamiliaDeducibleSIRADIG = Month(FechaNacimiento)
                    End If
                Else
                    MesHastaCargaDeFamiliaDeducibleSIRADIG = 0
                End If
            End If
        Else
            MesHastaCargaDeFamiliaDeducibleSIRADIG = 0
        End If
    Case "31", "32"  'Para Hijo/a o Hijastro/a incapacitado
        If ImporteDeduccionPersonal("Hijo", FechaControl) > 0 Then
            MesHastaCargaDeFamiliaDeducibleSIRADIG = 12
        Else
            MesHastaCargaDeFamiliaDeducibleSIRADIG = 0
        End If
    Case Else 'Otras Cargas de Familia (Requiere Mayor Desagregación)
        If ImporteDeduccionPersonal("OtrasCargasDeFamilia", FechaControl) > 0 Then
            MesHastaCargaDeFamiliaDeducibleSIRADIG = 12
        Else
            MesHastaCargaDeFamiliaDeducibleSIRADIG = 0
        End If
    End Select
    
    dat24Years = 0
    dat18Years = 0
    SQL = ""
    
End Function


Public Function DescuentoSeguroDeVidaOptativoAcumulado(PuestoLaboral As String, _
CodigoLiquidacion As String) As Double

    Dim SQL As String
    Dim dblSVOptativoCalculado As Double
    Dim strPeriodo As String
    Dim datFecha As Date
     
    strPeriodo = BuscarPeriodoLiquidacion(CodigoLiquidacion)
     
    'Buscamos el Descuento Seguro Optativo
    Set rstProcedimientoSlave = New ADODB.Recordset
    SQL = "Select Sum(Importe) As SVOptAcum From LIQUIDACIONSUELDOS Inner Join CODIGOLIQUIDACIONES" _
    & " On LIQUIDACIONSUELDOS.CodigoLiquidacion = CODIGOLIQUIDACIONES.CODIGO" _
    & " Where ((PUESTOLABORAL = '" & PuestoLaboral & "' And CODIGOLIQUIDACION <= '" & CodigoLiquidacion _
    & "' And Right(Periodo,4) = '" & Right(strPeriodo, 4) _
    & "' And CODIGOCONCEPTO = '0317') Or " _
    & "(PUESTOLABORAL = '" & PuestoLaboral & "' And CODIGOLIQUIDACION <= '" & CodigoLiquidacion _
    & "' And Right(Periodo,4) = '" & Right(strPeriodo, 4) _
    & "' And CODIGOCONCEPTO = '0361') Or " _
    & "(PUESTOLABORAL = '" & PuestoLaboral & "' And CODIGOLIQUIDACION <= '" & CodigoLiquidacion _
    & "' And Right(Periodo,4) = '" & Right(strPeriodo, 4) _
    & "' And CODIGOCONCEPTO = '0367') Or " _
    & "(PUESTOLABORAL = '" & PuestoLaboral & "' And CODIGOLIQUIDACION <= '" & CodigoLiquidacion _
    & "' And Right(Periodo,4) = '" & Right(strPeriodo, 4) _
    & "' And CODIGOCONCEPTO = '0370') Or " _
    & "(PUESTOLABORAL = '" & PuestoLaboral & "' And CODIGOLIQUIDACION <= '" & CodigoLiquidacion _
    & "' And Right(Periodo,4) = '" & Right(strPeriodo, 4) _
    & "' And CODIGOCONCEPTO = '0373') Or " _
    & "(PUESTOLABORAL = '" & PuestoLaboral & "' And CODIGOLIQUIDACION <= '" & CodigoLiquidacion _
    & "' And Right(Periodo,4) = '" & Right(strPeriodo, 4) _
    & "' And CODIGOCONCEPTO = '0374'))"
    If SQLNoMatch(SQL) = True Then
        'Si no existe descuento Seguro Optativo
        dblSVOptativoCalculado = 0
    Else
        'Si existe descuento Seguro Optativo
        datFecha = BuscarUltimoDiaDelPeriodo(strPeriodo)
        rstProcedimientoSlave.Open SQL, dbSlave, adOpenForwardOnly, adLockOptimistic
        Set rstBuscarSlave = New ADODB.Recordset
        SQL = "Select SEGURODEVIDA From DEDUCCIONES4TACATEGORIA" _
        & " Where FECHA <= #" & Format(datFecha, "MM/DD/YYYY") _
        & "# Order by FECHA Desc"
        rstBuscarSlave.Open SQL, dbSlave, adOpenForwardOnly, adLockOptimistic
        If rstProcedimientoSlave!SVOptAcum < (rstBuscarSlave!SeguroDeVida / 12 * Left(strPeriodo, 2)) Then
            dblSVOptativoCalculado = rstProcedimientoSlave!SVOptAcum
        Else
            dblSVOptativoCalculado = (rstBuscarSlave!SeguroDeVida / 12 * Left(strPeriodo, 2))
        End If
        rstProcedimientoSlave.Close
        rstBuscarSlave.Close
        Set rstBuscarSlave = Nothing
    End If
    
    DescuentoSeguroDeVidaOptativoAcumulado = Round(dblSVOptativoCalculado, 2)
    
    SQL = ""
    dblSVOptativoCalculado = 0
    If rstProcedimientoSlave.State = adStateOpen Then
        rstProcedimientoSlave.Close
    End If
    Set rstProcedimientoSlave = Nothing
    strPeriodo = ""
    datFecha = 0
   
End Function

Public Function ImporteDeduccionPersonal(Deduccion As String, FechaTope As Date) As Double

    Dim SQL As String
    
    SQL = "Select " & Deduccion & " From DEDUCCIONES4TACATEGORIA Where FECHA <= #" & Format(FechaTope, "MM/DD/YYYY") & "# Order by FECHA Desc"
    Set rstBuscarSlave = New ADODB.Recordset
    rstBuscarSlave.Open SQL, dbSlave, adOpenDynamic, adLockOptimistic
    ImporteDeduccionPersonal = rstBuscarSlave.Fields(0) / 12
    rstBuscarSlave.Close
    Set rstBuscarSlave = Nothing
    
    SQL = ""
    
End Function

Public Function ImporteDeduccionGeneral(PuestoLaboral As String, Deduccion As String, _
FechaTope As Date, Optional GananciaNeta As Double, Optional SeguroVidaEnRecibo As Double = 0) As Double

    Dim SQL As String
    Dim dblLimiteDeduccion As Double
    Dim strMesBarraAno As String
    Dim strCL As String
    
    SQL = "Select " & Deduccion & " From IMPORTEDEDUCCIONESGENERALES Where PUESTOLABORAL = '" & PuestoLaboral & "' " _
    & "And FECHA <= # " & Format(FechaTope, "MM/DD/YYYY") & " # Order by FECHA Desc"
    Set rstBuscarSlave = New ADODB.Recordset
    rstBuscarSlave.Open SQL, dbSlave, adOpenDynamic, adLockOptimistic
    If rstBuscarSlave.BOF = False Then
        ImporteDeduccionGeneral = rstBuscarSlave.Fields(0) / 12
    End If
    rstBuscarSlave.Close
    Set rstBuscarSlave = Nothing
    
    Select Case Deduccion
    Case Is = "SERVICIODOMESTICO"
        SQL = "Select SERVICIODOMESTICO From DEDUCCIONES4TACATEGORIA Where FECHA <= #" & Format(FechaTope, "MM/DD/YYYY") & "# Order by FECHA Desc"
        Set rstBuscarSlave = New ADODB.Recordset
        rstBuscarSlave.Open SQL, dbSlave, adOpenDynamic, adLockOptimistic
        dblLimiteDeduccion = rstBuscarSlave.Fields(0) / 12
        rstBuscarSlave.Close
        If dblLimiteDeduccion < ImporteDeduccionGeneral Then
            ImporteDeduccionGeneral = dblLimiteDeduccion
        End If
    Case Is = "SEGURODEVIDA"
        If ImporteDeduccionGeneral = 0 Then 'NO ME CONVENCE
            strMesBarraAno = Left(Format(FechaTope, "MM/DD/YYYY"), 2) & "/" & Right(Format(FechaTope, "MM/DD/YYYY"), 4)
            strCL = BuscarCodigoLiquidacion(strMesBarraAno, True)
            ImporteDeduccionGeneral = DescuentoSeguroDeVidaOptativoAcumulado(PuestoLaboral, strCL)
            If ImporteDeduccionGeneral = 0 Then
                ImporteDeduccionGeneral = -ImporteRegistradoAcumuladoDeduccionEspecifica("SeguroDeVidaOptativo", _
                PuestoLaboral, strMesBarraAno, strCL) / Left(strMesBarraAno, 2)
            Else
                ImporteDeduccionGeneral = 0
            End If
        End If
        ImporteDeduccionGeneral = ImporteDeduccionGeneral + SeguroVidaEnRecibo
        SQL = "Select SEGURODEVIDA From DEDUCCIONES4TACATEGORIA Where FECHA <= #" & Format(FechaTope, "MM/DD/YYYY") & "# Order by FECHA Desc"
        Set rstBuscarSlave = New ADODB.Recordset
        rstBuscarSlave.Open SQL, dbSlave, adOpenDynamic, adLockOptimistic
        dblLimiteDeduccion = rstBuscarSlave.Fields(0) / 12
        rstBuscarSlave.Close
        If dblLimiteDeduccion < ImporteDeduccionGeneral Then
            ImporteDeduccionGeneral = dblLimiteDeduccion
        End If
    Case Is = "CUOTAMEDICOASISTENCIAL"
        SQL = "Select CUOTAMEDICOASISTENCIAL From DEDUCCIONES4TACATEGORIA Where FECHA <= #" & Format(FechaTope, "MM/DD/YYYY") & "# Order by FECHA Desc"
        Set rstBuscarSlave = New ADODB.Recordset
        rstBuscarSlave.Open SQL, dbSlave, adOpenDynamic, adLockOptimistic
        dblLimiteDeduccion = GananciaNeta * rstBuscarSlave.Fields(0)
        ImporteDeduccionGeneral = ImporteDeduccionGeneral * Month(FechaTope)
        rstBuscarSlave.Close
        If dblLimiteDeduccion < ImporteDeduccionGeneral Then
            ImporteDeduccionGeneral = dblLimiteDeduccion
        End If
    Case Is = "DONACIONES"
        SQL = "Select DONACIONES From DEDUCCIONES4TACATEGORIA Where FECHA <= #" & Format(FechaTope, "MM/DD/YYYY") & "# Order by FECHA Desc"
        Set rstBuscarSlave = New ADODB.Recordset
        rstBuscarSlave.Open SQL, dbSlave, adOpenDynamic, adLockOptimistic
        dblLimiteDeduccion = GananciaNeta * rstBuscarSlave.Fields(0)
        ImporteDeduccionGeneral = ImporteDeduccionGeneral * Month(FechaTope)
        rstBuscarSlave.Close
        If dblLimiteDeduccion < ImporteDeduccionGeneral Then
            ImporteDeduccionGeneral = dblLimiteDeduccion
        End If
    Case Is = "HONORARIOSMEDICOS"
        SQL = "Select HONORARIOSMEDICOS From DEDUCCIONES4TACATEGORIA Where FECHA <= #" & Format(FechaTope, "MM/DD/YYYY") & "# Order by FECHA Desc"
        Set rstBuscarSlave = New ADODB.Recordset
        rstBuscarSlave.Open SQL, dbSlave, adOpenDynamic, adLockOptimistic
        dblLimiteDeduccion = GananciaNeta * rstBuscarSlave.Fields(0)
        ImporteDeduccionGeneral = ImporteDeduccionGeneral * 12 '* 0.4 (EL SIRADIG YA AJUSTA POR 40%)
        rstBuscarSlave.Close
        If dblLimiteDeduccion < ImporteDeduccionGeneral Then
            ImporteDeduccionGeneral = dblLimiteDeduccion
        End If
    Case Is = "ALQUILERES"
        SQL = "Select ALQUILERES From DEDUCCIONES4TACATEGORIA Where FECHA <= #" & Format(FechaTope, "MM/DD/YYYY") & "# Order by FECHA Desc"
        Set rstBuscarSlave = New ADODB.Recordset
        rstBuscarSlave.Open SQL, dbSlave, adOpenDynamic, adLockOptimistic
        dblLimiteDeduccion = rstBuscarSlave.Fields(0) / 12
        rstBuscarSlave.Close
        ImporteDeduccionGeneral = ImporteDeduccionGeneral '* 0.4 (EL SIRADIG YA AJUSTA POR 40%)
        If dblLimiteDeduccion < ImporteDeduccionGeneral Then
            ImporteDeduccionGeneral = dblLimiteDeduccion
        End If
    End Select
    
    Set rstBuscarSlave = Nothing
    SQL = ""
    dblLimiteDeduccion = 0
    
End Function

Public Sub LiquidacionPrueba(CodigoBase As String, CodigoDestino As String, _
ConceptoLiquidacion As String, Optional ImporteFijo As Variant = "NaN", _
Optional Alicuota As Variant = "NaN")

    Dim SQL As String
    
    SQL = "Select * From LIQUIDACIONSUELDOS Where CODIGOLIQUIDACION = " & "'" & CodigoBase & "'"
    If SQLNoMatch(SQL) = True Then
        MsgBox "La liquidación que desea copiar debe contener registros", vbCritical + vbOKOnly, "LIQUIDACIÓN DE ORIGEN VACÍA"
        LiquidacionPruebaSISPER.cmbCodigoLiquidacionBase.SetFocus
        Exit Sub
    End If
    SQL = "Select * From LIQUIDACIONSUELDOS Where CODIGOLIQUIDACION = " & "'" & CodigoDestino & "'"
    If SQLNoMatch(SQL) = False Then
        MsgBox "Proceda a borrar los registros de la liquidación destino antes de incorporar datos a la misma", vbCritical + vbOKOnly, "LIQUIDACIÓN DESTINO LLENA"
        LiquidacionPruebaSISPER.cmbCodigoLiquidacionDestino.SetFocus
        Exit Sub
    End If
    
    Select Case ConceptoLiquidacion
    
    Case "0150" 'SAC
        Call InfoGeneral("- INICIANDO LIQUIDACIÓN PRUEBA -", LiquidacionPruebaSISPER)
        Call InfoGeneral("", LiquidacionPruebaSISPER)
        'Buscamos la mitad del Haber Óptimo de la liquidación base
        SQL = "Select ((IMPORTE)/2) as SAC, PUESTOLABORAL from LIQUIDACIONSUELDOS" _
        & " Where CODIGOLIQUIDACION = '" & CodigoBase _
        & "' And CODIGOCONCEPTO = '0115'"
        If SQLNoMatch(SQL) = True Then
            Call InfoGeneral("La liquidación base no posee Haber Óptimo", LiquidacionPruebaSISPER)
        Else
            'Copiamos la mitad del Haber Óptimo de la liquidación base a la liquidación destino como SAC (0150)
            SQL = "Insert Into LIQUIDACIONSUELDOS (CodigoLiquidacion, PuestoLaboral, CodigoConcepto, Importe)" & _
            " Select '" & CodigoDestino & "', PuestoLaboral, '0150', ((Importe)/2)" & _
            " From LIQUIDACIONSUELDOS Where CodigoLiquidacion = " & "'" & CodigoBase & "' And CODIGOCONCEPTO = '0115'"
            dbSlave.BeginTrans
            dbSlave.Execute SQL
            dbSlave.CommitTrans
            'Copiamos el concepto SAC (0150) de la liquidación destino al concepto Total Bruto (9998) de la misma liquidación
            SQL = "Insert Into LIQUIDACIONSUELDOS (CodigoLiquidacion, PuestoLaboral, CodigoConcepto, Importe)" & _
            " Select '" & CodigoDestino & "', PuestoLaboral, '9998', Importe" & _
            " From LIQUIDACIONSUELDOS Where CodigoLiquidacion = " & "'" & CodigoDestino & "' And CODIGOCONCEPTO = '0150'"
            dbSlave.BeginTrans
            dbSlave.Execute SQL
            dbSlave.CommitTrans
            'Calculamos la Jubilacion Personal a partir del Total Bruto
            SQL = "Insert Into LIQUIDACIONSUELDOS (CodigoLiquidacion, PuestoLaboral, CodigoConcepto, Importe)" & _
            " Select '" & CodigoDestino & "', PuestoLaboral, '0208', Format((Importe * 0.185), '#.00')" & _
            " From LIQUIDACIONSUELDOS Where CodigoLiquidacion = " & "'" & CodigoDestino & "' And CODIGOCONCEPTO = '9998'"
            dbSlave.BeginTrans
            dbSlave.Execute SQL
            dbSlave.CommitTrans
            'Calculamos la Jubilacion Estatal a partir del Total Bruto
            SQL = "Insert Into LIQUIDACIONSUELDOS (CodigoLiquidacion, PuestoLaboral, CodigoConcepto, Importe)" & _
            " Select '" & CodigoDestino & "', PuestoLaboral, '0209', Format((Importe * 0.185), '#.00')" & _
            " From LIQUIDACIONSUELDOS Where CodigoLiquidacion = " & "'" & CodigoDestino & "' And CODIGOCONCEPTO = '9998'"
            dbSlave.BeginTrans
            dbSlave.Execute SQL
            dbSlave.CommitTrans
            'Calculamos la Obra Social Personal a partir del Total Bruto
            SQL = "Insert Into LIQUIDACIONSUELDOS (CodigoLiquidacion, PuestoLaboral, CodigoConcepto, Importe)" & _
            " Select '" & CodigoDestino & "', PuestoLaboral, '0212', Format((Importe * 0.05), '#.00')" & _
            " From LIQUIDACIONSUELDOS Where CodigoLiquidacion = " & "'" & CodigoDestino & "' And CODIGOCONCEPTO = '9998'"
            dbSlave.BeginTrans
            dbSlave.Execute SQL
            dbSlave.CommitTrans
            'Calculamos la Obra Social Estatal a partir del Total Bruto
            SQL = "Insert Into LIQUIDACIONSUELDOS (CodigoLiquidacion, PuestoLaboral, CodigoConcepto, Importe)" & _
            " Select '" & CodigoDestino & "', PuestoLaboral, '0213', Format((Importe * 0.04), '#.00')" & _
            " From LIQUIDACIONSUELDOS Where CodigoLiquidacion = " & "'" & CodigoDestino & "' And CODIGOCONCEPTO = '9998'"
            dbSlave.BeginTrans
            dbSlave.Execute SQL
            dbSlave.CommitTrans
            'Calculamos la ART a partir del Haber Óptimo
            SQL = "Insert Into LIQUIDACIONSUELDOS (CodigoLiquidacion, PuestoLaboral, CodigoConcepto, Importe)" & _
            " Select '" & CodigoDestino & "', PuestoLaboral, '0381', Format((Importe * 0.0173 + 0.6), '#.00')" & _
            " From LIQUIDACIONSUELDOS Where CodigoLiquidacion = " & "'" & CodigoDestino & "' And CODIGOCONCEPTO = '0115'"
            dbSlave.BeginTrans
            dbSlave.Execute SQL
            dbSlave.CommitTrans
            'Verificamos si existe el descuento cuota sindical en la liquidación base y lo copiamos en la de destino
            SQL = "Select * from LIQUIDACIONSUELDOS" _
            & " Where (CODIGOLIQUIDACION = '" & CodigoBase _
            & "' And CODIGOCONCEPTO = '0219')"
            If SQLNoMatch(SQL) = False Then
                'Si existe descuento Cuota Sindical, copiamos la mitad del importe en la liquidación destino con el código 0219
                SQL = "Insert Into LIQUIDACIONSUELDOS (CodigoLiquidacion, PuestoLaboral, CodigoConcepto, Importe)" & _
                " Select '" & CodigoDestino & "', PuestoLaboral, '0219', (Importe / 2)" & _
                " From LIQUIDACIONSUELDOS Where CODIGOLIQUIDACION = '" & CodigoBase & "' And CODIGOCONCEPTO = '0219'"
                dbSlave.BeginTrans
                dbSlave.Execute SQL
                dbSlave.CommitTrans
            End If
            SQL = "Select * from LIQUIDACIONSUELDOS" _
            & " Where (CODIGOLIQUIDACION = '" & CodigoBase _
            & "' And CODIGOCONCEPTO = '0227')"
            If SQLNoMatch(SQL) = False Then
                'Si existe descuento Cuota Sindical, copiamos la mitad del importe en la liquidación destino con el código 0227
                SQL = "Insert Into LIQUIDACIONSUELDOS (CodigoLiquidacion, PuestoLaboral, CodigoConcepto, Importe)" & _
                " Select '" & CodigoDestino & "', PuestoLaboral, '0227', (Importe / 2)" & _
                " From LIQUIDACIONSUELDOS Where CODIGOLIQUIDACION = '" & CodigoBase & "' And CODIGOCONCEPTO = '0227'"
                dbSlave.BeginTrans
                dbSlave.Execute SQL
                dbSlave.CommitTrans
            End If
        End If
        Call InfoGeneral("", LiquidacionPruebaSISPER)
        Call InfoGeneral("- LIQUIDACIÓN FINALIZADA -", LiquidacionPruebaSISPER)
    
    Case "0005" 'Antigüedad
        Call InfoGeneral("- INICIANDO LIQUIDACIÓN PRUEBA -", LiquidacionPruebaSISPER)
        Call InfoGeneral("", LiquidacionPruebaSISPER)
        'Buscamos el Haber Óptimo de la liquidación base
        SQL = "Select IMPORTE, PUESTOLABORAL from LIQUIDACIONSUELDOS" _
        & " Where CODIGOLIQUIDACION = '" & CodigoBase _
        & "' And CODIGOCONCEPTO = '0115'"
        If SQLNoMatch(SQL) = True Then
            Call InfoGeneral("La liquidación base no posee Haber Óptimo", LiquidacionPruebaSISPER)
        Else
            'Primero, copiamos de la liquidación base todos los conceptos remunerativos que no van a sufrir alteración
            'Copiamos el concepto Básico (0001) de la liquidación base a la liquidación destino
            SQL = "Insert Into LIQUIDACIONSUELDOS (CodigoLiquidacion, PuestoLaboral, CodigoConcepto, Importe)" & _
            " Select '" & CodigoDestino & "', PuestoLaboral, '0001', Importe" & _
            " From LIQUIDACIONSUELDOS Where CodigoLiquidacion = " & "'" & CodigoBase & "' And CODIGOCONCEPTO = '0001'"
            dbSlave.BeginTrans
            dbSlave.Execute SQL
            dbSlave.CommitTrans
            'Copiamos el concepto Título (0008) de la liquidación base a la liquidación destino
            SQL = "Insert Into LIQUIDACIONSUELDOS (CodigoLiquidacion, PuestoLaboral, CodigoConcepto, Importe)" & _
            " Select '" & CodigoDestino & "', PuestoLaboral, '0008', Importe" & _
            " From LIQUIDACIONSUELDOS Where CodigoLiquidacion = " & "'" & CodigoBase & "' And CODIGOCONCEPTO = '0008'"
            dbSlave.BeginTrans
            dbSlave.Execute SQL
            dbSlave.CommitTrans
            'Copiamos el concepto Gastos de Representación (0028) de la liquidación base a la liquidación destino
            SQL = "Insert Into LIQUIDACIONSUELDOS (CodigoLiquidacion, PuestoLaboral, CodigoConcepto, Importe)" & _
            " Select '" & CodigoDestino & "', PuestoLaboral, '0028', Importe" & _
            " From LIQUIDACIONSUELDOS Where CodigoLiquidacion = " & "'" & CodigoBase & "' And CODIGOCONCEPTO = '0028'"
            dbSlave.BeginTrans
            dbSlave.Execute SQL
            dbSlave.CommitTrans
            'Copiamos el concepto Dedicación Exclusiva (0064) de la liquidación base a la liquidación destino
            SQL = "Insert Into LIQUIDACIONSUELDOS (CodigoLiquidacion, PuestoLaboral, CodigoConcepto, Importe)" & _
            " Select '" & CodigoDestino & "', PuestoLaboral, '0064', Importe" & _
            " From LIQUIDACIONSUELDOS Where CodigoLiquidacion = " & "'" & CodigoBase & "' And CODIGOCONCEPTO = '0064'"
            dbSlave.BeginTrans
            dbSlave.Execute SQL
            dbSlave.CommitTrans
            'Copiamos el concepto AD.REM.DEC.2232/10-415/13 (0134) de la liquidación base a la liquidación destino
            SQL = "Insert Into LIQUIDACIONSUELDOS (CodigoLiquidacion, PuestoLaboral, CodigoConcepto, Importe)" & _
            " Select '" & CodigoDestino & "', PuestoLaboral, '0134', Importe" & _
            " From LIQUIDACIONSUELDOS Where CodigoLiquidacion = " & "'" & CodigoBase & "' And CODIGOCONCEPTO = '0134'"
            dbSlave.BeginTrans
            dbSlave.Execute SQL
            dbSlave.CommitTrans
            'Copiamos el concepto ADICIONAL DEC 1619/05 C/A (0169) de la liquidación base a la liquidación destino
            SQL = "Insert Into LIQUIDACIONSUELDOS (CodigoLiquidacion, PuestoLaboral, CodigoConcepto, Importe)" & _
            " Select '" & CodigoDestino & "', PuestoLaboral, '0169', Importe" & _
            " From LIQUIDACIONSUELDOS Where CodigoLiquidacion = " & "'" & CodigoBase & "' And CODIGOCONCEPTO = '0169'"
            dbSlave.BeginTrans
            dbSlave.Execute SQL
            dbSlave.CommitTrans
            'Copiamos el concepto Plus Remunerativo (0603) de la liquidación base a la liquidación destino
            SQL = "Insert Into LIQUIDACIONSUELDOS (CodigoLiquidacion, PuestoLaboral, CodigoConcepto, Importe)" & _
            " Select '" & CodigoDestino & "', PuestoLaboral, '0603', Importe" & _
            " From LIQUIDACIONSUELDOS Where CodigoLiquidacion = " & "'" & CodigoBase & "' And CODIGOCONCEPTO = '0603'"
            dbSlave.BeginTrans
            dbSlave.Execute SQL
            dbSlave.CommitTrans
            
            'Segundo, creamos una tabla adicional para calcular los conceptos que sufren cambio
            SQL = "Create Table CALCULOSAUXILIARES (" _
            & "PUESTOLABORAL Text (6), BASICO Currency, " _
            & "ANTIGUEDAD Currency, ANOS Byte, TITULO Currency, " _
            & "COMPENSACIONFUNCIONAL Currency, FUNCIONESPECIFICA Currency, " _
            & "BONIFICACIONFONAVI Currency, COEFICIENTEFONAVI Double, " _
            & "PRESENTISMO Currency)"
            dbSlave.BeginTrans
            dbSlave.Execute SQL
            dbSlave.CommitTrans
            'Copiamos el concepto Básico (0001) de la liquidación base a la tabla de cálculos auxiliares
            SQL = "Insert Into CALCULOSAUXILIARES (PUESTOLABORAL, BASICO, ANTIGUEDAD, ANOS, TITULO, COMPENSACIONFUNCIONAL, FUNCIONESPECIFICA, BONIFICACIONFONAVI, COEFICIENTEFONAVI, PRESENTISMO)" & _
            " Select PUESTOLABORAL, Importe, 0, 0, 0, 0, 0, 0, 0, 0" & _
            " From LIQUIDACIONSUELDOS Where CodigoLiquidacion = " & "'" & CodigoBase & "' And CODIGOCONCEPTO = '0001'"
            dbSlave.BeginTrans
            dbSlave.Execute SQL
            dbSlave.CommitTrans
            'Modificamos el campo TITULO de la Tabla de Calculos Auxiliares para incorporar el Titulo de la tabla base
            SQL = "Update CALCULOSAUXILIARES Inner Join " _
            & "(Select PUESTOLABORAL, IMPORTE From LIQUIDACIONSUELDOS " _
            & "Where CodigoLiquidacion = " & "'" & CodigoBase _
            & "' And CODIGOCONCEPTO = '0008') As LIQUIDACIONSUELDOS " _
            & "On CALCULOSAUXILIARES.PUESTOLABORAL = LIQUIDACIONSUELDOS.PUESTOLABORAL " _
            & "Set CALCULOSAUXILIARES.TITULO = LIQUIDACIONSUELDOS.IMPORTE"
            dbSlave.BeginTrans
            dbSlave.Execute SQL
            dbSlave.CommitTrans
            'Modificamos el campo Bonificación FONAVI de la Tabla de Calculos Auxiliares para incorporar la Bonificación FONAVI de la tabla base
            SQL = "Update CALCULOSAUXILIARES Inner Join " _
            & "(Select PUESTOLABORAL, IMPORTE From LIQUIDACIONSUELDOS " _
            & "Where CodigoLiquidacion = " & "'" & CodigoBase _
            & "' And CODIGOCONCEPTO = '0108') As LIQUIDACIONSUELDOS " _
            & "On CALCULOSAUXILIARES.PUESTOLABORAL = LIQUIDACIONSUELDOS.PUESTOLABORAL " _
            & "Set CALCULOSAUXILIARES.BONIFICACIONFONAVI = LIQUIDACIONSUELDOS.IMPORTE"
            dbSlave.BeginTrans
            dbSlave.Execute SQL
            dbSlave.CommitTrans
            'Modificamos el campo Antigüedad de la Tabla de Calculos Auxiliares para incorporar la antiguedad de la tabla base
            SQL = "Update CALCULOSAUXILIARES Inner Join " _
            & "(Select PUESTOLABORAL, IMPORTE From LIQUIDACIONSUELDOS " _
            & "Where CodigoLiquidacion = " & "'" & CodigoBase _
            & "' And CODIGOCONCEPTO = '0005') As LIQUIDACIONSUELDOS " _
            & "On CALCULOSAUXILIARES.PUESTOLABORAL = LIQUIDACIONSUELDOS.PUESTOLABORAL " _
            & "Set CALCULOSAUXILIARES.ANTIGUEDAD = LIQUIDACIONSUELDOS.IMPORTE"
            dbSlave.BeginTrans
            dbSlave.Execute SQL
            dbSlave.CommitTrans
            'Calculamos el Coeficiente de la Bonificación FONAVI
            SQL = "Update CALCULOSAUXILIARES " _
            & "Set COEFICIENTEFONAVI = Format(BONIFICACIONFONAVI / (BASICO + ANTIGUEDAD + TITULO), '#.00') " _
            & "Where BONIFICACIONFONAVI > 0"
            dbSlave.BeginTrans
            dbSlave.Execute SQL
            dbSlave.CommitTrans
            'Calculamos la nueva antiguedad en años (agregamos 1 año)
            SQL = "Update CALCULOSAUXILIARES " _
            & "Set ANOS = Format((ANTIGUEDAD / BASICO / 0.02), '#') + 1"
            dbSlave.BeginTrans
            dbSlave.Execute SQL
            dbSlave.CommitTrans
            'Recalculamos la antiguedad con los datos obtenidos
            SQL = "Update CALCULOSAUXILIARES " _
            & "Set ANTIGUEDAD = Format((ANOS * BASICO * 0.02), '#.00')"
            dbSlave.BeginTrans
            dbSlave.Execute SQL
            dbSlave.CommitTrans
            'Recalculamos la Bonificación FONAVI en funcion de la nueva antiguedad
            SQL = "Update CALCULOSAUXILIARES " _
            & "Set BONIFICACIONFONAVI = Format((BASICO + ANTIGUEDAD + TITULO) * COEFICIENTEFONAVI, '#.00') " _
            & "Where BONIFICACIONFONAVI > 0"
            dbSlave.BeginTrans
            dbSlave.Execute SQL
            dbSlave.CommitTrans
            'Modificamos el campo Compensacion Funcional de la Tabla de Calculos Auxiliares para incorporar la Compensacion Funcional de la tabla base
            SQL = "Update CALCULOSAUXILIARES Inner Join " _
            & "(Select PUESTOLABORAL, IMPORTE From LIQUIDACIONSUELDOS " _
            & "Where CodigoLiquidacion = " & "'" & CodigoBase _
            & "' And CODIGOCONCEPTO = '0002') As LIQUIDACIONSUELDOS " _
            & "On CALCULOSAUXILIARES.PUESTOLABORAL = LIQUIDACIONSUELDOS.PUESTOLABORAL " _
            & "Set CALCULOSAUXILIARES.COMPENSACIONFUNCIONAL = LIQUIDACIONSUELDOS.IMPORTE"
            dbSlave.BeginTrans
            dbSlave.Execute SQL
            dbSlave.CommitTrans
            'Recalculamos la Compensación Funcional
            SQL = "Update CALCULOSAUXILIARES " _
            & "Set COMPENSACIONFUNCIONAL = Format(((BASICO + TITULO + ANTIGUEDAD) * 0.51), '#.00') " _
            & "Where COMPENSACIONFUNCIONAL > 0"
            dbSlave.BeginTrans
            dbSlave.Execute SQL
            dbSlave.CommitTrans
            'Modificamos el campo Funcion Especifica de la Tabla de Calculos Auxiliares para incorporar la Funcion Especifica de la tabla base
            SQL = "Update CALCULOSAUXILIARES Inner Join " _
            & "(Select PUESTOLABORAL, IMPORTE From LIQUIDACIONSUELDOS " _
            & "Where CodigoLiquidacion = " & "'" & CodigoBase _
            & "' And CODIGOCONCEPTO = '0065') As LIQUIDACIONSUELDOS " _
            & "On CALCULOSAUXILIARES.PUESTOLABORAL = LIQUIDACIONSUELDOS.PUESTOLABORAL " _
            & "Set CALCULOSAUXILIARES.FUNCIONESPECIFICA = LIQUIDACIONSUELDOS.IMPORTE"
            dbSlave.BeginTrans
            dbSlave.Execute SQL
            dbSlave.CommitTrans
            'Recalculamos la Funcion Especifica
            SQL = "Update CALCULOSAUXILIARES " _
            & "Set FUNCIONESPECIFICA = Format(((BASICO + TITULO + ANTIGUEDAD + COMPENSACIONFUNCIONAL) * 1.10), '#.00') " _
            & "Where FUNCIONESPECIFICA > 0"
            dbSlave.BeginTrans
            dbSlave.Execute SQL
            dbSlave.CommitTrans
            'Modificamos el campo Presentismo de la Tabla de Calculos Auxiliares para incorporar el Presentismo de la tabla base
            SQL = "Update CALCULOSAUXILIARES Inner Join " _
            & "(Select PUESTOLABORAL, IMPORTE From LIQUIDACIONSUELDOS " _
            & "Where CodigoLiquidacion = " & "'" & CodigoBase _
            & "' And CODIGOCONCEPTO = '0054') As LIQUIDACIONSUELDOS " _
            & "On CALCULOSAUXILIARES.PUESTOLABORAL = LIQUIDACIONSUELDOS.PUESTOLABORAL " _
            & "Set CALCULOSAUXILIARES.PRESENTISMO = LIQUIDACIONSUELDOS.IMPORTE"
            dbSlave.BeginTrans
            dbSlave.Execute SQL
            dbSlave.CommitTrans
            'Recalculamos el Presentismo
            SQL = "Update CALCULOSAUXILIARES " _
            & "Set PRESENTISMO = Format(((BASICO + TITULO + ANTIGUEDAD + COMPENSACIONFUNCIONAL + BONIFICACIONFONAVI + FUNCIONESPECIFICA) * 0.15), '#.00') " _
            & "Where PRESENTISMO > 0"
            dbSlave.BeginTrans
            dbSlave.Execute SQL
            dbSlave.CommitTrans
            'Copiamos el concepto Antiguedad (0005) de la tabla Calculo Auxiliares a la liquidación destino
            SQL = "Insert Into LIQUIDACIONSUELDOS (CodigoLiquidacion, PuestoLaboral, CodigoConcepto, Importe)" & _
            " Select '" & CodigoDestino & "', PUESTOLABORAL, '0005', ANTIGUEDAD" & _
            " From CALCULOSAUXILIARES Where ANTIGUEDAD > 0"
            dbSlave.BeginTrans
            dbSlave.Execute SQL
            dbSlave.CommitTrans
            'Copiamos el concepto Compensación Funcional (0002) de la tabla Calculo Auxiliares a la liquidación destino
            SQL = "Insert Into LIQUIDACIONSUELDOS (CodigoLiquidacion, PuestoLaboral, CodigoConcepto, Importe)" & _
            " Select '" & CodigoDestino & "', PUESTOLABORAL, '0002', COMPENSACIONFUNCIONAL" & _
            " From CALCULOSAUXILIARES Where COMPENSACIONFUNCIONAL > 0"
            dbSlave.BeginTrans
            dbSlave.Execute SQL
            dbSlave.CommitTrans
            'Copiamos el concepto Función Especifica (0065) de la tabla Calculo Auxiliares a la liquidación destino
            SQL = "Insert Into LIQUIDACIONSUELDOS (CodigoLiquidacion, PuestoLaboral, CodigoConcepto, Importe)" & _
            " Select '" & CodigoDestino & "', PUESTOLABORAL, '0065', FUNCIONESPECIFICA" & _
            " From CALCULOSAUXILIARES Where FUNCIONESPECIFICA > 0"
            dbSlave.BeginTrans
            dbSlave.Execute SQL
            dbSlave.CommitTrans
            'Copiamos el concepto Bonificación FONAVI (0108) de la tabla Calculo Auxiliares a la liquidación destino
            SQL = "Insert Into LIQUIDACIONSUELDOS (CodigoLiquidacion, PuestoLaboral, CodigoConcepto, Importe)" & _
            " Select '" & CodigoDestino & "', PUESTOLABORAL, '0108', BONIFICACIONFONAVI" & _
            " From CALCULOSAUXILIARES Where BONIFICACIONFONAVI > 0"
            dbSlave.BeginTrans
            dbSlave.Execute SQL
            dbSlave.CommitTrans
            'Copiamos el concepto Presentismo (0054) de la tabla Calculo Auxiliares a la liquidación destino
            SQL = "Insert Into LIQUIDACIONSUELDOS (CodigoLiquidacion, PuestoLaboral, CodigoConcepto, Importe)" & _
            " Select '" & CodigoDestino & "', PUESTOLABORAL, '0054', PRESENTISMO" & _
            " From CALCULOSAUXILIARES Where PRESENTISMO > 0"
            dbSlave.BeginTrans
            dbSlave.Execute SQL
            dbSlave.CommitTrans
            
            'Tercero, borramos la tabla auxiliar
            SQL = "Drop Table CALCULOSAUXILIARES"
            dbSlave.BeginTrans
            dbSlave.Execute SQL
            dbSlave.CommitTrans
            
            'Cuarto, calculamos el nuevo Haber Óptimo (0115)
            SQL = "Insert Into LIQUIDACIONSUELDOS (CodigoLiquidacion, PuestoLaboral, CodigoConcepto, Importe)" & _
            " Select '" & CodigoDestino & "', PUESTOLABORAL, '0115', HABEROPTIMO" & _
            " From (Select PUESTOLABORAL, Sum(IMPORTE) As HABEROPTIMO" & _
            " From LIQUIDACIONSUELDOS Where CodigoLiquidacion = " & "'" & CodigoDestino _
            & "' Group By PUESTOLABORAL)"
            dbSlave.BeginTrans
            dbSlave.Execute SQL
            dbSlave.CommitTrans
            
            'Quinto, copiamos Salario Familiar y Discapacitados
            'Copiamos el concepto Salario Familiar (0003) de la liquidación base a la liquidación destino
            SQL = "Insert Into LIQUIDACIONSUELDOS (CodigoLiquidacion, PuestoLaboral, CodigoConcepto, Importe)" & _
            " Select '" & CodigoDestino & "', PuestoLaboral, '0003', Importe" & _
            " From LIQUIDACIONSUELDOS Where CodigoLiquidacion = " & "'" & CodigoBase & "' And CODIGOCONCEPTO = '0003'"
            dbSlave.BeginTrans
            dbSlave.Execute SQL
            dbSlave.CommitTrans
            'Copiamos el concepto Discapacitados (0058) de la liquidación base a la liquidación destino
            SQL = "Insert Into LIQUIDACIONSUELDOS (CodigoLiquidacion, PuestoLaboral, CodigoConcepto, Importe)" & _
            " Select '" & CodigoDestino & "', PuestoLaboral, '0058', Importe" & _
            " From LIQUIDACIONSUELDOS Where CodigoLiquidacion = " & "'" & CodigoBase & "' And CODIGOCONCEPTO = '0058'"
            dbSlave.BeginTrans
            dbSlave.Execute SQL
            dbSlave.CommitTrans
            
            'Sexto, calculamos el nuevo Total Bruto (9998)
            SQL = "Insert Into LIQUIDACIONSUELDOS (CodigoLiquidacion, PuestoLaboral, CodigoConcepto, Importe)" & _
            " Select '" & CodigoDestino & "', PUESTOLABORAL, '9998', TOTALBRUTO" & _
            " From (Select PUESTOLABORAL, Sum(IMPORTE) As TOTALBRUTO" & _
            " From LIQUIDACIONSUELDOS Where CodigoLiquidacion = " & "'" & CodigoDestino _
            & "' And CODIGOCONCEPTO <> '0115' Group By PUESTOLABORAL)"
            dbSlave.BeginTrans
            dbSlave.Execute SQL
            dbSlave.CommitTrans
            
            'Septimo, calculamos los descuentos
            'Calculamos la Jubilacion Personal a partir del Haber Óptimo
            SQL = "Insert Into LIQUIDACIONSUELDOS (CodigoLiquidacion, PuestoLaboral, CodigoConcepto, Importe)" & _
            " Select '" & CodigoDestino & "', PuestoLaboral, '0208', Format((Importe * 0.185), '#.00')" & _
            " From LIQUIDACIONSUELDOS Where CodigoLiquidacion = " & "'" & CodigoDestino & "' And CODIGOCONCEPTO = '0115'"
            dbSlave.BeginTrans
            dbSlave.Execute SQL
            dbSlave.CommitTrans
            'Calculamos la Jubilacion Estatal a partir del Haber Óptimo
            SQL = "Insert Into LIQUIDACIONSUELDOS (CodigoLiquidacion, PuestoLaboral, CodigoConcepto, Importe)" & _
            " Select '" & CodigoDestino & "', PuestoLaboral, '0209', Format((Importe * 0.185), '#.00')" & _
            " From LIQUIDACIONSUELDOS Where CodigoLiquidacion = " & "'" & CodigoDestino & "' And CODIGOCONCEPTO = '0115'"
            dbSlave.BeginTrans
            dbSlave.Execute SQL
            dbSlave.CommitTrans
            'Calculamos la Obra Social Personal a partir del Haber Óptimo
            SQL = "Insert Into LIQUIDACIONSUELDOS (CodigoLiquidacion, PuestoLaboral, CodigoConcepto, Importe)" & _
            " Select '" & CodigoDestino & "', PuestoLaboral, '0212', Format((Importe * 0.05), '#.00')" & _
            " From LIQUIDACIONSUELDOS Where CodigoLiquidacion = " & "'" & CodigoDestino & "' And CODIGOCONCEPTO = '0115'"
            dbSlave.BeginTrans
            dbSlave.Execute SQL
            dbSlave.CommitTrans
            'Calculamos la Obra Social Estatal a partir del Haber Óptimo
            SQL = "Insert Into LIQUIDACIONSUELDOS (CodigoLiquidacion, PuestoLaboral, CodigoConcepto, Importe)" & _
            " Select '" & CodigoDestino & "', PuestoLaboral, '0213', Format((Importe * 0.04), '#.00')" & _
            " From LIQUIDACIONSUELDOS Where CodigoLiquidacion = " & "'" & CodigoDestino & "' And CODIGOCONCEPTO = '0115'"
            dbSlave.BeginTrans
            dbSlave.Execute SQL
            dbSlave.CommitTrans
            'Calculamos la ART a partir del Haber Óptimo
            SQL = "Insert Into LIQUIDACIONSUELDOS (CodigoLiquidacion, PuestoLaboral, CodigoConcepto, Importe)" & _
            " Select '" & CodigoDestino & "', PuestoLaboral, '0381', Format((Importe * 0.0173 + 0.6), '#.00')" & _
            " From LIQUIDACIONSUELDOS Where CodigoLiquidacion = " & "'" & CodigoDestino & "' And CODIGOCONCEPTO = '0115'"
            dbSlave.BeginTrans
            dbSlave.Execute SQL
            dbSlave.CommitTrans
            'Copiamos el concepto Cuota Sindical (0219 y 0227) de la liquidación base a la liquidación destino (0219) - (VERIFICAR PORQUE COPIAMOS EL MISMO IMPORTE)
            SQL = "Insert Into LIQUIDACIONSUELDOS (CodigoLiquidacion, PuestoLaboral, CodigoConcepto, Importe)" & _
            " Select '" & CodigoDestino & "', PUESTOLABORAL, '0219', CUOTASINDICAL" & _
            " From (Select PUESTOLABORAL, Sum(IMPORTE) As CUOTASINDICAL" & _
            " From LIQUIDACIONSUELDOS Where" & _
            " ((CodigoLiquidacion = " & "'" & CodigoBase & "' And CODIGOCONCEPTO = '0219') Or" & _
            " (CodigoLiquidacion = " & "'" & CodigoBase & "' And CODIGOCONCEPTO = '0227')) Group By PUESTOLABORAL)"
            dbSlave.BeginTrans
            dbSlave.Execute SQL
            dbSlave.CommitTrans
            'Copiamos el concepto Seguro de Vida Optativo (Varios) de la liquidación base a la liquidación destino (0317) - (VERIFICAR PORQUE COPIAMOS EL MISMO IMPORTE)
            SQL = "Insert Into LIQUIDACIONSUELDOS (CodigoLiquidacion, PuestoLaboral, CodigoConcepto, Importe)" & _
            " Select '" & CodigoDestino & "', PUESTOLABORAL, '0317', SEGUROOPTATIVO" & _
            " From (Select PUESTOLABORAL, Sum(IMPORTE) As SEGUROOPTATIVO" & _
            " From LIQUIDACIONSUELDOS Where" & _
            " ((CodigoLiquidacion = " & "'" & CodigoBase & "' And CODIGOCONCEPTO = '0317') Or" & _
            " (CodigoLiquidacion = " & "'" & CodigoBase & "' And CODIGOCONCEPTO = '0361') Or" & _
            " (CodigoLiquidacion = " & "'" & CodigoBase & "' And CODIGOCONCEPTO = '0367') Or" & _
            " (CodigoLiquidacion = " & "'" & CodigoBase & "' And CODIGOCONCEPTO = '0370') Or" & _
            " (CodigoLiquidacion = " & "'" & CodigoBase & "' And CODIGOCONCEPTO = '0373') Or" & _
            " (CodigoLiquidacion = " & "'" & CodigoBase & "' And CODIGOCONCEPTO = '0374')) Group By PUESTOLABORAL)"
            dbSlave.BeginTrans
            dbSlave.Execute SQL
            dbSlave.CommitTrans
            'Copiamos el concepto Seguro de Vida Obligatorio (0360, 0366, 0369) de la liquidación base a la liquidación destino (0360)
            SQL = "Insert Into LIQUIDACIONSUELDOS (CodigoLiquidacion, PuestoLaboral, CodigoConcepto, Importe)" & _
            " Select '" & CodigoDestino & "', PUESTOLABORAL, '0360', SEGUROOBLIGATORIO" & _
            " From (Select PUESTOLABORAL, Sum(IMPORTE) As SEGUROOBLIGATORIO" & _
            " From LIQUIDACIONSUELDOS Where" & _
            " ((CodigoLiquidacion = " & "'" & CodigoBase & "' And CODIGOCONCEPTO = '0360') Or" & _
            " (CodigoLiquidacion = " & "'" & CodigoBase & "' And CODIGOCONCEPTO = '0366') Or" & _
            " (CodigoLiquidacion = " & "'" & CodigoBase & "' And CODIGOCONCEPTO = '0369')) Group By PUESTOLABORAL)"
            dbSlave.BeginTrans
            dbSlave.Execute SQL
            dbSlave.CommitTrans
        
            Call InfoGeneral("", LiquidacionPruebaSISPER)
            Call InfoGeneral("- LIQUIDACIÓN FINALIZADA -", LiquidacionPruebaSISPER)
        End If
        
    Case "0001" 'Basico
        Call InfoGeneral("- INICIANDO LIQUIDACIÓN PRUEBA -", LiquidacionPruebaSISPER)
        Call InfoGeneral("", LiquidacionPruebaSISPER)
        'Buscamos el Haber Óptimo de la liquidación base
        SQL = "Select IMPORTE, PUESTOLABORAL from LIQUIDACIONSUELDOS" _
        & " Where CODIGOLIQUIDACION = '" & CodigoBase _
        & "' And CODIGOCONCEPTO = '0115'"
        If SQLNoMatch(SQL) = True Then
            Call InfoGeneral("La liquidación base no posee Haber Óptimo", LiquidacionPruebaSISPER)
        Else
            'Primero, copiamos de la liquidación base todos los conceptos remunerativos que no van a sufrir alteración
            'Copiamos el concepto Gastos de Representación (0028) de la liquidación base a la liquidación destino
            SQL = "Insert Into LIQUIDACIONSUELDOS (CodigoLiquidacion, PuestoLaboral, CodigoConcepto, Importe)" & _
            " Select '" & CodigoDestino & "', PuestoLaboral, '0028', Importe" & _
            " From LIQUIDACIONSUELDOS Where CodigoLiquidacion = " & "'" & CodigoBase & "' And CODIGOCONCEPTO = '0028'"
            dbSlave.BeginTrans
            dbSlave.Execute SQL
            dbSlave.CommitTrans
            'Copiamos el concepto AD.REM.DEC.2232/10-415/13 (0134) de la liquidación base a la liquidación destino
            SQL = "Insert Into LIQUIDACIONSUELDOS (CodigoLiquidacion, PuestoLaboral, CodigoConcepto, Importe)" & _
            " Select '" & CodigoDestino & "', PuestoLaboral, '0134', Importe" & _
            " From LIQUIDACIONSUELDOS Where CodigoLiquidacion = " & "'" & CodigoBase & "' And CODIGOCONCEPTO = '0134'"
            dbSlave.BeginTrans
            dbSlave.Execute SQL
            dbSlave.CommitTrans
            'Copiamos el concepto ADICIONAL DEC 1619/05 C/A (0169) de la liquidación base a la liquidación destino
            SQL = "Insert Into LIQUIDACIONSUELDOS (CodigoLiquidacion, PuestoLaboral, CodigoConcepto, Importe)" & _
            " Select '" & CodigoDestino & "', PuestoLaboral, '0169', Importe" & _
            " From LIQUIDACIONSUELDOS Where CodigoLiquidacion = " & "'" & CodigoBase & "' And CODIGOCONCEPTO = '0169'"
            dbSlave.BeginTrans
            dbSlave.Execute SQL
            dbSlave.CommitTrans
            'Copiamos el concepto Plus Remunerativo (0603) de la liquidación base a la liquidación destino
            SQL = "Insert Into LIQUIDACIONSUELDOS (CodigoLiquidacion, PuestoLaboral, CodigoConcepto, Importe)" & _
            " Select '" & CodigoDestino & "', PuestoLaboral, '0603', Importe" & _
            " From LIQUIDACIONSUELDOS Where CodigoLiquidacion = " & "'" & CodigoBase & "' And CODIGOCONCEPTO = '0603'"
            dbSlave.BeginTrans
            dbSlave.Execute SQL
            dbSlave.CommitTrans
            
            'Segundo, creamos una tabla adicional para calcular los conceptos que sufren cambio
            SQL = "Create Table CALCULOSAUXILIARES (" _
            & "PUESTOLABORAL Text (6), BASICO Currency, " _
            & "ANTIGUEDAD Currency, ANOS Byte, TITULO Currency, " _
            & "COMPENSACIONFUNCIONAL Currency, FUNCIONESPECIFICA Currency, " _
            & "BONIFICACIONFONAVI Currency, COEFICIENTEFONAVI Double, " _
            & "PRESENTISMO Currency, COEFICIENTETITULO Double, " _
            & "DEDICACIONEXCLUSIVA Double)"
            dbSlave.BeginTrans
            dbSlave.Execute SQL
            dbSlave.CommitTrans
            'Copiamos el concepto Básico (0001) de la liquidación base a la tabla de cálculos auxiliares
            SQL = "Insert Into CALCULOSAUXILIARES (PUESTOLABORAL, BASICO, ANTIGUEDAD, ANOS, TITULO, COMPENSACIONFUNCIONAL, FUNCIONESPECIFICA, BONIFICACIONFONAVI, COEFICIENTEFONAVI, PRESENTISMO, COEFICIENTETITULO, DEDICACIONEXCLUSIVA)" & _
            " Select PUESTOLABORAL, Importe, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0" & _
            " From LIQUIDACIONSUELDOS Where CodigoLiquidacion = " & "'" & CodigoBase & "' And CODIGOCONCEPTO = '0001'"
            dbSlave.BeginTrans
            dbSlave.Execute SQL
            dbSlave.CommitTrans
            'Modificamos el campo TITULO (0008) de la Tabla de Calculos Auxiliares para incorporar el Titulo de la tabla base
            SQL = "Update CALCULOSAUXILIARES Inner Join " _
            & "(Select PUESTOLABORAL, IMPORTE From LIQUIDACIONSUELDOS " _
            & "Where CodigoLiquidacion = " & "'" & CodigoBase _
            & "' And CODIGOCONCEPTO = '0008') As LIQUIDACIONSUELDOS " _
            & "On CALCULOSAUXILIARES.PUESTOLABORAL = LIQUIDACIONSUELDOS.PUESTOLABORAL " _
            & "Set CALCULOSAUXILIARES.TITULO = LIQUIDACIONSUELDOS.IMPORTE"
            dbSlave.BeginTrans
            dbSlave.Execute SQL
            dbSlave.CommitTrans
            'Modificamos el campo Bonificación FONAVI de la Tabla de Calculos Auxiliares para incorporar la Bonificación FONAVI de la tabla base
            SQL = "Update CALCULOSAUXILIARES Inner Join " _
            & "(Select PUESTOLABORAL, IMPORTE From LIQUIDACIONSUELDOS " _
            & "Where CodigoLiquidacion = " & "'" & CodigoBase _
            & "' And CODIGOCONCEPTO = '0108') As LIQUIDACIONSUELDOS " _
            & "On CALCULOSAUXILIARES.PUESTOLABORAL = LIQUIDACIONSUELDOS.PUESTOLABORAL " _
            & "Set CALCULOSAUXILIARES.BONIFICACIONFONAVI = LIQUIDACIONSUELDOS.IMPORTE"
            dbSlave.BeginTrans
            dbSlave.Execute SQL
            dbSlave.CommitTrans
            'Modificamos el campo Antigüedad de la Tabla de Calculos Auxiliares para incorporar la antiguedad de la tabla base
            SQL = "Update CALCULOSAUXILIARES Inner Join " _
            & "(Select PUESTOLABORAL, IMPORTE From LIQUIDACIONSUELDOS " _
            & "Where CodigoLiquidacion = " & "'" & CodigoBase _
            & "' And CODIGOCONCEPTO = '0005') As LIQUIDACIONSUELDOS " _
            & "On CALCULOSAUXILIARES.PUESTOLABORAL = LIQUIDACIONSUELDOS.PUESTOLABORAL " _
            & "Set CALCULOSAUXILIARES.ANTIGUEDAD = LIQUIDACIONSUELDOS.IMPORTE"
            dbSlave.BeginTrans
            dbSlave.Execute SQL
            dbSlave.CommitTrans
            'Calculamos el Coeficiente de la Bonificación FONAVI
            SQL = "Update CALCULOSAUXILIARES " _
            & "Set COEFICIENTEFONAVI = Format(BONIFICACIONFONAVI / (BASICO + ANTIGUEDAD + TITULO), '#.00')"
            dbSlave.BeginTrans
            dbSlave.Execute SQL
            dbSlave.CommitTrans
            'Calculamos la antiguedad en años
            SQL = "Update CALCULOSAUXILIARES " _
            & "Set ANOS = Format((ANTIGUEDAD / BASICO / 0.02), '0')"
            dbSlave.BeginTrans
            dbSlave.Execute SQL
            dbSlave.CommitTrans
            'Calculamos el Coeficiente deL Titulo
            SQL = "Update CALCULOSAUXILIARES " _
            & "Set COEFICIENTETITULO = Format((TITULO / BASICO), '#.00')"
            dbSlave.BeginTrans
            dbSlave.Execute SQL
            dbSlave.CommitTrans
            'Recalculamos el Básico
            If IsNumeric(Alicuota) Then
                SQL = "Update CALCULOSAUXILIARES " _
                & "Set BASICO = Format((BASICO * " & De_Num_a_Tx_01(Alicuota, , 2) & "), '#.00')"
            ElseIf IsNumeric(ImporteFijo) Then
                SQL = "Update CALCULOSAUXILIARES " _
                & "Set BASICO = Format((BASICO + " & De_Num_a_Tx_01(ImporteFijo, , 2) & "), '#.00')"
            End If
            dbSlave.BeginTrans
            dbSlave.Execute SQL
            dbSlave.CommitTrans
            'Recalculamos la antiguedad con los datos obtenidos
            SQL = "Update CALCULOSAUXILIARES " _
            & "Set ANTIGUEDAD = Format((ANOS * BASICO * 0.02), '#.00')"
            dbSlave.BeginTrans
            dbSlave.Execute SQL
            dbSlave.CommitTrans
            'Recalculamos el título con los datos obtenidos
            SQL = "Update CALCULOSAUXILIARES " _
            & "Set TITULO = Format((BASICO * COEFICIENTETITULO), '#.00')"
            dbSlave.BeginTrans
            dbSlave.Execute SQL
            dbSlave.CommitTrans
            'Recalculamos la Bonificación FONAVI con los datos obtenidos
            SQL = "Update CALCULOSAUXILIARES " _
            & "Set BONIFICACIONFONAVI = Format((BASICO + ANTIGUEDAD + TITULO) * COEFICIENTEFONAVI, '#.00') " _
            & "Where BONIFICACIONFONAVI > 0"
            dbSlave.BeginTrans
            dbSlave.Execute SQL
            dbSlave.CommitTrans
            'Modificamos el campo Compensacion Funcional de la Tabla de Calculos Auxiliares para incorporar la Compensacion Funcional de la tabla base
            SQL = "Update CALCULOSAUXILIARES Inner Join " _
            & "(Select PUESTOLABORAL, IMPORTE From LIQUIDACIONSUELDOS " _
            & "Where CodigoLiquidacion = " & "'" & CodigoBase _
            & "' And CODIGOCONCEPTO = '0002') As LIQUIDACIONSUELDOS " _
            & "On CALCULOSAUXILIARES.PUESTOLABORAL = LIQUIDACIONSUELDOS.PUESTOLABORAL " _
            & "Set CALCULOSAUXILIARES.COMPENSACIONFUNCIONAL = LIQUIDACIONSUELDOS.IMPORTE"
            dbSlave.BeginTrans
            dbSlave.Execute SQL
            dbSlave.CommitTrans
            'Recalculamos la Compensación Funcional
            SQL = "Update CALCULOSAUXILIARES " _
            & "Set COMPENSACIONFUNCIONAL = Format(((BASICO + TITULO + ANTIGUEDAD) * 0.51), '#.00') " _
            & "Where COMPENSACIONFUNCIONAL > 0"
            dbSlave.BeginTrans
            dbSlave.Execute SQL
            dbSlave.CommitTrans
            'Modificamos el campo Funcion Especifica de la Tabla de Calculos Auxiliares para incorporar la Funcion Especifica de la tabla base
            SQL = "Update CALCULOSAUXILIARES Inner Join " _
            & "(Select PUESTOLABORAL, IMPORTE From LIQUIDACIONSUELDOS " _
            & "Where CodigoLiquidacion = " & "'" & CodigoBase _
            & "' And CODIGOCONCEPTO = '0065') As LIQUIDACIONSUELDOS " _
            & "On CALCULOSAUXILIARES.PUESTOLABORAL = LIQUIDACIONSUELDOS.PUESTOLABORAL " _
            & "Set CALCULOSAUXILIARES.FUNCIONESPECIFICA = LIQUIDACIONSUELDOS.IMPORTE"
            dbSlave.BeginTrans
            dbSlave.Execute SQL
            dbSlave.CommitTrans
            'Recalculamos la Funcion Especifica
            SQL = "Update CALCULOSAUXILIARES " _
            & "Set FUNCIONESPECIFICA = Format(((BASICO + TITULO + ANTIGUEDAD + COMPENSACIONFUNCIONAL) * 1.10), '#.00') " _
            & "Where FUNCIONESPECIFICA > 0"
            dbSlave.BeginTrans
            dbSlave.Execute SQL
            dbSlave.CommitTrans
            'Modificamos el campo Presentismo de la Tabla de Calculos Auxiliares para incorporar el Presentismo de la tabla base
            SQL = "Update CALCULOSAUXILIARES Inner Join " _
            & "(Select PUESTOLABORAL, IMPORTE From LIQUIDACIONSUELDOS " _
            & "Where CodigoLiquidacion = " & "'" & CodigoBase _
            & "' And CODIGOCONCEPTO = '0054') As LIQUIDACIONSUELDOS " _
            & "On CALCULOSAUXILIARES.PUESTOLABORAL = LIQUIDACIONSUELDOS.PUESTOLABORAL " _
            & "Set CALCULOSAUXILIARES.PRESENTISMO = LIQUIDACIONSUELDOS.IMPORTE"
            dbSlave.BeginTrans
            dbSlave.Execute SQL
            dbSlave.CommitTrans
            'Recalculamos el Presentismo
            SQL = "Update CALCULOSAUXILIARES " _
            & "Set PRESENTISMO = Format(((BASICO + TITULO + ANTIGUEDAD + COMPENSACIONFUNCIONAL + BONIFICACIONFONAVI + FUNCIONESPECIFICA) * 0.15), '#.00') " _
            & "Where PRESENTISMO > 0"
            dbSlave.BeginTrans
            dbSlave.Execute SQL
            dbSlave.CommitTrans
            'Modificamos el campo Dedicación Exclusiva de la Tabla de Calculos Auxiliares para incorporar el Presentismo de la tabla base
            SQL = "Update CALCULOSAUXILIARES Inner Join " _
            & "(Select PUESTOLABORAL, IMPORTE From LIQUIDACIONSUELDOS " _
            & "Where CodigoLiquidacion = " & "'" & CodigoBase _
            & "' And CODIGOCONCEPTO = '0064') As LIQUIDACIONSUELDOS " _
            & "On CALCULOSAUXILIARES.PUESTOLABORAL = LIQUIDACIONSUELDOS.PUESTOLABORAL " _
            & "Set CALCULOSAUXILIARES.DEDICACIONEXCLUSIVA = LIQUIDACIONSUELDOS.IMPORTE"
            dbSlave.BeginTrans
            dbSlave.Execute SQL
            dbSlave.CommitTrans
            'Recalculamos la Dedicación Exclusiva
            SQL = "Update CALCULOSAUXILIARES " _
            & "Set DEDICACIONEXCLUSIVA = Format(((BASICO) * 0.80), '#.00') " _
            & "Where DEDICACIONEXCLUSIVA > 0"
            dbSlave.BeginTrans
            dbSlave.Execute SQL
            dbSlave.CommitTrans
            'Copiamos el concepto Antiguedad (0005) de la tabla Calculo Auxiliares a la liquidación destino
            SQL = "Insert Into LIQUIDACIONSUELDOS (CodigoLiquidacion, PuestoLaboral, CodigoConcepto, Importe)" & _
            " Select '" & CodigoDestino & "', PUESTOLABORAL, '0005', ANTIGUEDAD" & _
            " From CALCULOSAUXILIARES Where ANTIGUEDAD > 0"
            dbSlave.BeginTrans
            dbSlave.Execute SQL
            dbSlave.CommitTrans
            'Copiamos el concepto Compensación Funcional (0002) de la tabla Calculo Auxiliares a la liquidación destino
            SQL = "Insert Into LIQUIDACIONSUELDOS (CodigoLiquidacion, PuestoLaboral, CodigoConcepto, Importe)" & _
            " Select '" & CodigoDestino & "', PUESTOLABORAL, '0002', COMPENSACIONFUNCIONAL" & _
            " From CALCULOSAUXILIARES Where COMPENSACIONFUNCIONAL > 0"
            dbSlave.BeginTrans
            dbSlave.Execute SQL
            dbSlave.CommitTrans
            'Copiamos el concepto Función Especifica (0065) de la tabla Calculo Auxiliares a la liquidación destino
            SQL = "Insert Into LIQUIDACIONSUELDOS (CodigoLiquidacion, PuestoLaboral, CodigoConcepto, Importe)" & _
            " Select '" & CodigoDestino & "', PUESTOLABORAL, '0065', FUNCIONESPECIFICA" & _
            " From CALCULOSAUXILIARES Where FUNCIONESPECIFICA > 0"
            dbSlave.BeginTrans
            dbSlave.Execute SQL
            dbSlave.CommitTrans
            'Copiamos el concepto Bonificación FONAVI (0108) de la tabla Calculo Auxiliares a la liquidación destino
            SQL = "Insert Into LIQUIDACIONSUELDOS (CodigoLiquidacion, PuestoLaboral, CodigoConcepto, Importe)" & _
            " Select '" & CodigoDestino & "', PUESTOLABORAL, '0108', BONIFICACIONFONAVI" & _
            " From CALCULOSAUXILIARES Where BONIFICACIONFONAVI > 0"
            dbSlave.BeginTrans
            dbSlave.Execute SQL
            dbSlave.CommitTrans
            'Copiamos el concepto Presentismo (0054) de la tabla Calculo Auxiliares a la liquidación destino
            SQL = "Insert Into LIQUIDACIONSUELDOS (CodigoLiquidacion, PuestoLaboral, CodigoConcepto, Importe)" & _
            " Select '" & CodigoDestino & "', PUESTOLABORAL, '0054', PRESENTISMO" & _
            " From CALCULOSAUXILIARES Where PRESENTISMO > 0"
            dbSlave.BeginTrans
            dbSlave.Execute SQL
            dbSlave.CommitTrans
            'Copiamos el concepto DedicacionExclusiva (0064) de la tabla Calculo Auxiliares a la liquidación destino
            SQL = "Insert Into LIQUIDACIONSUELDOS (CodigoLiquidacion, PuestoLaboral, CodigoConcepto, Importe)" & _
            " Select '" & CodigoDestino & "', PUESTOLABORAL, '0064', DEDICACIONEXCLUSIVA" & _
            " From CALCULOSAUXILIARES Where DEDICACIONEXCLUSIVA > 0"
            dbSlave.BeginTrans
            dbSlave.Execute SQL
            dbSlave.CommitTrans
            'Copiamos el concepto TITULO (0008) de la tabla Calculo Auxiliares a la liquidación destino
            SQL = "Insert Into LIQUIDACIONSUELDOS (CodigoLiquidacion, PuestoLaboral, CodigoConcepto, Importe)" & _
            " Select '" & CodigoDestino & "', PUESTOLABORAL, '0008', TITULO" & _
            " From CALCULOSAUXILIARES Where TITULO > 0"
            dbSlave.BeginTrans
            dbSlave.Execute SQL
            dbSlave.CommitTrans
            'Copiamos el concepto Basico (0001) de la tabla Calculo Auxiliares a la liquidación destino
            SQL = "Insert Into LIQUIDACIONSUELDOS (CodigoLiquidacion, PuestoLaboral, CodigoConcepto, Importe)" & _
            " Select '" & CodigoDestino & "', PUESTOLABORAL, '0001', BASICO" & _
            " From CALCULOSAUXILIARES Where BASICO > 0"
            dbSlave.BeginTrans
            dbSlave.Execute SQL
            dbSlave.CommitTrans
            
            'Tercero, borramos la tabla auxiliar
            SQL = "Drop Table CALCULOSAUXILIARES"
            dbSlave.BeginTrans
            dbSlave.Execute SQL
            dbSlave.CommitTrans
            
            'Cuarto, calculamos el nuevo Haber Óptimo (0115)
            SQL = "Insert Into LIQUIDACIONSUELDOS (CodigoLiquidacion, PuestoLaboral, CodigoConcepto, Importe)" & _
            " Select '" & CodigoDestino & "', PUESTOLABORAL, '0115', HABEROPTIMO" & _
            " From (Select PUESTOLABORAL, Sum(IMPORTE) As HABEROPTIMO" & _
            " From LIQUIDACIONSUELDOS Where CodigoLiquidacion = " & "'" & CodigoDestino _
            & "' Group By PUESTOLABORAL)"
            dbSlave.BeginTrans
            dbSlave.Execute SQL
            dbSlave.CommitTrans
            
            'Quinto, copiamos Salario Familiar y Discapacitados
            'Copiamos el concepto Salario Familiar (0003) de la liquidación base a la liquidación destino
            SQL = "Insert Into LIQUIDACIONSUELDOS (CodigoLiquidacion, PuestoLaboral, CodigoConcepto, Importe)" & _
            " Select '" & CodigoDestino & "', PuestoLaboral, '0003', Importe" & _
            " From LIQUIDACIONSUELDOS Where CodigoLiquidacion = " & "'" & CodigoBase & "' And CODIGOCONCEPTO = '0003'"
            dbSlave.BeginTrans
            dbSlave.Execute SQL
            dbSlave.CommitTrans
            'Copiamos el concepto Discapacitados (0058) de la liquidación base a la liquidación destino
            SQL = "Insert Into LIQUIDACIONSUELDOS (CodigoLiquidacion, PuestoLaboral, CodigoConcepto, Importe)" & _
            " Select '" & CodigoDestino & "', PuestoLaboral, '0058', Importe" & _
            " From LIQUIDACIONSUELDOS Where CodigoLiquidacion = " & "'" & CodigoBase & "' And CODIGOCONCEPTO = '0058'"
            dbSlave.BeginTrans
            dbSlave.Execute SQL
            dbSlave.CommitTrans
            
            'Sexto, calculamos el nuevo Total Bruto (9998)
            SQL = "Insert Into LIQUIDACIONSUELDOS (CodigoLiquidacion, PuestoLaboral, CodigoConcepto, Importe)" & _
            " Select '" & CodigoDestino & "', PUESTOLABORAL, '9998', TOTALBRUTO" & _
            " From (Select PUESTOLABORAL, Sum(IMPORTE) As TOTALBRUTO" & _
            " From LIQUIDACIONSUELDOS Where CodigoLiquidacion = " & "'" & CodigoDestino _
            & "' And CODIGOCONCEPTO <> '0115' Group By PUESTOLABORAL)"
            dbSlave.BeginTrans
            dbSlave.Execute SQL
            dbSlave.CommitTrans
            
            'Septimo, calculamos los descuentos
            'Calculamos la Jubilacion Personal a partir del Haber Óptimo
            SQL = "Insert Into LIQUIDACIONSUELDOS (CodigoLiquidacion, PuestoLaboral, CodigoConcepto, Importe)" & _
            " Select '" & CodigoDestino & "', PuestoLaboral, '0208', Format((Importe * 0.185), '#.00')" & _
            " From LIQUIDACIONSUELDOS Where CodigoLiquidacion = " & "'" & CodigoDestino & "' And CODIGOCONCEPTO = '0115'"
            dbSlave.BeginTrans
            dbSlave.Execute SQL
            dbSlave.CommitTrans
            'Calculamos la Jubilacion Estatal a partir del Haber Óptimo
            SQL = "Insert Into LIQUIDACIONSUELDOS (CodigoLiquidacion, PuestoLaboral, CodigoConcepto, Importe)" & _
            " Select '" & CodigoDestino & "', PuestoLaboral, '0209', Format((Importe * 0.185), '#.00')" & _
            " From LIQUIDACIONSUELDOS Where CodigoLiquidacion = " & "'" & CodigoDestino & "' And CODIGOCONCEPTO = '0115'"
            dbSlave.BeginTrans
            dbSlave.Execute SQL
            dbSlave.CommitTrans
            'Calculamos la Obra Social Personal a partir del Haber Óptimo
            SQL = "Insert Into LIQUIDACIONSUELDOS (CodigoLiquidacion, PuestoLaboral, CodigoConcepto, Importe)" & _
            " Select '" & CodigoDestino & "', PuestoLaboral, '0212', Format((Importe * 0.05), '#.00')" & _
            " From LIQUIDACIONSUELDOS Where CodigoLiquidacion = " & "'" & CodigoDestino & "' And CODIGOCONCEPTO = '0115'"
            dbSlave.BeginTrans
            dbSlave.Execute SQL
            dbSlave.CommitTrans
            'Calculamos la Obra Social Estatal a partir del Haber Óptimo
            SQL = "Insert Into LIQUIDACIONSUELDOS (CodigoLiquidacion, PuestoLaboral, CodigoConcepto, Importe)" & _
            " Select '" & CodigoDestino & "', PuestoLaboral, '0213', Format((Importe * 0.04), '#.00')" & _
            " From LIQUIDACIONSUELDOS Where CodigoLiquidacion = " & "'" & CodigoDestino & "' And CODIGOCONCEPTO = '0115'"
            dbSlave.BeginTrans
            dbSlave.Execute SQL
            dbSlave.CommitTrans
            'Calculamos la ART a partir del Haber Óptimo
            SQL = "Insert Into LIQUIDACIONSUELDOS (CodigoLiquidacion, PuestoLaboral, CodigoConcepto, Importe)" & _
            " Select '" & CodigoDestino & "', PuestoLaboral, '0381', Format((Importe * 0.0173 + 0.6), '#.00')" & _
            " From LIQUIDACIONSUELDOS Where CodigoLiquidacion = " & "'" & CodigoDestino & "' And CODIGOCONCEPTO = '0115'"
            dbSlave.BeginTrans
            dbSlave.Execute SQL
            dbSlave.CommitTrans
            'Copiamos el concepto Cuota Sindical (0219 y 0227) de la liquidación base a la liquidación destino (0219) - (VERIFICAR PORQUE COPIAMOS EL MISMO IMPORTE)
            SQL = "Insert Into LIQUIDACIONSUELDOS (CodigoLiquidacion, PuestoLaboral, CodigoConcepto, Importe)" & _
            " Select '" & CodigoDestino & "', PUESTOLABORAL, '0219', CUOTASINDICAL" & _
            " From (Select PUESTOLABORAL, Sum(IMPORTE) As CUOTASINDICAL" & _
            " From LIQUIDACIONSUELDOS Where" & _
            " ((CodigoLiquidacion = " & "'" & CodigoBase & "' And CODIGOCONCEPTO = '0219') Or" & _
            " (CodigoLiquidacion = " & "'" & CodigoBase & "' And CODIGOCONCEPTO = '0227')) Group By PUESTOLABORAL)"
            dbSlave.BeginTrans
            dbSlave.Execute SQL
            dbSlave.CommitTrans
            'Copiamos el concepto Seguro de Vida Optativo (Varios) de la liquidación base a la liquidación destino (0317) - (VERIFICAR PORQUE COPIAMOS EL MISMO IMPORTE)
            SQL = "Insert Into LIQUIDACIONSUELDOS (CodigoLiquidacion, PuestoLaboral, CodigoConcepto, Importe)" & _
            " Select '" & CodigoDestino & "', PUESTOLABORAL, '0317', SEGUROOPTATIVO" & _
            " From (Select PUESTOLABORAL, Sum(IMPORTE) As SEGUROOPTATIVO" & _
            " From LIQUIDACIONSUELDOS Where" & _
            " ((CodigoLiquidacion = " & "'" & CodigoBase & "' And CODIGOCONCEPTO = '0317') Or" & _
            " (CodigoLiquidacion = " & "'" & CodigoBase & "' And CODIGOCONCEPTO = '0361') Or" & _
            " (CodigoLiquidacion = " & "'" & CodigoBase & "' And CODIGOCONCEPTO = '0367') Or" & _
            " (CodigoLiquidacion = " & "'" & CodigoBase & "' And CODIGOCONCEPTO = '0370') Or" & _
            " (CodigoLiquidacion = " & "'" & CodigoBase & "' And CODIGOCONCEPTO = '0373') Or" & _
            " (CodigoLiquidacion = " & "'" & CodigoBase & "' And CODIGOCONCEPTO = '0374')) Group By PUESTOLABORAL)"
            dbSlave.BeginTrans
            dbSlave.Execute SQL
            dbSlave.CommitTrans
            'Copiamos el concepto Seguro de Vida Obligatorio (0360, 0366, 0369) de la liquidación base a la liquidación destino (0360)
            SQL = "Insert Into LIQUIDACIONSUELDOS (CodigoLiquidacion, PuestoLaboral, CodigoConcepto, Importe)" & _
            " Select '" & CodigoDestino & "', PUESTOLABORAL, '0360', SEGUROOBLIGATORIO" & _
            " From (Select PUESTOLABORAL, Sum(IMPORTE) As SEGUROOBLIGATORIO" & _
            " From LIQUIDACIONSUELDOS Where" & _
            " ((CodigoLiquidacion = " & "'" & CodigoBase & "' And CODIGOCONCEPTO = '0360') Or" & _
            " (CodigoLiquidacion = " & "'" & CodigoBase & "' And CODIGOCONCEPTO = '0366') Or" & _
            " (CodigoLiquidacion = " & "'" & CodigoBase & "' And CODIGOCONCEPTO = '0369')) Group By PUESTOLABORAL)"
            dbSlave.BeginTrans
            dbSlave.Execute SQL
            dbSlave.CommitTrans
        
            Call InfoGeneral("", LiquidacionPruebaSISPER)
            Call InfoGeneral("- LIQUIDACIÓN FINALIZADA -", LiquidacionPruebaSISPER)
        End If
        
    Case "0603" 'Plus Remunerativo
        Call InfoGeneral("- INICIANDO LIQUIDACIÓN PRUEBA -", LiquidacionPruebaSISPER)
        Call InfoGeneral("", LiquidacionPruebaSISPER)
        'Buscamos el Haber Óptimo de la liquidación base
        SQL = "Select IMPORTE, PUESTOLABORAL from LIQUIDACIONSUELDOS" _
        & " Where CODIGOLIQUIDACION = '" & CodigoBase _
        & "' And CODIGOCONCEPTO = '0115'"
        If SQLNoMatch(SQL) = True Then
            Call InfoGeneral("La liquidación base no posee Haber Óptimo", LiquidacionPruebaSISPER)
        Else
            'Primero, copiamos de la liquidación base todos los conceptos remunerativos que no van a sufrir alteración
            'Copiamos el concepto Básico (0001) de la liquidación base a la liquidación destino
            SQL = "Insert Into LIQUIDACIONSUELDOS (CodigoLiquidacion, PuestoLaboral, CodigoConcepto, Importe)" & _
            " Select '" & CodigoDestino & "', PuestoLaboral, '0001', Importe" & _
            " From LIQUIDACIONSUELDOS Where CodigoLiquidacion = " & "'" & CodigoBase & "' And CODIGOCONCEPTO = '0001'"
            dbSlave.BeginTrans
            dbSlave.Execute SQL
            dbSlave.CommitTrans
            'Copiamos el concepto Antigüedad (0005) de la liquidación base a la liquidación destino
            SQL = "Insert Into LIQUIDACIONSUELDOS (CodigoLiquidacion, PuestoLaboral, CodigoConcepto, Importe)" & _
            " Select '" & CodigoDestino & "', PuestoLaboral, '0005', Importe" & _
            " From LIQUIDACIONSUELDOS Where CodigoLiquidacion = " & "'" & CodigoBase & "' And CODIGOCONCEPTO = '0005'"
            dbSlave.BeginTrans
            dbSlave.Execute SQL
            dbSlave.CommitTrans
            'Copiamos el concepto Título (0008) de la liquidación base a la liquidación destino
            SQL = "Insert Into LIQUIDACIONSUELDOS (CodigoLiquidacion, PuestoLaboral, CodigoConcepto, Importe)" & _
            " Select '" & CodigoDestino & "', PuestoLaboral, '0008', Importe" & _
            " From LIQUIDACIONSUELDOS Where CodigoLiquidacion = " & "'" & CodigoBase & "' And CODIGOCONCEPTO = '0008'"
            dbSlave.BeginTrans
            dbSlave.Execute SQL
            dbSlave.CommitTrans
            'Copiamos el concepto Gastos de Representación (0028) de la liquidación base a la liquidación destino
            SQL = "Insert Into LIQUIDACIONSUELDOS (CodigoLiquidacion, PuestoLaboral, CodigoConcepto, Importe)" & _
            " Select '" & CodigoDestino & "', PuestoLaboral, '0028', Importe" & _
            " From LIQUIDACIONSUELDOS Where CodigoLiquidacion = " & "'" & CodigoBase & "' And CODIGOCONCEPTO = '0028'"
            dbSlave.BeginTrans
            dbSlave.Execute SQL
            dbSlave.CommitTrans
            'Copiamos el concepto Dedicación Exclusiva (0064) de la liquidación base a la liquidación destino
            SQL = "Insert Into LIQUIDACIONSUELDOS (CodigoLiquidacion, PuestoLaboral, CodigoConcepto, Importe)" & _
            " Select '" & CodigoDestino & "', PuestoLaboral, '0064', Importe" & _
            " From LIQUIDACIONSUELDOS Where CodigoLiquidacion = " & "'" & CodigoBase & "' And CODIGOCONCEPTO = '0064'"
            dbSlave.BeginTrans
            dbSlave.Execute SQL
            dbSlave.CommitTrans
            'Copiamos el concepto Bonificación FONAVI (0108) de la liquidación base a la liquidación destino
            SQL = "Insert Into LIQUIDACIONSUELDOS (CodigoLiquidacion, PuestoLaboral, CodigoConcepto, Importe)" & _
            " Select '" & CodigoDestino & "', PuestoLaboral, '0108', Importe" & _
            " From LIQUIDACIONSUELDOS Where CodigoLiquidacion = " & "'" & CodigoBase & "' And CODIGOCONCEPTO = '0108'"
            dbSlave.BeginTrans
            dbSlave.Execute SQL
            dbSlave.CommitTrans
            'Copiamos el concepto AD.REM.DEC.2232/10-415/13 (0134) de la liquidación base a la liquidación destino
            SQL = "Insert Into LIQUIDACIONSUELDOS (CodigoLiquidacion, PuestoLaboral, CodigoConcepto, Importe)" & _
            " Select '" & CodigoDestino & "', PuestoLaboral, '0134', Importe" & _
            " From LIQUIDACIONSUELDOS Where CodigoLiquidacion = " & "'" & CodigoBase & "' And CODIGOCONCEPTO = '0134'"
            dbSlave.BeginTrans
            dbSlave.Execute SQL
            dbSlave.CommitTrans
            'Copiamos el concepto ADICIONAL DEC 1619/05 C/A (0169) de la liquidación base a la liquidación destino
            SQL = "Insert Into LIQUIDACIONSUELDOS (CodigoLiquidacion, PuestoLaboral, CodigoConcepto, Importe)" & _
            " Select '" & CodigoDestino & "', PuestoLaboral, '0169', Importe" & _
            " From LIQUIDACIONSUELDOS Where CodigoLiquidacion = " & "'" & CodigoBase & "' And CODIGOCONCEPTO = '0169'"
            dbSlave.BeginTrans
            dbSlave.Execute SQL
            dbSlave.CommitTrans
            'Copiamos y Recalculamos el Plus Remunerativo (0603)
            If IsNumeric(Alicuota) Then
                SQL = "Insert Into LIQUIDACIONSUELDOS (CodigoLiquidacion, PuestoLaboral, CodigoConcepto, Importe)" & _
                " Select '" & CodigoDestino & "', PuestoLaboral, '0603', Importe * " & De_Num_a_Tx_01(Alicuota, , 2) & _
                " From LIQUIDACIONSUELDOS Where CodigoLiquidacion = " & "'" & CodigoBase & "' And CODIGOCONCEPTO = '0603'"
            ElseIf IsNumeric(ImporteFijo) Then
                SQL = "Insert Into LIQUIDACIONSUELDOS (CodigoLiquidacion, PuestoLaboral, CodigoConcepto, Importe)" & _
                " Select '" & CodigoDestino & "', PuestoLaboral, '0603', Importe + " & De_Num_a_Tx_01(ImporteFijo, , 2) & _
                " From LIQUIDACIONSUELDOS Where CodigoLiquidacion = " & "'" & CodigoBase & "' And CODIGOCONCEPTO = '0603'"
            End If
            dbSlave.BeginTrans
            dbSlave.Execute SQL
            dbSlave.CommitTrans
            
            'Segundo, creamos una tabla adicional para calcular los conceptos que sufren cambio
            
            'Tercero, borramos la tabla auxiliar
            
            'Cuarto, calculamos el nuevo Haber Óptimo (0115)
            SQL = "Insert Into LIQUIDACIONSUELDOS (CodigoLiquidacion, PuestoLaboral, CodigoConcepto, Importe)" & _
            " Select '" & CodigoDestino & "', PUESTOLABORAL, '0115', HABEROPTIMO" & _
            " From (Select PUESTOLABORAL, Sum(IMPORTE) As HABEROPTIMO" & _
            " From LIQUIDACIONSUELDOS Where CodigoLiquidacion = " & "'" & CodigoDestino _
            & "' Group By PUESTOLABORAL)"
            dbSlave.BeginTrans
            dbSlave.Execute SQL
            dbSlave.CommitTrans
            
            'Quinto, copiamos Salario Familiar y Discapacitados
            'Copiamos el concepto Salario Familiar (0003) de la liquidación base a la liquidación destino
            SQL = "Insert Into LIQUIDACIONSUELDOS (CodigoLiquidacion, PuestoLaboral, CodigoConcepto, Importe)" & _
            " Select '" & CodigoDestino & "', PuestoLaboral, '0003', Importe" & _
            " From LIQUIDACIONSUELDOS Where CodigoLiquidacion = " & "'" & CodigoBase & "' And CODIGOCONCEPTO = '0003'"
            dbSlave.BeginTrans
            dbSlave.Execute SQL
            dbSlave.CommitTrans
            'Copiamos el concepto Discapacitados (0058) de la liquidación base a la liquidación destino
            SQL = "Insert Into LIQUIDACIONSUELDOS (CodigoLiquidacion, PuestoLaboral, CodigoConcepto, Importe)" & _
            " Select '" & CodigoDestino & "', PuestoLaboral, '0058', Importe" & _
            " From LIQUIDACIONSUELDOS Where CodigoLiquidacion = " & "'" & CodigoBase & "' And CODIGOCONCEPTO = '0058'"
            dbSlave.BeginTrans
            dbSlave.Execute SQL
            dbSlave.CommitTrans
            
            'Sexto, calculamos el nuevo Total Bruto (9998)
            SQL = "Insert Into LIQUIDACIONSUELDOS (CodigoLiquidacion, PuestoLaboral, CodigoConcepto, Importe)" & _
            " Select '" & CodigoDestino & "', PUESTOLABORAL, '9998', TOTALBRUTO" & _
            " From (Select PUESTOLABORAL, Sum(IMPORTE) As TOTALBRUTO" & _
            " From LIQUIDACIONSUELDOS Where CodigoLiquidacion = " & "'" & CodigoDestino _
            & "' And CODIGOCONCEPTO <> '0115' Group By PUESTOLABORAL)"
            dbSlave.BeginTrans
            dbSlave.Execute SQL
            dbSlave.CommitTrans
            
            'Septimo, calculamos los descuentos
            'Calculamos la Jubilacion Personal a partir del Haber Óptimo
            SQL = "Insert Into LIQUIDACIONSUELDOS (CodigoLiquidacion, PuestoLaboral, CodigoConcepto, Importe)" & _
            " Select '" & CodigoDestino & "', PuestoLaboral, '0208', Format((Importe * 0.185), '#.00')" & _
            " From LIQUIDACIONSUELDOS Where CodigoLiquidacion = " & "'" & CodigoDestino & "' And CODIGOCONCEPTO = '0115'"
            dbSlave.BeginTrans
            dbSlave.Execute SQL
            dbSlave.CommitTrans
            'Calculamos la Jubilacion Estatal a partir del Haber Óptimo
            SQL = "Insert Into LIQUIDACIONSUELDOS (CodigoLiquidacion, PuestoLaboral, CodigoConcepto, Importe)" & _
            " Select '" & CodigoDestino & "', PuestoLaboral, '0209', Format((Importe * 0.185), '#.00')" & _
            " From LIQUIDACIONSUELDOS Where CodigoLiquidacion = " & "'" & CodigoDestino & "' And CODIGOCONCEPTO = '0115'"
            dbSlave.BeginTrans
            dbSlave.Execute SQL
            dbSlave.CommitTrans
            'Calculamos la Obra Social Personal a partir del Haber Óptimo
            SQL = "Insert Into LIQUIDACIONSUELDOS (CodigoLiquidacion, PuestoLaboral, CodigoConcepto, Importe)" & _
            " Select '" & CodigoDestino & "', PuestoLaboral, '0212', Format((Importe * 0.05), '#.00')" & _
            " From LIQUIDACIONSUELDOS Where CodigoLiquidacion = " & "'" & CodigoDestino & "' And CODIGOCONCEPTO = '0115'"
            dbSlave.BeginTrans
            dbSlave.Execute SQL
            dbSlave.CommitTrans
            'Calculamos la Obra Social Estatal a partir del Haber Óptimo
            SQL = "Insert Into LIQUIDACIONSUELDOS (CodigoLiquidacion, PuestoLaboral, CodigoConcepto, Importe)" & _
            " Select '" & CodigoDestino & "', PuestoLaboral, '0213', Format((Importe * 0.04), '#.00')" & _
            " From LIQUIDACIONSUELDOS Where CodigoLiquidacion = " & "'" & CodigoDestino & "' And CODIGOCONCEPTO = '0115'"
            dbSlave.BeginTrans
            dbSlave.Execute SQL
            dbSlave.CommitTrans
            'Calculamos la ART a partir del Haber Óptimo
            SQL = "Insert Into LIQUIDACIONSUELDOS (CodigoLiquidacion, PuestoLaboral, CodigoConcepto, Importe)" & _
            " Select '" & CodigoDestino & "', PuestoLaboral, '0381', Format((Importe * 0.0173 + 0.6), '#.00')" & _
            " From LIQUIDACIONSUELDOS Where CodigoLiquidacion = " & "'" & CodigoDestino & "' And CODIGOCONCEPTO = '0115'"
            dbSlave.BeginTrans
            dbSlave.Execute SQL
            dbSlave.CommitTrans
            'Copiamos el concepto Cuota Sindical (0219 y 0227) de la liquidación base a la liquidación destino (0219) - (VERIFICAR PORQUE COPIAMOS EL MISMO IMPORTE)
            SQL = "Insert Into LIQUIDACIONSUELDOS (CodigoLiquidacion, PuestoLaboral, CodigoConcepto, Importe)" & _
            " Select '" & CodigoDestino & "', PUESTOLABORAL, '0219', CUOTASINDICAL" & _
            " From (Select PUESTOLABORAL, Sum(IMPORTE) As CUOTASINDICAL" & _
            " From LIQUIDACIONSUELDOS Where" & _
            " ((CodigoLiquidacion = " & "'" & CodigoBase & "' And CODIGOCONCEPTO = '0219') Or" & _
            " (CodigoLiquidacion = " & "'" & CodigoBase & "' And CODIGOCONCEPTO = '0227')) Group By PUESTOLABORAL)"
            dbSlave.BeginTrans
            dbSlave.Execute SQL
            dbSlave.CommitTrans
            'Copiamos el concepto Seguro de Vida Optativo (Varios) de la liquidación base a la liquidación destino (0317) - (VERIFICAR PORQUE COPIAMOS EL MISMO IMPORTE)
            SQL = "Insert Into LIQUIDACIONSUELDOS (CodigoLiquidacion, PuestoLaboral, CodigoConcepto, Importe)" & _
            " Select '" & CodigoDestino & "', PUESTOLABORAL, '0317', SEGUROOPTATIVO" & _
            " From (Select PUESTOLABORAL, Sum(IMPORTE) As SEGUROOPTATIVO" & _
            " From LIQUIDACIONSUELDOS Where" & _
            " ((CodigoLiquidacion = " & "'" & CodigoBase & "' And CODIGOCONCEPTO = '0317') Or" & _
            " (CodigoLiquidacion = " & "'" & CodigoBase & "' And CODIGOCONCEPTO = '0361') Or" & _
            " (CodigoLiquidacion = " & "'" & CodigoBase & "' And CODIGOCONCEPTO = '0367') Or" & _
            " (CodigoLiquidacion = " & "'" & CodigoBase & "' And CODIGOCONCEPTO = '0370') Or" & _
            " (CodigoLiquidacion = " & "'" & CodigoBase & "' And CODIGOCONCEPTO = '0373') Or" & _
            " (CodigoLiquidacion = " & "'" & CodigoBase & "' And CODIGOCONCEPTO = '0374')) Group By PUESTOLABORAL)"
            dbSlave.BeginTrans
            dbSlave.Execute SQL
            dbSlave.CommitTrans
            'Copiamos el concepto Seguro de Vida Obligatorio (0360, 0366, 0369) de la liquidación base a la liquidación destino (0360)
            SQL = "Insert Into LIQUIDACIONSUELDOS (CodigoLiquidacion, PuestoLaboral, CodigoConcepto, Importe)" & _
            " Select '" & CodigoDestino & "', PUESTOLABORAL, '0360', SEGUROOBLIGATORIO" & _
            " From (Select PUESTOLABORAL, Sum(IMPORTE) As SEGUROOBLIGATORIO" & _
            " From LIQUIDACIONSUELDOS Where" & _
            " ((CodigoLiquidacion = " & "'" & CodigoBase & "' And CODIGOCONCEPTO = '0360') Or" & _
            " (CodigoLiquidacion = " & "'" & CodigoBase & "' And CODIGOCONCEPTO = '0366') Or" & _
            " (CodigoLiquidacion = " & "'" & CodigoBase & "' And CODIGOCONCEPTO = '0369')) Group By PUESTOLABORAL)"
            dbSlave.BeginTrans
            dbSlave.Execute SQL
            dbSlave.CommitTrans
        
            Call InfoGeneral("", LiquidacionPruebaSISPER)
            Call InfoGeneral("- LIQUIDACIÓN FINALIZADA -", LiquidacionPruebaSISPER)
        End If
               
    Case Else
        Call InfoGeneral("- PROCEDIMIENTO DE LIQUIDACIÓN AÚN NO PREVISTO PARA EL CONCEPTO ESPECIFICADO -", LiquidacionPruebaSISPER)
    
    End Select
    
End Sub

Public Function ImporteLimiteDeduccionGeneralSIRADIG(CodigoSIRADIG As String, PeriodoLiquidacion As String, _
Optional GananciaNeta As Double, Optional SeguroVidaEnRecibo As Double = 0) As Double

    Dim SQL As String
    Dim dblLimiteDeduccion As Double
    Dim datFechaTope As Date
    
    datFechaTope = BuscarUltimoDiaDelPeriodo(PeriodoLiquidacion)
    
    Select Case CodigoSIRADIG
    Case Is = "8"
        SQL = "Select SERVICIODOMESTICO From DEDUCCIONES4TACATEGORIA " _
            & "Where FECHA <= #" & Format(datFechaTope, "MM/DD/YYYY") & "# " _
            & "Order by FECHA Desc"
        Set rstBuscarSlave = New ADODB.Recordset
        rstBuscarSlave.Open SQL, dbSlave, adOpenDynamic, adLockOptimistic
        dblLimiteDeduccion = rstBuscarSlave.Fields(0) / 12
        dblLimiteDeduccion = dblLimiteDeduccion * Val(Left(PeriodoLiquidacion, 2))
        rstBuscarSlave.Close
    Case Is = "2"
        SQL = "Select SEGURODEVIDA From DEDUCCIONES4TACATEGORIA " _
            & "Where FECHA <= #" & Format(datFechaTope, "MM/DD/YYYY") & "# " _
            & "Order by FECHA Desc"
        Set rstBuscarSlave = New ADODB.Recordset
        rstBuscarSlave.Open SQL, dbSlave, adOpenDynamic, adLockOptimistic
        dblLimiteDeduccion = rstBuscarSlave.Fields(0) / 12
        dblLimiteDeduccion = dblLimiteDeduccion * Val(Left(PeriodoLiquidacion, 2))
        rstBuscarSlave.Close
    Case Is = "1"
        SQL = "Select CUOTAMEDICOASISTENCIAL From DEDUCCIONES4TACATEGORIA " _
            & "Where FECHA <= #" & Format(datFechaTope, "MM/DD/YYYY") & "# " _
            & "Order by FECHA Desc"
        Set rstBuscarSlave = New ADODB.Recordset
        rstBuscarSlave.Open SQL, dbSlave, adOpenDynamic, adLockOptimistic
        dblLimiteDeduccion = GananciaNeta * rstBuscarSlave.Fields(0)
        rstBuscarSlave.Close
    Case Is = "3"
        SQL = "Select DONACIONES From DEDUCCIONES4TACATEGORIA " _
            & "Where FECHA <= #" & Format(datFechaTope, "MM/DD/YYYY") & "# " _
            & "Order by FECHA Desc"
        Set rstBuscarSlave = New ADODB.Recordset
        rstBuscarSlave.Open SQL, dbSlave, adOpenDynamic, adLockOptimistic
        dblLimiteDeduccion = GananciaNeta * rstBuscarSlave.Fields(0)
        rstBuscarSlave.Close
    Case Is = "7"
        SQL = "Select HONORARIOSMEDICOS From DEDUCCIONES4TACATEGORIA " _
            & "Where FECHA <= #" & Format(datFechaTope, "MM/DD/YYYY") & "# " _
            & "Order by FECHA Desc"
        Set rstBuscarSlave = New ADODB.Recordset
        rstBuscarSlave.Open SQL, dbSlave, adOpenDynamic, adLockOptimistic
        dblLimiteDeduccion = GananciaNeta * rstBuscarSlave.Fields(0)
        rstBuscarSlave.Close
    Case Is = "22"
        SQL = "Select ALQUILERES From DEDUCCIONES4TACATEGORIA " _
            & "Where FECHA <= #" & Format(datFechaTope, "MM/DD/YYYY") & "# " _
            & "Order by FECHA Desc"
        Set rstBuscarSlave = New ADODB.Recordset
        rstBuscarSlave.Open SQL, dbSlave, adOpenDynamic, adLockOptimistic
        dblLimiteDeduccion = rstBuscarSlave.Fields(0) / 12
        dblLimiteDeduccion = dblLimiteDeduccion * Val(Left(PeriodoLiquidacion, 2))
        rstBuscarSlave.Close
    End Select
    
    ImporteLimiteDeduccionGeneralSIRADIG = Round(dblLimiteDeduccion, 2)
    
    Set rstBuscarSlave = Nothing
    SQL = ""
    dblLimiteDeduccion = 0
    
End Function

Public Function ImporteLimiteDeduccionPersonalSIRADIG(DenominacionDeduccion As String, Fecha As Date, _
Optional GciaNetaAntesDeDeduccionEspecial As Double) As Double

    Dim SQL As String
    Dim dblImporteMensual As Double
    Dim dblLimiteDeduccion As Double
    Dim datFechaControl As Date
    
    Select Case DenominacionDeduccion
    Case "MinimoNoImponible", "DeduccionEspecial"
        dblImporteMensual = ImporteDeduccionPersonal(DenominacionDeduccion, DateSerial(Year(Fecha) - 1, 12, 31))
        SQL = "Select " & DenominacionDeduccion & " From DEDUCCIONES4TACATEGORIA Where FECHA <= #" & Format(Fecha, "MM/DD/YYYY") & "# " _
        & "And Year(FECHA) = '" & Year(Fecha) & "' Order by FECHA Desc"
        If SQLNoMatch(SQL) = True Then
            dblImporteCalculado = dblImporteMensual * Month(Fecha)
        Else
            For i = 1 To Month(datFecha)
                datFechaControl = DateSerial(Year(Fecha), i, 1)
                datFechaControl = DateAdd("m", 1, datFechaControl)
                datFechaControl = DateAdd("d", -1, datFechaControl)
                dblImporteMensual = ImporteDeduccionPersonal(DenominacionDeduccion, datFechaControl)
                dblLimiteDeduccion = dblLimiteDeduccion + dblImporteMensual
            Next i
            datFechaControl = 0
        End If
        If DenominacionDeduccion = "DeduccionEspecial" Then
            If GciaNetaAntesDeDeduccionEspecial < dblLimiteDeduccion Then
                dblLimiteDeduccion = GciaNetaAntesDeDeduccionEspecial
            End If
        End If
    Case Else
    
    End Select
    
    
    
    ImporteLimiteDeduccionPersonalSIRADIG = Round(dblLimiteDeduccion, 2)
    
    Set rstBuscarSlave = Nothing
    SQL = ""
    dblLimiteDeduccion = 0
    
End Function


