Attribute VB_Name = "EliminarRegistros"
Public Sub EliminarAgente()

    Dim i As String
    Dim Borrar As Integer
    Dim SQL As String
    
    With ListadoAgentes.dgAgentes
        i = .Row
        If i <> 0 Then
            Borrar = MsgBox("Desea Borrar DEFINITIVAMENTE al Agente: " & .TextMatrix(i, 1) & "?" & vbCrLf & "Tenga en cuenta que toda la información relacionada al mismo será igualmente eliminada", vbQuestion + vbYesNo, "BORRANDO AGENTE")
            If Borrar = 6 Then
                Set rstRegistroSlave = New ADODB.Recordset
                SQL = "Select * from AGENTES Where PUESTOLABORAL = " & "'" & .TextMatrix(i, 2) & "'"
                rstRegistroSlave.Open SQL, dbSlave, adOpenForwardOnly, adLockOptimistic
                rstRegistroSlave.Delete
                rstRegistroSlave.Close
                Set rstRegistroSlave = Nothing
                Call ConfigurardgAgentes(ListadoAgentes)
                Call CargardgAgentes(ListadoAgentes)
            End If
        End If
        i = ""
    End With
    
End Sub

Public Sub EliminarPrecarizado()

    Dim i As String
    Dim Borrar As Integer
    Dim SQL As String
    
    With ListadoPrecarizados.dgPrecarizados
        i = .Row
        If i <> 0 Then
            Borrar = MsgBox("Desea Borrar DEFINITIVAMENTE al Facturero: " & .TextMatrix(i, 0) & "?" & vbCrLf & "Tenga en cuenta que toda la información relacionada al mismo será igualmente eliminada", vbQuestion + vbYesNo, "BORRANDO FACTURERO")
            If Borrar = 6 Then
                Set rstRegistroSlave = New ADODB.Recordset
                SQL = "Select * from PRECARIZADOS Where AGENTES = " & "'" & .TextMatrix(i, 0) & "'"
                rstRegistroSlave.Open SQL, dbSlave, adOpenForwardOnly, adLockOptimistic
                rstRegistroSlave.Delete
                rstRegistroSlave.Close
                Set rstRegistroSlave = Nothing
                Call ConfigurardgPrecarizados
                Call CargardgPrecarizados
            End If
        End If
        i = ""
    End With
    
End Sub


Public Sub EliminarConcepto()

    Dim i As String
    Dim Borrar As Integer
    Dim SQL As String
    
    With ListadoConceptos.dgConceptos
        i = .Row
        If i <> 0 Then
            Borrar = MsgBox("Desea Borrar DEFINITIVAMENTE al Concepto: " & .TextMatrix(i, 1) & "?" & vbCrLf & "Tenga en cuenta que toda la información relacionada al mismo será igualmente eliminada", vbQuestion + vbYesNo, "BORRANDO CONCEPTO")
            If Borrar = 6 Then
                Set rstRegistroSlave = New ADODB.Recordset
                SQL = "Select * from CONCEPTOS Where CODIGO = " & "'" & .TextMatrix(i, 0) & "'"
                rstRegistroSlave.Open SQL, dbSlave, adOpenForwardOnly, adLockOptimistic
                rstRegistroSlave.Delete
                rstRegistroSlave.Close
                Set rstRegistroSlave = Nothing
                ConfigurardgConceptos
                CargardgConceptos
            End If
        End If
        i = ""
    End With
    
End Sub

Public Sub EliminarNormaEscalaGanancias()

    Dim i As String
    Dim Borrar As Integer
    Dim SQL As String
    
    With ListadoEscalaGanancias.dgNormasEscalaGanancias
        i = .Row
        If i > 1 Then
            Borrar = MsgBox("Desea Borrar DEFINITIVAMENTE la Norma Legal: " & .TextMatrix(i, 0) & "?" & vbCrLf & "Tenga en cuenta que toda la información relacionada a la misma será igualmente eliminada", vbQuestion + vbYesNo, "BORRANDO NORMA LEGAL")
            If Borrar = 6 Then
                SQL = "Delete * from ESCALAAPLICABLEGANANCIAS" _
                & " Where NORMALEGAL = " & "'" & .TextMatrix(i, 0) & "'"
                dbSlave.BeginTrans
                dbSlave.Execute SQL
                dbSlave.CommitTrans
                ConfigurardgNormasEscalaGanancias
                CargardgNormasEscalaGanancias
                ConfigurardgEscalaGanancias
                CargardgEscalaGanancias (ListadoEscalaGanancias.dgNormasEscalaGanancias.TextMatrix(1, 0))
            End If
        End If
        i = ""
    End With
    
End Sub

Public Sub EliminarEscalaGanancias()

    Dim i As String
    Dim Borrar As Integer
    Dim SQL As String
    Dim strNormaLegal As String
    
    With ListadoEscalaGanancias.dgNormasEscalaGanancias
        i = .Row
        strNormaLegal = .TextMatrix(i, 0)
    End With
    
    With ListadoEscalaGanancias.dgEscalaGanancias
        i = .Row
        If i > 1 Then
            Borrar = MsgBox("Desea Borrar DEFINITIVAMENTE el Tramo Nro: " & .TextMatrix(i, 0) & " de la Norma: " & strNormaLegal & "?" & vbCrLf & "Tenga en cuenta que toda la información relacionada al mismo será igualmente eliminada", vbQuestion + vbYesNo, "BORRANDO TRAMO DE ESCALA")
            If Borrar = 6 Then
                Set rstRegistroSlave = New ADODB.Recordset
                SQL = "Select * from ESCALAAPLICABLEGANANCIAS" _
                & " Where TRAMO = '" & .TextMatrix(i, 0) _
                & "' And NORMALEGAL = '" & strNormaLegal & "'"
                rstRegistroSlave.Open SQL, dbSlave, adOpenForwardOnly, adLockOptimistic
                rstRegistroSlave.Delete
                rstRegistroSlave.Close
                Set rstRegistroSlave = Nothing
                ConfigurardgEscalaGanancias
                CargardgEscalaGanancias (ListadoEscalaGanancias.dgNormasEscalaGanancias.TextMatrix(1, 0))
            End If
        End If
        i = ""
    End With
    
End Sub

Public Sub EliminarLimitesDeducciones()

    Dim i As String
    Dim Borrar As Integer
    Dim SQL As String
    
    With ListadoLimitesDeducciones.dgLimitesDeducciones
        i = .Row
        If i > 1 Then
            Borrar = MsgBox("Desea Borrar DEFINITIVAMENTE la Norma Legal: " & .TextMatrix(i, 0) & "?" & vbCrLf & "Tenga en cuenta que toda la información relacionada a la misma será igualmente eliminada", vbQuestion + vbYesNo, "BORRANDO NORMA LEGAL DE DEDUCCIÓN")
            If Borrar = 6 Then
                Set rstRegistroSlave = New ADODB.Recordset
                SQL = "Select * from DEDUCCIONES4TACATEGORIA Where NORMALEGAL = " & "'" & .TextMatrix(i, 0) & "'"
                rstRegistroSlave.Open SQL, dbSlave, adOpenForwardOnly, adLockOptimistic
                rstRegistroSlave.Delete
                rstRegistroSlave.Close
                Set rstRegistroSlave = Nothing
                ConfigurardgLimitesDeducciones
                CargardgLimitesDeducciones
            End If
        End If
        i = ""
    End With
    
End Sub

Public Sub EliminarParentesco()

    Dim i As String
    Dim Borrar As Integer
    Dim SQL As String
    
    With ListadoParentesco.dgParentesco
        i = .Row
        If i <> 0 Then
            Borrar = MsgBox("Desea Borrar DEFINITIVAMENTE al Código de Parentesco: " & .TextMatrix(i, 0) & "?" & vbCrLf & "Tenga en cuenta que toda la información relacionada al mismo será igualmente eliminada", vbQuestion + vbYesNo, "BORRANDO PARENTESCO")
            If Borrar = 6 Then
                Set rstRegistroSlave = New ADODB.Recordset
                SQL = "Select * from ASIGNACIONESFAMILIARES Where CODIGO = " & "'" & .TextMatrix(i, 0) & "'"
                rstRegistroSlave.Open SQL, dbSlave, adOpenForwardOnly, adLockOptimistic
                rstRegistroSlave.Delete
                rstRegistroSlave.Close
                Set rstRegistroSlave = Nothing
                ConfigurardgParentesco
                CargardgParentesco
            End If
        End If
        i = ""
    End With
    
End Sub

Public Sub EliminarDeduccionesGenerales()

    Dim i As String
    Dim strPuestoLaboral As String
    Dim Borrar As Integer
    Dim SQL As String
    
    
    With ListadoDeduccionesGenerales.dgDeduccionesGenerales
        i = .Row
        If i <> 0 Then
            strPuestoLaboral = ListadoDeduccionesGenerales.dgAgentes.Row
            strPuestoLaboral = ListadoDeduccionesGenerales.dgAgentes.TextMatrix(strPuestoLaboral, 2)
            Borrar = MsgBox("Desea Borrar DEFINITIVAMENTE las DEDUCCIONES GENERALES de FECHA: " & .TextMatrix(i, 0) & " correspondiente al PUESTO " & strPuestoLaboral & "?" & vbCrLf & "Tenga en cuenta que esto puede afectar futuras Retenciones de Ganancias a efectuar al Agente en cuestión", vbQuestion + vbYesNo, "BORRANDO DEDUCCIONES GENERALES")
            If Borrar = 6 Then
                Set rstRegistroSlave = New ADODB.Recordset
                SQL = "Select * from IMPORTEDEDUCCIONESGENERALES Where PUESTOLABORAL= '" & strPuestoLaboral & "' And FECHA = #" & Format(.TextMatrix(i, 0), "YYYY/MM/DD") & "#"
                rstRegistroSlave.Open SQL, dbSlave, adOpenForwardOnly, adLockOptimistic
                rstRegistroSlave.Delete
                rstRegistroSlave.Close
                Set rstRegistroSlave = Nothing
                ConfigurardgDeduccionesGenerales
                CargardgDeduccionesGenerales (strPuestoLaboral)
            End If
            strPuestoLaboral = ""
        End If
        i = ""
    End With
    
End Sub

Public Sub EliminarCodigoLiquidacion()

    Dim i As String
    Dim Borrar As Integer
    Dim SQL As String
    
    With ListadoCodigoLiquidaciones.dgCodigoLiquidacion
        i = .Row
        If i <> 0 Then
            Borrar = MsgBox("Desea Borrar DEFINITIVAMENTE al Código de Liquidación: " & .TextMatrix(i, 0) & "?" & vbCrLf & "Tenga en cuenta que toda la información relacionada al mismo será igualmente eliminada", vbQuestion + vbYesNo, "BORRANDO CÓDIGO DE LIQUIDACIÓN")
            If Borrar = 6 Then
                Set rstRegistroSlave = New ADODB.Recordset
                SQL = "Select * from CODIGOLIQUIDACIONES Where CODIGO = " & "'" & .TextMatrix(i, 0) & "'"
                rstRegistroSlave.Open SQL, dbSlave, adOpenForwardOnly, adLockOptimistic
                rstRegistroSlave.Delete
                rstRegistroSlave.Close
                Set rstRegistroSlave = Nothing
                ConfigurardgCodigoLiquidacion
                CargardgCodigoLiquidacion
            End If
        End If
        i = ""
    End With
    
End Sub

Public Sub EliminarFamiliar()

    Dim i As String
    Dim strPuestoLaboral As String
    Dim strAgente As String
    Dim strDNI As String
    Dim Borrar As Integer
    Dim SQL As String
    
    
    With ListadoFamiliares.dgFamiliares
        i = .Row
        If i <> 0 Then
            strPuestoLaboral = ListadoFamiliares.dgAgentes.Row
            strAgente = ListadoFamiliares.dgAgentes.TextMatrix(strPuestoLaboral, 1)
            strPuestoLaboral = ListadoFamiliares.dgAgentes.TextMatrix(strPuestoLaboral, 2)
            strDNI = .TextMatrix(i, 3)
            Borrar = MsgBox("Desea Borrar DEFINITIVAMENTE al Familiar: " & .TextMatrix(i, 1) & " del Agente: " & strAgente & "?" & vbCrLf & "Tenga en cuenta que se borrarán todos los datos del familiar en cuestión", vbQuestion + vbYesNo, "BORRANDO FAMILIAR")
            If Borrar = 6 Then
                Set rstRegistroSlave = New ADODB.Recordset
                SQL = "Select * from CARGASDEFAMILIA Where PUESTOLABORAL= '" & strPuestoLaboral & "' And DNI = '" & strDNI & "'"
                rstRegistroSlave.Open SQL, dbSlave, adOpenForwardOnly, adLockOptimistic
                rstRegistroSlave.Delete
                rstRegistroSlave.Close
                Set rstRegistroSlave = Nothing
                ConfigurardgFamiliares
                Call CargardgFamiliares(strPuestoLaboral)
            End If
            strPuestoLaboral = ""
        End If
        i = ""
    End With
    
End Sub

Public Sub EliminarRetencionGanancias()

    Dim strPuestoLaboral As String
    Dim strCodigoLiquidacion As String
    Dim strPeriodo As String
    Dim strAgente As String
    Dim i As String
    Dim X As String
    Dim Borrar As Integer
    Dim SQL As String
    
    
    With ListadoLiquidacionGanancias
        i = .dgAgentesRetenidos.Row
        X = .dgCodigosLiquidacionesGanancias.Row
        If i <> 0 Then
            If IsNumeric(ListadoLiquidacionGanancias.dgAgentesRetenidos.TextMatrix(i, 2)) = True Then
                strPuestoLaboral = .dgAgentesRetenidos.TextMatrix(i, 0)
                strAgente = .dgAgentesRetenidos.TextMatrix(i, 1)
                strCodigoLiquidacion = .dgCodigosLiquidacionesGanancias.TextMatrix(X, 0)
                strPeriodo = .dgCodigosLiquidacionesGanancias.TextMatrix(X, 1)
                Borrar = MsgBox("Desea Borrar DEFINITIVAMENTE la Retención de Ganancias del: " & strPeriodo & " del Agente: " & strAgente & "?", vbQuestion + vbYesNo, "BORRANDO RETENCIÓN")
                If Borrar = 6 Then
                    Set rstRegistroSlave = New ADODB.Recordset
                    SQL = "Select * from LIQUIDACIONGANANCIAS4TACATEGORIA Where PUESTOLABORAL= '" & strPuestoLaboral & "' And CODIGOLIQUIDACION = '" & strCodigoLiquidacion & "'"
                    rstRegistroSlave.Open SQL, dbSlave, adOpenForwardOnly, adLockOptimistic
                    rstRegistroSlave.Delete
                    rstRegistroSlave.Close
                    Set rstRegistroSlave = Nothing
                    ConfigurardgAgentesRetenidos
                    Call CargardgAgentesRetenidos(strCodigoLiquidacion)
                End If
                strPuestoLaboral = ""
                strAgente = ""
                strCodigoLiquidacion = ""
                strPeriodo = ""
            Else
                MsgBox "No se puede eliminar la retención de un Agente SUGERIDO", vbCritical + vbOKOnly, "IMPOSIBLE ELIMINAR AGENTE SUGERIDO"
            End If
        End If
        i = ""
        X = ""
    End With
    
End Sub

Public Sub EliminarLiquidacionSISPER()

    Dim i As String
    Dim Borrar As Integer
    Dim SQL As String
    
    With ListadoLiquidacionesSISPER.dgCodigoLiquidacion
        i = .Row
        If i <> 0 Then
            If .TextMatrix(i, 2) = 0 Then
                MsgBox "La liquidación seleccionada no tiene datos para eliminar", vbCritical + vbOKOnly, "LIQUIDACIÓN SISPER VACÍA"
            Else
                Borrar = MsgBox("Desea Borrar DEFINITIVAMENTE Liquidación de Haberes " & .TextMatrix(i, 1) & "?" & vbCrLf & "Tenga en cuenta que toda la información relacionada al mismo será igualmente eliminada", vbQuestion + vbYesNo, "BORRANDO LIQUIDACIÓN SISPER")
                If Borrar = 6 Then
                    SQL = "DELETE FROM LIQUIDACIONSUELDOS Where CODIGOLIQUIDACION = " & "'" & .TextMatrix(i, 0) & "'"
                    dbSlave.BeginTrans
                    dbSlave.Execute SQL
                    dbSlave.CommitTrans
                    ConfigurardgLiquidacionSISPER
                    CargardgLiquidacionSISPER
                End If
            End If
        End If
        i = ""
    End With
    
End Sub

Public Sub EliminarAutocarga()

    Dim i As String
    Dim Borrar As Integer
    Dim SQL As String
    
    With Autocarga.dgListadoAutocarga
        i = .Row
        If i <> 0 Then
            Borrar = MsgBox("Desea Borrar DEFINITIVAMENTE el comprobante de Autocarga de fecha " & .TextMatrix(i, 1) & "?" & vbCrLf & "Tenga en cuenta que toda la información relacionada al mismo será igualmente eliminada", vbQuestion + vbYesNo, "BORRANDO LIQUIDACIÓN AUTOCARGA")
            If Borrar = 6 Then
                SQL = "DELETE FROM LIQUIDACIONHONORARIOS Where COMPROBANTE = " & "'" & .TextMatrix(i, 0) & "'"
                dbSlave.BeginTrans
                dbSlave.Execute SQL
                dbSlave.CommitTrans
                ConfigurardgAutocarga
                CargardgAutocarga
            End If
        End If
        i = ""
    End With
    
End Sub

Public Sub EliminarComprobanteSIIF()

    Dim i As String
    Dim Borrar As Integer
    Dim SQL As String
    Dim X As Integer
    Dim strComprobante As String
    
    With ListadoComprobantesSIIF.dgListadoComprobante
        i = .Row
        If i <> 0 Then
            Borrar = MsgBox("Desea Borrar DEFINITIVAMENTE el Comprobante SIIF Nro: " & .TextMatrix(i, 0) & "?" & vbCrLf & "Tenga en cuenta que podrá volver a reimputar el mismo desde Autocarga", vbQuestion + vbYesNo, "BORRANDO COMPROBANTE SIIF")
            If Borrar = 6 Then
                For X = 1 To 99
                    strComprobante = "NoSIIF" & Format(X, "00")
                    SQL = "Select * From LIQUIDACIONHONORARIOS" _
                    & " Where COMPROBANTE = '" & strComprobante & "'"
                    If SQLNoMatch(SQL) = True Then
                        Exit For
                    End If
                Next X
                'Insertamos una copia del grupo de registros a editar cambiando el Nro. de Comprobantey Tipo
                SQL = "Insert Into LIQUIDACIONHONORARIOS (Comprobante, Fecha, Tipo," _
                & " Proveedor, MontoBruto, Sellos, LibramientoPago, IIBB, OtraRetencion, Anticipo, Seguro, Descuento, Actividad, Partida)" _
                & " Select '" & strComprobante & "', Fecha, 0," _
                & " Proveedor, MontoBruto, Sellos, LibramientoPago, IIBB, OtraRetencion, Anticipo, Seguro, Descuento, 0, 0" _
                & " From LIQUIDACIONHONORARIOS Where Comprobante = " & "'" & .TextMatrix(i, 0) & "'"
                Debug.Print SQL
                dbSlave.BeginTrans
                dbSlave.Execute SQL
                dbSlave.CommitTrans
                'Borramos los registros que debían editarse
                SQL = "DELETE FROM LIQUIDACIONHONORARIOS Where COMPROBANTE = " & "'" & .TextMatrix(i, 0) & "'"
                dbSlave.BeginTrans
                dbSlave.Execute SQL
                dbSlave.CommitTrans
                ConfigurardgComprobantesSIIF
                Call CargardgComprobantesSIIF(, Year(Now()))
                X = .Row
                ConfigurardgImputacion
                CargardgImputacion (.TextMatrix(X, 0))
                ConfigurardgRetencion
                CargardgRetencion (.TextMatrix(X, 0))
            End If
        End If
        i = ""
    End With
    
End Sub

Public Sub EliminarHaberLiquidado()

    Dim i As String
    Dim Borrar As Integer
    Dim SQL As String
    Dim strPL As String
    Dim strCL As String
    Dim dblImporteEliminado As Double
    
    With ReciboDeSueldo.dgHaberesLiquidados
        i = .Row
        If i <> 0 Then
            Borrar = MsgBox("Desea Borrar DEFINITIVAMENTE el Concepto: " & .TextMatrix(i, 1) & " del recibo seleccionado?", vbQuestion + vbYesNo, "BORRANDO CONCEPTO")
            If Borrar = 6 Then
                'Buscamos el Puesto Laboral
                SQL = "Select PUESTOLABORAL From AGENTES" _
                & " Where NOMBRECOMPLETO = '" & ReciboDeSueldo.cmbAgente.Text & "'"
                Set rstBuscarSlave = New ADODB.Recordset
                rstBuscarSlave.Open SQL, dbSlave, adOpenDynamic, adLockOptimistic
                strPL = rstBuscarSlave!PuestoLaboral
                strCL = Left(ReciboDeSueldo.cmbPeriodo.Text, 4)
                rstBuscarSlave.Close
                Set rstBuscarSlave = Nothing
                Set rstRegistroSlave = New ADODB.Recordset
                SQL = "Select * from LIQUIDACIONSUELDOS " _
                & "Where PUESTOLABORAL = " & "'" & strPL & "'" _
                & "And CODIGOCONCEPTO = " & "'" & .TextMatrix(i, 0) & "'" _
                & "And CODIGOLIQUIDACION = " & "'" & strCL & "'"
                rstRegistroSlave.Open SQL, dbSlave, adOpenForwardOnly, adLockOptimistic
                dblImporteEliminado = rstRegistroSlave!Importe
                rstRegistroSlave.Delete
                rstRegistroSlave.Close
                'Ajustamos el Haber Óptimo
                SQL = "Select * from LIQUIDACIONSUELDOS " _
                & "Where PUESTOLABORAL = " & "'" & strPL & "'" _
                & "And CODIGOCONCEPTO = " & "'" & "9998" & "'" _
                & "And CODIGOLIQUIDACION = " & "'" & strCL & "'"
                rstRegistroSlave.Open SQL, dbSlave, adOpenForwardOnly, adLockOptimistic
                rstRegistroSlave!Importe = rstRegistroSlave!Importe - dblImporteEliminado
                rstRegistroSlave.Update
                rstRegistroSlave.Close
                'Ajustamos la Jubilación Personal
                SQL = "Select * from LIQUIDACIONSUELDOS " _
                & "Where PUESTOLABORAL = " & "'" & strPL & "'" _
                & "And CODIGOCONCEPTO = " & "'" & "0208" & "'" _
                & "And CODIGOLIQUIDACION = " & "'" & strCL & "'"
                rstRegistroSlave.Open SQL, dbSlave, adOpenForwardOnly, adLockOptimistic
                rstRegistroSlave!Importe = rstRegistroSlave!Importe - (dblImporteEliminado * 0.185)
                rstRegistroSlave.Update
                rstRegistroSlave.Close
                'Ajustamos la Jubilación Estatal
                SQL = "Select * from LIQUIDACIONSUELDOS " _
                & "Where PUESTOLABORAL = " & "'" & strPL & "'" _
                & "And CODIGOCONCEPTO = " & "'" & "0209" & "'" _
                & "And CODIGOLIQUIDACION = " & "'" & strCL & "'"
                rstRegistroSlave.Open SQL, dbSlave, adOpenForwardOnly, adLockOptimistic
                rstRegistroSlave!Importe = rstRegistroSlave!Importe - (dblImporteEliminado * 0.185)
                rstRegistroSlave.Update
                rstRegistroSlave.Close
                'Ajustamos la O.Social Personal
                SQL = "Select * from LIQUIDACIONSUELDOS " _
                & "Where PUESTOLABORAL = " & "'" & strPL & "'" _
                & "And CODIGOCONCEPTO = " & "'" & "0212" & "'" _
                & "And CODIGOLIQUIDACION = " & "'" & strCL & "'"
                rstRegistroSlave.Open SQL, dbSlave, adOpenForwardOnly, adLockOptimistic
                rstRegistroSlave!Importe = rstRegistroSlave!Importe - (dblImporteEliminado * 0.05)
                rstRegistroSlave.Update
                rstRegistroSlave.Close
                'Ajustamos la O.Social Estatal
                SQL = "Select * from LIQUIDACIONSUELDOS " _
                & "Where PUESTOLABORAL = " & "'" & strPL & "'" _
                & "And CODIGOCONCEPTO = " & "'" & "0213" & "'" _
                & "And CODIGOLIQUIDACION = " & "'" & strCL & "'"
                rstRegistroSlave.Open SQL, dbSlave, adOpenForwardOnly, adLockOptimistic
                rstRegistroSlave!Importe = rstRegistroSlave!Importe - (dblImporteEliminado * 0.04)
                rstRegistroSlave.Update
                rstRegistroSlave.Close

                Set rstRegistroSlave = Nothing
                ConfigurardgHaberesLiquidados
                ConfigurardgDescuentosLiquidados
                Call CargardgHaberesLiquidados(strPL, strCL)
                Call CargardgDescuentosLiquidados(strPL, strCL)
            End If
        End If
        i = ""
    End With
    
    SQL = ""
    strCL = ""
    strPL = ""
    
End Sub

Public Sub EliminarPrecarizadoImputado()

    Dim i As String
    Dim Borrar As Integer
    Dim SQL As String
    
    With ListadoHonorariosImputados.dgListadoHonorariosImputados
        i = .Row
        If i <> 0 Then
            Borrar = MsgBox("Desea Borrar DEFINITIVAMENTE la IMPUTACION del Facturero: " & .TextMatrix(i, 1) & " con monto: $" & .TextMatrix(i, 2) & " ?", vbQuestion + vbYesNo, "BORRANDO IMPUTACION")
            If Borrar = 6 Then
                Set rstRegistroSlave = New ADODB.Recordset
                SQL = "Select * from LIQUIDACIONHONORARIOS " _
                & "Where PROVEEDOR = " & "'" & .TextMatrix(i, 1) & "' " _
                & "And COMPROBANTE = " & "'" & ListadoHonorariosImputados.txtComprobante.Text & "' " _
                & "And MONTOBRUTO = " & De_Txt_a_Num_01(.TextMatrix(i, 2), 2, ".")
                rstRegistroSlave.Open SQL, dbSlave, adOpenForwardOnly, adLockOptimistic
                rstRegistroSlave.Delete
                rstRegistroSlave.Close
                Set rstRegistroSlave = Nothing
                ConfigurardgListadoHonorariosImputados
                Call CargardgListadoHonorariosImputados(ListadoHonorariosImputados.txtComprobante.Text)
            End If
        End If
        i = ""
    End With
    
End Sub

Public Sub EliminarCodigoSIRADIG(Tabla As String)

    Dim i As String
    Dim Borrar As Integer
    Dim SQL As String
    
    With ListadoCodigosSIRADIG.dgCodigosSIRADIG
        i = .Row
        If i <> 0 Then
            Borrar = MsgBox("Desea Borrar DEFINITIVAMENTE el Codigo: " & .TextMatrix(i, 0) & " - " & .TextMatrix(i, 1) & "?" _
            & vbCrLf & "Tenga en cuenta que toda la información relacionada al mismo será igualmente eliminada", vbQuestion + vbYesNo, "BORRANDO " _
            & ListadoCodigosSIRADIG.Caption)
            If Borrar = 6 Then
                Set rstRegistroSlave = New ADODB.Recordset
                SQL = "Select * from " & Tabla & " Where CODIGO = " & "'" & .TextMatrix(i, 0) & "'"
                rstRegistroSlave.Open SQL, dbSlave, adOpenForwardOnly, adLockOptimistic
                rstRegistroSlave.Delete
                rstRegistroSlave.Close
                Set rstRegistroSlave = Nothing
                ConfigurardgCodigosSIRADIG
                Call CargardgCodigosSIRADIG(strListadoCodigoSIRADIG)
            End If
        End If
        i = ""
    End With
    
End Sub

Public Sub EliminarF572()

    Dim i As String
    Dim Borrar As Integer
    Dim SQL As String
    Dim strCUIL As String
    
    With ListadoF572.dgPresentacionesF572
        i = .Row
        If i <> 0 And .TextMatrix(i, 0) <> "" Then
            Borrar = MsgBox("Desea Borrar DEFINITIVAMENTE la DDJJ del Agente: " & ListadoF572.cmbAgente.Text _
            & " , Periodo: " & .TextMatrix(i, 1) & "(" & .TextMatrix(i, 2) & ") ?" _
            & vbCrLf & "Tenga en cuenta que toda la información relacionada al mismo será igualmente eliminada", vbQuestion + vbYesNo, "BORRANDO DDJJ")
            If Borrar = 6 Then
                SQL = "Delete From DeduccionesGeneralesSIRADIG " _
                    & "Where ID = " & " '" & .TextMatrix(i, 0) & "'"
                dbSlave.BeginTrans
                dbSlave.Execute SQL
                dbSlave.CommitTrans
                SQL = "Delete From CargasFamiliaSIRADIG " _
                & "Where ID = " & " '" & .TextMatrix(i, 0) & "'"
                dbSlave.BeginTrans
                dbSlave.Execute SQL
                dbSlave.CommitTrans
                SQL = "Delete From PresentacionSIRADIG " _
                & "Where ID = " & " '" & .TextMatrix(i, 0) & "'"
                dbSlave.BeginTrans
                dbSlave.Execute SQL
                dbSlave.CommitTrans
                strCUIL = "Select CUIL From AGENTES" _
                & " Where NOMBRECOMPLETO = '" & ListadoF572.cmbAgente.Text & "'"
                Set rstBuscarSlave = New ADODB.Recordset
                rstBuscarSlave.Open strCUIL, dbSlave, adOpenDynamic, adLockOptimistic
                strCUIL = rstBuscarSlave!CUIL
                rstBuscarSlave.Close
                Set rstBuscarSlave = Nothing
                ConfigurardgPresentacionesF572
                Call CargardgPresentacionesF572(strCUIL)
                ConfigurardgCargasDeFamiliaF572
                ConfigurardgDeduccionesGeneralesF572
                If .Rows > 1 Then
                    Call CargardgCargasDeFamiliaF572(.TextMatrix(1, 0))
                    Call CargardgDeduccionesGeneralesF572(.TextMatrix(1, 0))
                End If
            End If
        End If
        i = ""
    End With
    
End Sub
