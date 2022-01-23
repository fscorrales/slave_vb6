Attribute VB_Name = "CopiarRegistros"
'VER EL SIGUIENTE LINK

'http://stackoverflow.com/questions/9225038/copy-data-to-and-from-the-same-table-and-change-the-value-of-copied-data-in-one

'INSERT INTO Metric(Key,Name,MetricValue)
'SELECT 387,Name,MetricValue
'From Metric
'Where Key = 112

'Another option
'Db = "C:\Docs\ltd.mdb"

'Set cn = CreateObject("ADODB.Connection")

'cn.Open "Provider = Microsoft.Jet.OLEDB.4.0; " & _
   "Data Source =" & Db

'sSQL = "INSERT INTO New (id,schedno) " _
'& "SELECT id,schedno FROM [new.txt] IN '' " _
'& "'text;HDR=Yes;FMT=Delimited;database=C:\Docs\';"

'cn.Execute sSQL

Public Sub LiquidacionOrigenALiquidacionDestino(CodigoOrigen As String, CodigoDestino As String)

    Dim SQL As String
    'Dim qdf As QueryDef
    
    
    
    'Conectamos Recordset Origen de Datos
    'SQL = "Select * From LIQUIDACIONSUELDOS Where CODIGOLIQUIDACION = " & "'" & CodigoOrigen & "'"
    'Set rstListadoSlave = New ADODB.Recordset
    'rstListadoSlave.Open SQL, dbSlave, adOpenDynamic, adLockOptimistic
    'Conectamos Recordset Destino de Datos
    'Set rstRegistroSlave = New ADODB.Recordset
    'rstRegistroSlave.Open "LIQUIDACIONSUELDOS", dbSlave, adOpenForwardOnly, adLockOptimistic
    'If rstListadoSlave.BOF = False Then
    '    rstListadoSlave.MoveFirst
    '    While rstListadoSlave.EOF = False
    '        rstRegistroSlave.AddNew
    '        rstRegistroSlave!CodigoLiquidacion = CodigoDestino
    '        rstRegistroSlave!PuestoLaboral = rstListadoSlave!PuestoLaboral
    '        rstRegistroSlave!CodigoConcepto = rstListadoSlave!CodigoConcepto
    '        rstRegistroSlave!Importe = rstListadoSlave!Importe
    '        rstRegistroSlave.Update
    '        'rstRegistroSlave.Close
    '        rstListadoSlave.MoveNext
    '    Wend
    'End If

    'Descargamos todas las variables usadas
    'rstListadoSlave.Close
    'Set rstListadoSlave = Nothing
    'rstRegistroSlave.Close
    'Set rstRegistroSlave = Nothing
    'SQL = ""
    
    'SQL INSERT INTO SELECT Statement
    SQL = "Insert Into LIQUIDACIONSUELDOS (CodigoLiquidacion, PuestoLaboral, CodigoConcepto, Importe) " & _
    "Select '" & CodigoDestino & "', PuestoLaboral, CodigoConcepto, Importe " & _
    "From LIQUIDACIONSUELDOS Where CodigoLiquidacion = " & "'" & CodigoOrigen & "'"
    'Debug.Print SQL
    
    dbSlave.BeginTrans
    dbSlave.Execute SQL
    dbSlave.CommitTrans
    
    
    'Change connection from ADO to DAO
    'Desconectar
    'ConectarDAO
    
    'Create a dummy QueryDef object.
    'Set qdf = dbSlaveDAO.CreateQueryDef("", "Select * from LIQUIDACIONSUELDOS")
    
    'SQL INSERT INTO SELECT Statement
    'SQL = "Insert Into LIQUIDACIONSUELDOS (CodigoLiquidacion, PuestoLaboral, CodigoConcepto, Importe) " & _
    '"Select CodigoLiquidacion, PuestoLaboral, CodigoConcepto, Importe " & _
    '"From LIQUIDACIONSUELDOS Where CodigoLiquidacion = " & "'" & CodigoOrigen & "'"
    'Debug.Print SQL
    'qdf.SQL = SQL
    'qdf.Execute
        
    'Unload variable
    'qdf.Close
    'Set qdf = Nothing
    'SQL = ""
    
    'Change connection from DAO to ADO
    'DesconectarDAO
    'Conectar
    
End Sub

Public Sub MigrarDeduccionesSIRADIG(PeriodoOrigen As String, PeriodoDestino As String, TipoDatos As String)

    Dim SQL As String
    Dim intContar As Integer
    Dim strID As String
    Dim datFechaControl As Date
    Dim intMesHasta As Integer
    Dim strCodigoParentescoSIRADIG As String

    
    Call InfoGeneral("- INICIANDO MIGRACIÓN DATOS F. 572 Web -", MigrarDeducciones)
    Call InfoGeneral("", MigrarDeducciones)
    
    If PeriodoOrigen = "BD Previa" Then
        'Precargamos las DDJJ en período origen
        SQL = "Select Agentes.CUIL From " _
            & "Agentes Inner Join CargasDeFamilia On " _
            & "Agentes.PuestoLaboral = CargasDeFamilia.PuestoLaboral " _
            & "Where DeducibleGanancias = True " _
            & "Group by Agentes.CUIL"
        Set rstRegistroSlave = New ADODB.Recordset
        rstRegistroSlave.Open SQL, dbSlave, adOpenForwardOnly, adLockOptimistic
        Call InfoGeneral("Cantidad de Cargas de Familia en Origen: " & rstRegistroSlave.RecordCount, MigrarDeducciones)
        Call InfoGeneral("", MigrarDeducciones)
        'Bucle para migrar las DDJJ
        If rstRegistroSlave.BOF = False Then
            rstRegistroSlave.MoveFirst
            intContar = 0
            While rstRegistroSlave.EOF = False
                'Verificamos si la Carga de Familia ya no esta cargada en periodo destino
                SQL = "Select * From PresentacionSIRADIG " _
                    & "Where Right(PresentacionSIRADIG.ID,2) = '" & Right(PeriodoDestino, 2) & "' " _
                    & "And PresentacionSIRADIG.CUIL = '" & rstRegistroSlave!CUIL & "' " _
                    & "And PresentacionSIRADIG.NroPresentacion = '0'"
                If SQLNoMatch(SQL) = True Then
                    'Si no existe, procedemos a importar
                    'Buscamos el último número de ID utilizado en el Período Destino
                    strID = ""
                    SQL = "SELECT MAX(Left(ID,5)) AS LastID FROM PresentacionSIRADIG " _
                        & "WHERE RIGHT(ID,2) = '" & Right(PeriodoDestino, 2) & "'"
                    If SQLNoMatch(SQL) = True Then
                        strID = "00001/" & Right(PeriodoDestino, 2)
                    Else
                        Set rstBuscarSlave = New ADODB.Recordset
                        rstBuscarSlave.Open SQL, dbSlave, adOpenForwardOnly, adLockOptimistic
                        strID = CStr(Val(rstBuscarSlave!LastID) + 1)
                        strID = Format(strID, "00000") & "/" & Right(PeriodoDestino, 2)
                        rstBuscarSlave.Close
                        Set rstBuscarSlave = Nothing
                    End If
                    'Cargamos la DDJJ en la BD destino
                    SQL = "INSERT INTO PresentacionSIRADIG " _
                        & "(ID, CUIL, Fecha, NroPresentacion) " _
                        & "VALUES( '" & strID & "' , '" & rstRegistroSlave!CUIL _
                        & "' , #" & Format("01/01/" & PeriodoDestino, "MM/DD/YYYY") _
                        & "# , '0')"
                    dbSlave.BeginTrans
                    dbSlave.Execute SQL
                    dbSlave.CommitTrans
                    'Migramos las Cargas de Familia de la BD origen a la BD destino
                    SQL = "Select CargasDeFamilia.DNI, CargasDeFamilia.FechaAlta, " _
                        & "CargasDeFamilia.CodigoParentesco, CargasDeFamilia.Discapacitado From " _
                        & "Agentes Inner Join CargasDeFamilia On " _
                        & "Agentes.PuestoLaboral = CargasDeFamilia.PuestoLaboral " _
                        & "Where DeducibleGanancias = True " _
                        & "And Agentes.CUIL = '" & rstRegistroSlave!CUIL & "'"
                    Set rstListadoSlave = New ADODB.Recordset
                    rstListadoSlave.Open SQL, dbSlave, adOpenForwardOnly, adLockOptimistic
                    If rstListadoSlave.BOF = False Then
                        rstListadoSlave.MoveFirst
                        While rstListadoSlave.EOF = False
                            datFechaControl = DateTime.DateSerial(PeriodoDestino, "01", "01")
                            strCodigoParentescoSIRADIG = EquipararCodigoParentescoSISPERconSIRADIG(rstListadoSlave!CodigoParentesco, _
                            rstListadoSlave!Discapacitado)
                            intMesHasta = MesHastaCargaDeFamiliaDeducibleSIRADIG(strCodigoParentescoSIRADIG, _
                            rstListadoSlave!FechaAlta, datFechaControl)
                            If intMesHasta > 0 Then
                                Set rstBuscarSlave = New ADODB.Recordset
                                rstBuscarSlave.Open "CargasFamiliaSIRADIG", dbSlave, adOpenForwardOnly, adLockOptimistic
                                rstBuscarSlave.AddNew
                                rstBuscarSlave!ID = strID
                                rstBuscarSlave!CUIL = "00" & rstListadoSlave!DNI & "0"
                                rstBuscarSlave!MesDesde = "1"
                                rstBuscarSlave!MesHasta = intMesHasta
                                rstBuscarSlave!CodigoParentesco = strCodigoParentescoSIRADIG
                                If intMesHasta = 12 Then
                                    rstBuscarSlave!ProximoPeriodo = -1
                                Else
                                    rstBuscarSlave!ProximoPeriodo = 0
                                End If
                                rstBuscarSlave!FechaNacimiento = rstListadoSlave!FechaAlta
                                rstBuscarSlave.Update
                                rstBuscarSlave.Close
                                Set rstBuscarSlave = Nothing
                            End If
                            rstListadoSlave.MoveNext
                        Wend
                    End If
                    rstListadoSlave.Close
                    Set rstListadoSlave = Nothing
                    'Contamos las DDJJ importadas
                    intContar = intContar + 1
                End If
                rstRegistroSlave.MoveNext
            Wend
            Call InfoGeneral("Cantidad de DDJJ para importar (falta procedimiento): " & intContar, MigrarDeducciones)
            Call InfoGeneral("", MigrarDeducciones)
        End If
        rstRegistroSlave.Close
        Set rstRegistroSlave = Nothing
    Else
        Select Case TipoDatos
        Case "Todas las Deducciones"
            'Precargamos las DDJJ en período origen
            SQL = "Select * From PresentacionSIRADIG " _
                & "Where Right(ID,2) = " & "'" & Right(PeriodoOrigen, 2) & "'"
            Set rstRegistroSlave = New ADODB.Recordset
            rstRegistroSlave.Open SQL, dbSlave, adOpenForwardOnly, adLockOptimistic
            Call InfoGeneral("Cantidad de DDJJ en Origen: " & rstRegistroSlave.RecordCount, MigrarDeducciones)
            Call InfoGeneral("", MigrarDeducciones)
            'Bucle para migrar las DDJJ
            If rstRegistroSlave.BOF = False Then
                rstRegistroSlave.MoveFirst
                intContar = 0
                While rstRegistroSlave.EOF = False
                    'Verificamos si no existe una migración previa en la BD Destino (NroPresentación = 0)
                    SQL = "Select * From PresentacionSIRADIG " _
                        & "Where Right(ID,2) = " & "'" & Right(PeriodoDestino, 2) & "' " _
                        & "And CUIL = '" & rstRegistroSlave!CUIL & "' " _
                        & "And NroPresentacion = '0' "
                    If SQLNoMatch(SQL) = True Then
                        'Si no existe, procedemos a importar
                        'Buscamos el último número de ID utilizado en el Período Destino
                        strID = ""
                        SQL = "SELECT MAX(Left(ID,5)) AS LastID FROM PresentacionSIRADIG " _
                            & "WHERE RIGHT(ID,2) = '" & Right(PeriodoDestino, 2) & "'"
                        If SQLNoMatch(SQL) = True Then
                            strID = "00001/" & Right(PeriodoDestino, 2)
                        Else
                            Set rstBuscarSlave = New ADODB.Recordset
                            rstBuscarSlave.Open SQL, dbSlave, adOpenForwardOnly, adLockOptimistic
                            strID = CStr(Val(rstBuscarSlave!LastID) + 1)
                            strID = Format(strID, "00000") & "/" & Right(PeriodoDestino, 2)
                            rstBuscarSlave.Close
                            Set rstBuscarSlave = Nothing
                        End If
                        'Cargamos la DDJJ en la BD destino
                        SQL = "INSERT INTO PresentacionSIRADIG " _
                            & "(ID, CUIL, Fecha, NroPresentacion) " _
                            & "VALUES( '" & strID & "' , '" & rstRegistroSlave!CUIL _
                            & "' , #" & Format("01/01/" & PeriodoDestino, "MM/DD/YYYY") _
                            & "# , '0')"
                        dbSlave.BeginTrans
                        dbSlave.Execute SQL
                        dbSlave.CommitTrans
                        'Migramos las Deducciones Generales de la BD origen a la BD destino
                        SQL = "Insert Into DeduccionesGeneralesSIRADIG (ID, CodigoDeduccion, Mes01, Mes02, Mes03, Mes04, Mes05, Mes06, Mes07, Mes08, Mes09, Mes10, Mes11, Mes12) " _
                            & "Select '" & strID & "', CodigoDeduccion, Mes01, Mes02, Mes03, Mes04, Mes05, Mes06, Mes07, Mes08, Mes09, Mes10, Mes11, Mes12 " _
                            & "From DeduccionesGeneralesSIRADIG " _
                            & "Where ID = '" & rstRegistroSlave!ID & "'"
                        'Debug.Print SQL
                        dbSlave.BeginTrans
                        dbSlave.Execute SQL
                        dbSlave.CommitTrans
                        'Migramos las Cargas de Familia de la BD origen a la BD destino
                        SQL = "Insert Into CargasFamiliaSIRADIG (ID, CUIL, MesDesde, MesHasta, CodigoParentesco, ProximoPeriodo, FechaNacimiento) " _
                            & "Select '" & strID & "', CUIL, MesDesde, MesHasta, CodigoParentesco, ProximoPeriodo, FechaNacimiento " _
                            & "From CargasFamiliaSIRADIG " _
                            & "Where ID = '" & rstRegistroSlave!ID & "' " _
                            & "And ProximoPeriodo = True"
                        'Debug.Print SQL
                        dbSlave.BeginTrans
                        dbSlave.Execute SQL
                        dbSlave.CommitTrans
                        'Controlamos las Cargas de Familia recién importadas para ver si son deducibles en el período destino
                        datFechaControl = DateTime.DateSerial(PeriodoDestino, "01", "01")
                        intMesHasta = 0
                        SQL = "Select * From CargasFamiliaSIRADIG " _
                        & "Where ID = '" & strID & "'"
                        Set rstListadoSlave = New ADODB.Recordset
                        rstListadoSlave.Open SQL, dbSlave, adOpenForwardOnly, adLockOptimistic
                        If rstListadoSlave.BOF = False Then
                            rstListadoSlave.MoveFirst
                            While rstListadoSlave.EOF = False
                                intMesHasta = MesHastaCargaDeFamiliaDeducibleSIRADIG(rstListadoSlave!CodigoParentesco, _
                                rstListadoSlave!FechaNacimiento, datFechaControl)
                                Select Case intMesHasta
                                Case 0
                                    SQL = "Delete * from CargasFamiliaSIRADIG " _
                                    & "Where ID = '" & strID & "' " _
                                    & "And CUIL = '" & rstListadoSlave!CUIL & "'"
                                    dbSlave.BeginTrans
                                    dbSlave.Execute SQL
                                    dbSlave.CommitTrans
                                Case Is < 12
                                    SQL = "Update CargasFamiliaSIRADIG " _
                                    & "Set MesHasta = '" & intMesHasta & "' " _
                                    & "Where ID = '" & strID & "' " _
                                    & "And CUIL = '" & rstListadoSlave!CUIL & "'"
                                    dbSlave.BeginTrans
                                    dbSlave.Execute SQL
                                    dbSlave.CommitTrans
                                End Select
                                rstListadoSlave.MoveNext
                            Wend
                        End If
                        rstListadoSlave.Close
                        Set rstListadoSlave = Nothing
                        'Contamos las DDJJ importadas
                        intContar = intContar + 1
                    End If
                    rstRegistroSlave.MoveNext
                Wend
                Call InfoGeneral("Cantidad de DDJJ importadas: " & intContar, MigrarDeducciones)
                Call InfoGeneral("", MigrarDeducciones)
            End If
            rstRegistroSlave.Close
            Set rstRegistroSlave = Nothing
        Case "Solo Deducciones Personales"
            'Precargamos las DDJJ en período origen
            SQL = "Select * From PresentacionSIRADIG " _
                & "Where Right(ID,2) = " & "'" & Right(PeriodoOrigen, 2) & "'"
            Set rstRegistroSlave = New ADODB.Recordset
            rstRegistroSlave.Open SQL, dbSlave, adOpenForwardOnly, adLockOptimistic
            Call InfoGeneral("Cantidad de DDJJ en Origen: " & rstRegistroSlave.RecordCount, MigrarDeducciones)
            Call InfoGeneral("", MigrarDeducciones)
            'Bucle para migrar las DDJJ
            If rstRegistroSlave.BOF = False Then
                rstRegistroSlave.MoveFirst
                intContar = 0
                While rstRegistroSlave.EOF = False
                    'Verificamos si no existe una migración previa en la BD Destino (NroPresentación = 0)
                    SQL = "Select * From PresentacionSIRADIG " _
                        & "Where Right(ID,2) = " & "'" & Right(PeriodoDestino, 2) & "' " _
                        & "And CUIL = '" & rstRegistroSlave!CUIL & "' " _
                        & "And NroPresentacion = '0' "
                    If SQLNoMatch(SQL) = True Then
                        'Si no existe, procedemos a importar
                        'Buscamos el último número de ID utilizado en el Período Destino
                        strID = ""
                        SQL = "SELECT MAX(Left(ID,5)) AS LastID FROM PresentacionSIRADIG " _
                            & "WHERE RIGHT(ID,2) = '" & Right(PeriodoDestino, 2) & "'"
                        If SQLNoMatch(SQL) = True Then
                            strID = "00001/" & Right(PeriodoDestino, 2)
                        Else
                            Set rstBuscarSlave = New ADODB.Recordset
                            rstBuscarSlave.Open SQL, dbSlave, adOpenForwardOnly, adLockOptimistic
                            strID = CStr(Val(rstBuscarSlave!LastID) + 1)
                            strID = Format(strID, "00000") & "/" & Right(PeriodoDestino, 2)
                            rstBuscarSlave.Close
                            Set rstBuscarSlave = Nothing
                        End If
                        'Cargamos la DDJJ en la BD destino
                        SQL = "INSERT INTO PresentacionSIRADIG " _
                            & "(ID, CUIL, Fecha, NroPresentacion) " _
                            & "VALUES( '" & strID & "' , '" & rstRegistroSlave!CUIL _
                            & "' , #" & Format("01/01/" & PeriodoDestino, "MM/DD/YYYY") _
                            & "# , '0')"
                        dbSlave.BeginTrans
                        dbSlave.Execute SQL
                        dbSlave.CommitTrans
                        'Migramos las Cargas de Familia de la BD origen a la BD destino
                        SQL = "Insert Into CargasFamiliaSIRADIG (ID, CUIL, MesDesde, MesHasta, CodigoParentesco, ProximoPeriodo, FechaNacimiento) " _
                            & "Select '" & strID & "', CUIL, MesDesde, MesHasta, CodigoParentesco, ProximoPeriodo, FechaNacimiento " _
                            & "From CargasFamiliaSIRADIG " _
                            & "Where ID = '" & rstRegistroSlave!ID & "' " _
                            & "And ProximoPeriodo = True"
                        'Debug.Print SQL
                        dbSlave.BeginTrans
                        dbSlave.Execute SQL
                        dbSlave.CommitTrans
                        'Controlamos las Cargas de Familia recién importadas para ver si son deducibles en el período destino
                        datFechaControl = DateTime.DateSerial(PeriodoDestino, "01", "01")
                        intMesHasta = 0
                        SQL = "Select * From CargasFamiliaSIRADIG " _
                        & "Where ID = '" & strID & "'"
                        Set rstListadoSlave = New ADODB.Recordset
                        rstListadoSlave.Open SQL, dbSlave, adOpenForwardOnly, adLockOptimistic
                        If rstListadoSlave.BOF = False Then
                            rstListadoSlave.MoveFirst
                            While rstListadoSlave.EOF = False
                                intMesHasta = MesHastaCargaDeFamiliaDeducibleSIRADIG(rstListadoSlave!CodigoParentesco, _
                                rstListadoSlave!FechaNacimiento, datFechaControl)
                                Select Case intMesHasta
                                Case 0
                                    SQL = "Delete * from CargasFamiliaSIRADIG " _
                                    & "Where ID = '" & strID & "' " _
                                    & "And CUIL = '" & rstListadoSlave!CUIL & "'"
                                    dbSlave.BeginTrans
                                    dbSlave.Execute SQL
                                    dbSlave.CommitTrans
                                Case Is < 12
                                    SQL = "Update CargasFamiliaSIRADIG " _
                                    & "Set MesHasta = '" & intMesHasta & "' " _
                                    & "Where ID = '" & strID & "' " _
                                    & "And CUIL = '" & rstListadoSlave!CUIL & "'"
                                    dbSlave.BeginTrans
                                    dbSlave.Execute SQL
                                    dbSlave.CommitTrans
                                End Select
                                rstListadoSlave.MoveNext
                            Wend
                        End If
                        rstListadoSlave.Close
                        Set rstListadoSlave = Nothing
                        'Contamos las DDJJ importadas
                        intContar = intContar + 1
                    End If
                    rstRegistroSlave.MoveNext
                Wend
                Call InfoGeneral("Cantidad de DDJJ importadas: " & intContar, MigrarDeducciones)
                Call InfoGeneral("", MigrarDeducciones)
            End If
            rstRegistroSlave.Close
            Set rstRegistroSlave = Nothing
        Case "Solo Deducciones Generales"
            'Precargamos las DDJJ en período origen
            SQL = "Select * From PresentacionSIRADIG " _
                & "Where Right(ID,2) = " & "'" & Right(PeriodoOrigen, 2) & "'"
            Set rstRegistroSlave = New ADODB.Recordset
            rstRegistroSlave.Open SQL, dbSlave, adOpenForwardOnly, adLockOptimistic
            Call InfoGeneral("Cantidad de DDJJ en Origen: " & rstRegistroSlave.RecordCount, MigrarDeducciones)
            Call InfoGeneral("", MigrarDeducciones)
            'Bucle para migrar las DDJJ
            If rstRegistroSlave.BOF = False Then
                rstRegistroSlave.MoveFirst
                intContar = 0
                While rstRegistroSlave.EOF = False
                    'Verificamos si no existe una migración previa en la BD Destino (NroPresentación = 0)
                    SQL = "Select * From PresentacionSIRADIG " _
                        & "Where Right(ID,2) = " & "'" & Right(PeriodoDestino, 2) & "' " _
                        & "And CUIL = '" & rstRegistroSlave!CUIL & "' " _
                        & "And NroPresentacion = '0' "
                    If SQLNoMatch(SQL) = True Then
                        'Si no existe, procedemos a importar
                        'Buscamos el último número de ID utilizado en el Período Destino
                        strID = ""
                        SQL = "SELECT MAX(Left(ID,5)) AS LastID FROM PresentacionSIRADIG " _
                            & "WHERE RIGHT(ID,2) = '" & Right(PeriodoDestino, 2) & "'"
                        If SQLNoMatch(SQL) = True Then
                            strID = "00001/" & Right(PeriodoDestino, 2)
                        Else
                            Set rstBuscarSlave = New ADODB.Recordset
                            rstBuscarSlave.Open SQL, dbSlave, adOpenForwardOnly, adLockOptimistic
                            strID = CStr(Val(rstBuscarSlave!LastID) + 1)
                            strID = Format(strID, "00000") & "/" & Right(PeriodoDestino, 2)
                            rstBuscarSlave.Close
                            Set rstBuscarSlave = Nothing
                        End If
                        'Cargamos la DDJJ en la BD destino
                        SQL = "INSERT INTO PresentacionSIRADIG " _
                            & "(ID, CUIL, Fecha, NroPresentacion) " _
                            & "VALUES( '" & strID & "' , '" & rstRegistroSlave!CUIL _
                            & "' , #" & Format("01/01/" & PeriodoDestino, "MM/DD/YYYY") _
                            & "# , '0')"
                        dbSlave.BeginTrans
                        dbSlave.Execute SQL
                        dbSlave.CommitTrans
                        'Migramos las Deducciones Generales de la BD origen a la BD destino
                        SQL = "Insert Into DeduccionesGeneralesSIRADIG (ID, CodigoDeduccion, Mes01, Mes02, Mes03, Mes04, Mes05, Mes06, Mes07, Mes08, Mes09, Mes10, Mes11, Mes12) " _
                            & "Select '" & strID & "', CodigoDeduccion, Mes01, Mes02, Mes03, Mes04, Mes05, Mes06, Mes07, Mes08, Mes09, Mes10, Mes11, Mes12 " _
                            & "From DeduccionesGeneralesSIRADIG " _
                            & "Where ID = '" & rstRegistroSlave!ID & "'"
                        'Debug.Print SQL
                        dbSlave.BeginTrans
                        dbSlave.Execute SQL
                        dbSlave.CommitTrans
                        'Contamos las DDJJ importadas
                        intContar = intContar + 1
                    End If
                    rstRegistroSlave.MoveNext
                Wend
                Call InfoGeneral("Cantidad de DDJJ importadas: " & intContar, MigrarDeducciones)
                Call InfoGeneral("", MigrarDeducciones)
            End If
            rstRegistroSlave.Close
            Set rstRegistroSlave = Nothing
        End Select
    End If
    
    Call InfoGeneral("- MIGRACIÓN FINALIZADA -", MigrarDeducciones)
        
End Sub
