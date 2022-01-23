Attribute VB_Name = "ImportarDatos"
Public bolImportandoLiquidacionHonorariosSLAVE As Boolean
Public bolImportandoPrecarizadosSLAVE As Boolean

Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
 
Public Sub ImportarPadronFamiliaresV2()

    Dim strDireccion As String
    Dim strLineaImportada As String
    'Variable de tipo Aplicación de Excel
    Dim objExcel As Object
    'Una variable de tipo Libro de Excel
    Dim xLibro As Object
    Dim i As Integer
    Dim intPosicion As Integer
    Dim strPuestoLaboral As String
    Dim strParentesco As String
    Dim strNombreFamiliar As String
    Dim datFechaAlta As Date
    Dim strDNIFamiliar As String
    Dim strDiscapacidad As String
    Dim strNivelDeEstudio As String
    Dim strVive As String
    Dim strCobraSalario As String
    Dim strObraSocial As String
    Dim SQL As String
    
    On Error GoTo HayError
    
    With ImportacionPadronFamiliares
        .txtInforme.Text = ""
        Call InfoGeneral("- INICIANDO IMPORTACIÓN PADRÓN FAMILIARES -", ImportacionPadronFamiliares)
        Call InfoGeneral("", ImportacionPadronFamiliares)
        '.dlgMultifuncion.Filter = "Todos los Excel(*.csv)|*.csv|"
        .dlgMultifuncion.FileName = ""
        .dlgMultifuncion.ShowOpen
        strDireccion = .dlgMultifuncion.FileName
    End With
    If strDireccion = "" Then 'En caso de apretar cancelar
        Call InfoGeneral("IMPORTACION CANCELADA", ImportacionPadronFamiliares)
        Exit Sub
    Else
        'creamos un nuevo objeto excel
        Set objExcel = CreateObject("Excel.Application")
        'Usamos el método open para abrir el archivo que está _
         en el directorio del programa llamado archivo.csv
        objExcel.Workbooks.Open FileName:=(strDireccion)
        'Hacemos el Excel NO sea Visible
        Set xLibro = objExcel
        objExcel.Visible = False
        With xLibro.Sheets(1)
            i = 2
            Set rstRegistroSlave = New ADODB.Recordset
            While Trim(.Cells(i, 1)) <> ""
                strLineaImportada = .Cells(i, 1)
                If Left(strLineaImportada, 10) = "Documento:" Then
                    'Pamos de largo el Nro. Documento
                    intPosicion = InStr(strLineaImportada, ",,")
                    strLineaImportada = Right(strLineaImportada, Len(strLineaImportada) - intPosicion - 1)
                    intPosicion = InStr(strLineaImportada, ",")
                    strLineaImportada = Right(strLineaImportada, Len(strLineaImportada) - intPosicion)
                    'Pamos de largo el Nro. Agente
                    intPosicion = InStr(strLineaImportada, ",,")
                    strLineaImportada = Right(strLineaImportada, Len(strLineaImportada) - intPosicion - 1)
                    intPosicion = InStr(strLineaImportada, ",")
                    strLineaImportada = Right(strLineaImportada, Len(strLineaImportada) - intPosicion)
                    'Obtenemos el Puesto Laboral
                    intPosicion = InStr(strLineaImportada, ",")
                    strLineaImportada = Right(strLineaImportada, Len(strLineaImportada) - intPosicion)
                    intPosicion = InStr(strLineaImportada, ",")
                    strPuestoLaboral = Left(strLineaImportada, intPosicion - 1)
                    SQL = "Select * from AGENTES Where PUESTOLABORAL= '" & strPuestoLaboral & "'"
                    If SQLNoMatch(SQL) = True Then
                        'Procedemos a Buscar el Nro de Puesto a partir del nombre del agente
                        strLineaImportada = Right(strLineaImportada, Len(strLineaImportada) - intPosicion)
                        intPosicion = InStr(strLineaImportada, ",,,")
                        strLineaImportada = Right(strLineaImportada, Len(strLineaImportada) - intPosicion - 2)
                        intPosicion = InStr(strLineaImportada, ",,,,,,,")
                        strLineaImportada = Left(strLineaImportada, intPosicion - 1)
                        'Verificamos que exista el agente en la base principal
                        SQL = "Select * from AGENTES Where NOMBRECOMPLETO= '" & strLineaImportada & "'"
                        If SQLNoMatch(SQL) = False Then
                            rstRegistroSlave.Open SQL, dbSlave, adOpenForwardOnly, adLockOptimistic
                            strPuestoLaboral = rstRegistroSlave!PuestoLaboral
                            rstRegistroSlave.Close
                        Else
                            i = i + 1
                            strLineaImportada = .Cells(i, 1)
                            While Left(strLineaImportada, 10) <> "Documento:"
                                i = i + 1
                                strLineaImportada = .Cells(i, 1)
                            Wend
                            i = i - 1
                        End If
                    End If
                Else
                    'Pamos de largo el Nro. Familiar
                    intPosicion = InStr(strLineaImportada, ",")
                    strLineaImportada = Right(strLineaImportada, Len(strLineaImportada) - intPosicion)
                    'Obtenemos Descripcion de Parentesco
                    intPosicion = InStr(strLineaImportada, ",,,")
                    strParentesco = Left(strLineaImportada, intPosicion - 1)
                    SQL = "Select * from ASIGNACIONESFAMILIARES Where PARENTESCO = " & "'" & strParentesco & "'"
                    rstRegistroSlave.Open SQL, dbSlave, adOpenForwardOnly, adLockOptimistic
                    strParentesco = rstRegistroSlave!Codigo
                    rstRegistroSlave.Close
                    strLineaImportada = Right(strLineaImportada, Len(strLineaImportada) - intPosicion - 2)
                    'Obtenemos Nombre del Familiar
                    intPosicion = InStr(strLineaImportada, ",,,,,")
                    strNombreFamiliar = Left(strLineaImportada, intPosicion - 1)
                    strLineaImportada = Right(strLineaImportada, Len(strLineaImportada) - intPosicion - 4)
                    'Obtenemos Fecha de Alta
                    intPosicion = InStr(strLineaImportada, ",")
                    If Trim(Left(strLineaImportada, intPosicion - 1)) <> "" Then
                        datFechaAlta = CDate(Left(strLineaImportada, intPosicion - 1)) 'Verificar que Funcione Cdate
                    Else
                        datFechaAlta = Date
                    End If
                    strLineaImportada = Right(strLineaImportada, Len(strLineaImportada) - intPosicion)
                    'Obtenemos Nro DNI Familiar
                    intPosicion = InStr(strLineaImportada, ",,")
                    strDNIFamiliar = Left(strLineaImportada, intPosicion - 1)
                    strLineaImportada = Right(strLineaImportada, Len(strLineaImportada) - intPosicion - 1)
                    'Pamos de largo Género
                    intPosicion = InStr(strLineaImportada, ",")
                    strLineaImportada = Right(strLineaImportada, Len(strLineaImportada) - intPosicion)
                    'Obtenemos Si es o No Discapacitado
                    intPosicion = InStr(strLineaImportada, ",")
                    strDiscapacidad = Left(strLineaImportada, intPosicion - 1)
                    strLineaImportada = Right(strLineaImportada, Len(strLineaImportada) - intPosicion)
                    'Obtenemos Nivel de Estudio
                    intPosicion = InStr(strLineaImportada, ",")
                    strNivelDeEstudio = Left(strLineaImportada, intPosicion - 1)
                    strLineaImportada = Right(strLineaImportada, Len(strLineaImportada) - intPosicion)
                    'Pamos de largo Vive
                    intPosicion = InStr(strLineaImportada, ",")
                    strVive = Left(strLineaImportada, intPosicion - 1)
                    If strVive = "NO" Then
                        Call InfoGeneral(strNombreFamiliar & " NO vive", ImportacionPadronFamiliares)
                    End If
                    strLineaImportada = Right(strLineaImportada, Len(strLineaImportada) - intPosicion)
                    'Obtenemos si Cobra Salario
                    intPosicion = InStr(strLineaImportada, ",")
                    strCobraSalario = Left(strLineaImportada, intPosicion - 1)
                    strLineaImportada = Right(strLineaImportada, Len(strLineaImportada) - intPosicion)
                    'Obtenemos si Cobra Obra Social
                    intPosicion = InStr(strLineaImportada, ",")
                    strObraSocial = Left(strLineaImportada, intPosicion - 1)
                    'Verificamos si el familiar ya está cargado
                    SQL = "Select * from CARGASDEFAMILIA Where PUESTOLABORAL= '" & strPuestoLaboral & "' And DNI= '" & strDNIFamiliar & "'"
                    If SQLNoMatch(SQL) = True Then
                        rstRegistroSlave.Open "CARGASDEFAMILIA", dbSlave, adOpenForwardOnly, adLockOptimistic
                        rstRegistroSlave.AddNew
                        rstRegistroSlave!DeducibleGanancias = False
                    Else
                        rstRegistroSlave.Open SQL, dbSlave, adOpenForwardOnly, adLockOptimistic
                    End If
                    rstRegistroSlave!PuestoLaboral = strPuestoLaboral
                    rstRegistroSlave!DNI = strDNIFamiliar
                    rstRegistroSlave!NombreCompleto = strNombreFamiliar
                    rstRegistroSlave!FechaAlta = datFechaAlta
                    rstRegistroSlave!CodigoParentesco = strParentesco
                    rstRegistroSlave!NivelDeEstudio = strNivelDeEstudio
                    If strDiscapacidad = "SI" Then
                        rstRegistroSlave!Discapacitado = True
                    ElseIf strDiscapacidad = "NO" Then
                        rstRegistroSlave!Discapacitado = False
                    End If
                    If strCobraSalario = "SI" Then
                        rstRegistroSlave!CobraSalario = True
                    ElseIf strCobraSalario = "NO" Then
                        rstRegistroSlave!CobraSalario = False
                    End If
                    If strObraSocial = "1" Then
                        rstRegistroSlave!AdherenteObraSocial = True
                    ElseIf strObraSocial = "0" Then
                        rstRegistroSlave!AdherenteObraSocial = False
                    End If
                    rstRegistroSlave.Update
                    rstRegistroSlave.Close
                End If
                i = i + 1
            Wend
        End With
        Set rstRegistroSlave = Nothing
        objExcel.ActiveWorkbook.Close False
        objExcel.Quit
        Set objExcel = Nothing
        Set xLibro = Nothing
    End If
Exit Sub

HayError:
    MsgBox Err.Number & " - " & Err.Description, , "Error!!!"
    objExcel.ActiveWorkbook.Close False
    objExcel.Quit
    Set objExcel = Nothing
    Set xLibro = Nothing
    rstRegistroSlave.Close
    Set rstRegistroSlave = Nothing

End Sub

Public Function InfoGeneral(StrAgregar As String, FormularioDestino As Form)
    
    FormularioDestino.txtInforme.Text = FormularioDestino.txtInforme.Text & vbCrLf & StrAgregar
    
End Function

Public Sub ImportarPadronFamiliares()

    Dim strDireccion As String
    Dim intNumeroArchivo As Integer
    Dim Result As Collection
    Dim strMiArray() As String
    Dim strVarios As String
    Dim i As Integer
    Dim j As Integer
    Dim strPuestoLaboral As String
    Dim strParentesco As String
    Dim strNombreFamiliar As String
    Dim datFechaAlta As Date
    Dim strDNIFamiliar As String
    Dim strDiscapacidad As String
    Dim strNivelDeEstudio As String
    Dim strVive As String
    Dim strCobraSalario As String
    Dim strObraSocial As String
    Dim SQL As String
    
    With ImportacionPadronFamiliares
        .txtInforme.Text = ""
        Call InfoGeneral("- INICIANDO IMPORTACIÓN PADRÓN FAMILIARES -", ImportacionPadronFamiliares)
        Call InfoGeneral("", ImportacionPadronFamiliares)
        .dlgMultifuncion.Filter = "Todos los Excel(*.csv)|*.csv|"
        .dlgMultifuncion.FileName = ""
        .dlgMultifuncion.ShowOpen
        strDireccion = .dlgMultifuncion.FileName
    End With
    If strDireccion = "" Then 'En caso de apretar cancelar
        Call InfoGeneral("IMPORTACION CANCELADA", ImportacionPadronFamiliares)
        Exit Sub
    Else
        intNumeroArchivo = FreeFile
        Open strDireccion For Input As #intNumeroArchivo
        strDireccion = ""
        Set Result = New Collection
        Do Until EOF(intNumeroArchivo)
            Line Input #intNumeroArchivo, strVarios
            Debug.Print strVarios
            Result.Add strVarios
        Loop
        Close #intNumeroArchivo
        Set rstRegistroSlave = New ADODB.Recordset
        'If Left(Result(0), 4) = "Date" Then
        For i = 2 To Result.Count
            strMiArray = Split(Result(i), ",")
            If UBound(strMiArray) = 18 Then
                If strMiArray(0) = "Documento:" Then
                    strPuestoLaboral = strMiArray(7) 'Obtenemos el Puesto Laboral
                    SQL = "Select * from AGENTES Where PUESTOLABORAL= '" & strPuestoLaboral & "'"
                    If SQLNoMatch(SQL) = True Then
                        strVarios = strMiArray(11) 'Procedemos a Buscar el Nro de Puesto a partir del nombre del agente
                        'Verificamos que exista el agente en la base principal
                        SQL = "Select * from AGENTES Where NOMBRECOMPLETO= '" & strVarios & "'"
                        If SQLNoMatch(SQL) = False Then
                            rstRegistroSlave.Open SQL, dbSlave, adOpenForwardOnly, adLockOptimistic
                            strPuestoLaboral = rstRegistroSlave!PuestoLaboral
                            rstRegistroSlave.Close
                        Else 'Pasamos al siguiente agente
                            Call InfoGeneral("EL Agente " & strVarios & " no existe en la base de datos principal, por favor verificar", ImportacionPadronFamiliares)
                            i = i + 1
                            strMiArray = Split(Result(i), ",")
                            While strMiArray(0) <> "Documento:" And i < Result.Count
                                i = i + 1
                                strMiArray = Split(Result(i), ",")
                            Wend
                            If Not i = Result.Count Then
                                i = i - 1
                            End If
                        End If
                    End If
                Else
                    'Obtenemos Descripcion de Parentesco
                    strParentesco = strMiArray(1)
                    SQL = "Select * from ASIGNACIONESFAMILIARES Where PARENTESCO = " & "'" & strParentesco & "'"
                    rstRegistroSlave.Open SQL, dbSlave, adOpenForwardOnly, adLockOptimistic
                    strParentesco = rstRegistroSlave!Codigo
                    rstRegistroSlave.Close
                    'Obtenemos Nombre del Familiar
                    strNombreFamiliar = strMiArray(4)
                    'Ajustamos el Nombre del Familiar
                    strVarios = strNombreFamiliar
                    strNombreFamiliar = Trim(strNombreFamiliar)
                    If strNombreFamiliar <> strVarios Then
                        Call InfoGeneral("Se procedió a eliminar los espacios en blanco de los extremos del Agente " & strNombreFamiliar, ImportacionPadronFamiliares)
                        strNombreFamiliar = Left(strNombreFamiliar, Len(strNombreFamiliar) - 1)
                    End If
                    strVarios = Replace(strNombreFamiliar, " ", "")
                    If Len(strVarios) Mod 2 = 0 Then
                        j = Len(strVarios) / 2
                        If Left(strVarios, j) = Right(strVarios, j) Then
                            strVarios = strNombreFamiliar
                            j = Len(strNombreFamiliar)
                            If j Mod 2 = 0 Then
                                strNombreFamiliar = Trim(Left(strNombreFamiliar, j / 2))
                            Else
                                strNombreFamiliar = Trim(Left(strNombreFamiliar, (j - 1) / 2))
                            End If
                            Call InfoGeneral("Se modificó el nombre del Agente " & strVarios & " por el nombre " & strNombreFamiliar, ImportacionPadronFamiliares)
                        End If
                    End If
                    'Obtenemos Fecha de Alta
                    strVarios = strMiArray(9)
                    If Trim(strVarios) <> "" Then
                        datFechaAlta = DateTime.DateSerial(Right(strVarios, 4), Mid(strVarios, 4, 2), Left(strVarios, 2))
                    Else
                        datFechaAlta = Date
                    End If
                    'Obtenemos Nro DNI Familiar
                    strDNIFamiliar = strMiArray(10)
                    'Obtenemos Si es o No Discapacitado
                    strDiscapacidad = strMiArray(13)
                    'Obtenemos Nivel de Estudio
                    strNivelDeEstudio = strMiArray(14)
                    'Pamos de largo Vive
                    strVive = strMiArray(15)
                    If strVive = "NO" Then
                        Call InfoGeneral(strNombreFamiliar & " NO vive", ImportacionPadronFamiliares)
                    End If
                    'Obtenemos si Cobra Salario
                    strCobraSalario = strMiArray(16)
                    'Obtenemos si Cobra Obra Social
                    strObraSocial = strMiArray(17)
                    'Verificamos si el familiar ya está cargado
                    SQL = "Select * from CARGASDEFAMILIA Where PUESTOLABORAL= '" & strPuestoLaboral & "' And DNI= '" & strDNIFamiliar & "'"
                    If SQLNoMatch(SQL) = True Then
                        rstRegistroSlave.Open "CARGASDEFAMILIA", dbSlave, adOpenForwardOnly, adLockOptimistic
                        rstRegistroSlave.AddNew
                        rstRegistroSlave!DeducibleGanancias = False
                    Else
                        rstRegistroSlave.Open SQL, dbSlave, adOpenForwardOnly, adLockOptimistic
                    End If
                    rstRegistroSlave!PuestoLaboral = strPuestoLaboral
                    rstRegistroSlave!DNI = strDNIFamiliar
                    rstRegistroSlave!NombreCompleto = strNombreFamiliar
                    rstRegistroSlave!FechaAlta = datFechaAlta
                    rstRegistroSlave!CodigoParentesco = strParentesco
                    rstRegistroSlave!NivelDeEstudio = strNivelDeEstudio
                    If strDiscapacidad = "SI" Then
                        rstRegistroSlave!Discapacitado = True
                    ElseIf strDiscapacidad = "NO" Then
                        rstRegistroSlave!Discapacitado = False
                    End If
                    If strCobraSalario = "SI" Then
                        rstRegistroSlave!CobraSalario = True
                    ElseIf strCobraSalario = "NO" Then
                        rstRegistroSlave!CobraSalario = False
                    End If
                    If strObraSocial = "1" Then
                        rstRegistroSlave!AdherenteObraSocial = True
                    ElseIf strObraSocial = "0" Then
                        rstRegistroSlave!AdherenteObraSocial = False
                    End If
                    rstRegistroSlave.Update
                    rstRegistroSlave.Close
                End If
            Else 'Intentamos quitar las comas que están de más
                If Left(strMiArray(11), 1) = Chr(34) Then
                    strMiArray(11) = Right(strMiArray(11), Len(strMiArray(11)) - 2) 'Le quitamos las comillas
                    j = 12
                    While strMiArray(j) <> ""
                        If Left(strMiArray(j), 1) = " " Then
                            strMiArray(11) = strMiArray(11) & strMiArray(j)
                        Else
                            strMiArray(11) = strMiArray(11) & " " & strMiArray(j)
                        End If
                        j = j + 1
                    Wend
                    strMiArray(11) = Left(strMiArray(11), Len(strMiArray(11)) - 2) 'Le quitamos las comillas
                    Result.Remove (i)
                    strVarios = "Documento:,,,,,,," & strMiArray(7) & ",,,," & strMiArray(11) & ",,,,,,,"
                    Result.Add strVarios, , i
                    i = i - 1
                ElseIf Left(strMiArray(4), 1) = Chr(34) Then
                    strMiArray(4) = Right(strMiArray(4), Len(strMiArray(4)) - 2) 'Le quitamos las comillas
                    j = 5
                    While strMiArray(j) <> ""
                        If Left(strMiArray(j), 1) = " " Then
                            strMiArray(4) = strMiArray(4) & strMiArray(j)
                        Else
                            strMiArray(4) = strMiArray(4) & " " & strMiArray(j)
                        End If
                        j = j + 1
                    Wend
                    strMiArray(4) = Left(strMiArray(4), Len(strMiArray(4)) - 2) 'Le quitamos las comillas
                    Result.Remove (i)
                    strVarios = strMiArray(0) & "," & strMiArray(1) & ",,," & strMiArray(4) & ",,,,," & strMiArray(j + 4) & "," & strMiArray(j + 5) & ",,," & strMiArray(j + 8) & "," & strMiArray(j + 9) & "," & strMiArray(j + 10) & "," & strMiArray(j + 11) & "," & strMiArray(j + 12) & "," & strMiArray(j + 13)
                    Result.Add strVarios, , i
                    i = i - 1
                Else
                    Call InfoGeneral("La fila Nro. " & i & " tiene comas de más, por favor verificar", ImportacionPadronFamiliares)
                    If strMiArray(0) = "Documento:" Then 'Pasamos al Siguiente Agente
                        i = i + 1
                        strMiArray = Split(Result(i), ",")
                        While strMiArray(0) <> "Documento:" And i < Result.Count
                            i = i + 1
                            strMiArray = Split(Result(i), ",")
                        Wend
                        If Not i = Result.Count Then
                            i = i - 1
                        End If
                    End If
                End If
            End If
        Next i
        'Else
            'Call InfoDownload(" ===> " & rstDownloadStockFollower![Simbolo] & "- No Existen Datos en la Web para el Ticker de Yahoo cargado, favor de verificar")
        'End If
        Set rstRegistroSlave = Nothing
        Set Result = Nothing
    End If

End Sub

Public Sub ImportarLiquidacionSueldos(CodigoLiquidacion As String)

    Dim strDireccion As String
    Dim intNumeroArchivo As Integer
    Dim strVarios As String
    Dim dblPuestoLaboral As Double
    Dim strCodigoConcepto As String
    Dim strDNI As String
    Dim strDenominacion As String
    Dim dblMonto As Double
    Dim SQL As String
    Dim i As Integer
    
    SQL = "Select * From LIQUIDACIONSUELDOS Where CODIGOLIQUIDACION = " & "'" & CodigoLiquidacion & "'"
    If SQLNoMatch(SQL) = False Then
        MsgBox "Proceda a borrar los registros de la liquidación destino antes de incorporar datos a la misma", vbCritical + vbOKOnly, "LIQUIDACIÓN A IMPORTAR LLENA"
        ImportacionLiquidacionSueldo.cmbCodigoLiquidacion.SetFocus
        Exit Sub
    End If
    
    With ImportacionLiquidacionSueldo
        .txtInforme.Text = ""
        Call InfoGeneral("- INICIANDO IMPORTACIÓN SUELDOS -", ImportacionLiquidacionSueldo)
        Call InfoGeneral("", ImportacionLiquidacionSueldo)
        '.dlgMultifuncion.Filter = "Todos los Excel(*.csv)|*.csv|"
        .dlgMultifuncion.FileName = ""
        .dlgMultifuncion.ShowOpen
        strDireccion = .dlgMultifuncion.FileName
    End With
    If strDireccion = "" Then 'En caso de apretar cancelar
        Call InfoGeneral("IMPORTACION CANCELADA", ImportacionLiquidacionSueldo)
        Exit Sub
    Else
        intNumeroArchivo = FreeFile
        Open strDireccion For Input As #intNumeroArchivo
        strDireccion = ""
        Set rstRegistroSlave = New ADODB.Recordset
        Line Input #intNumeroArchivo, strVarios
        'Pasamos de largo fecha (8 lugares) y tipo (3 lugares)
        strVarios = Right(strVarios, Len(strVarios) - 11)
        'Obtenemos DNI (8 lugares)
        strDNI = Left(strVarios, 8)
        'Pasamos de largo DNI (8 lugares)
        strVarios = Right(strVarios, Len(strVarios) - 8)
        Do Until EOF(intNumeroArchivo)
            'Obtenemos y formateamos el Puesto Laboral
            dblPuestoLaboral = Val(Left(strVarios, 6))
            'Verificar que exista el puesto
            SQL = "Select * from AGENTES Where PUESTOLABORAL= '" & CStr(dblPuestoLaboral) & "'"
            If SQLNoMatch(SQL) = True Then
                'Si no existe el puesto laboral, verificamos si existe el DNI
                SQL = "Select MID(CUIL,3,8) As DNI From AGENTES Where MID(CUIL,3,8) = '" & strDNI & "'"
                If SQLNoMatch(SQL) = True Then
                    'Si no existe el DNI, quiere decir que hay que agregar el agente
                    'Obtenemos Nombre del Agente
                    strVarios = Right(strVarios, Len(strVarios) - 12)
                    strVarios = Left(strVarios, 40)
                    Call InfoGeneral("El Agente " & RTrim(strVarios) & ", con DNI " & strDNI & " y P.L. " & CStr(dblPuestoLaboral) & ", no existe en la base de datos principal, por favor verificar", ImportacionLiquidacionSueldo)
                    'Pasamos de largo todos los registros asociados con el Puesto Laboral
                    Line Input #intNumeroArchivo, strVarios
                    'Pasamos de largo fecha, tipo y número de documento
                    strVarios = Right(strVarios, Len(strVarios) - 19)
                    While Val(Left(strVarios, 6)) = dblPuestoLaboral And EOF(intNumeroArchivo) = False
                        Line Input #intNumeroArchivo, strVarios
                        'Pasamos de largo fecha (8 lugares) y tipo (3 lugares)
                        strVarios = Right(strVarios, Len(strVarios) - 11)
                        'Obtenemos DNI (8 lugares)
                        strDNI = Left(strVarios, 8)
                        'Pasamos de largo DNI (8 lugares)
                        strVarios = Right(strVarios, Len(strVarios) - 8)
                    Wend
                Else
                    'Si existe el DNI, cambiamos el Puesto Laboral del Agente
                    rstRegistroSlave.Open SQL, dbSlave, adOpenForwardOnly, adLockOptimistic
                    rstRegistroSlave!PuestoLaboral = CStr(dblPuestoLaboral)
                    rstRegistroSlave.Update
                    rstRegistroSlave.Close
                    GoTo Principio_Incoporacion
                End If
            Else
Principio_Incoporacion:
                While Val(Left(strVarios, 6)) = dblPuestoLaboral And EOF(intNumeroArchivo) = False
                    'Pasamos de largo Puesto Laboral, Categoría, Clase, Permanente y Nombre del Agente
                    strVarios = Right(strVarios, Len(strVarios) - 52)
                    'Obtenemos y formateamos el Código Concepto
                    strCodigoConcepto = Format(Left(strVarios, 4), "0000")
                    'Verificar que exita el Código de Concepto
                    SQL = "Select * from CONCEPTOS Where CODIGO= '" & strCodigoConcepto & "'"
                    If SQLNoMatch(SQL) = True Then
                        'Obtenemos y formateamos denominación concepto para agregar
                        strDenominacion = Left(strVarios, 29)
                        strDenominacion = Right(strDenominacion, Len(strDenominacion) - 4)
                        strDenominacion = RTrim(strDenominacion)
                        rstRegistroSlave.Open "CONCEPTOS", dbSlave, adOpenForwardOnly, adLockOptimistic
                        rstRegistroSlave.AddNew
                        rstRegistroSlave!Codigo = strCodigoConcepto
                        rstRegistroSlave!Denominacion = strDenominacion
                        rstRegistroSlave.Update
                        rstRegistroSlave.Close
                        Call InfoGeneral("El Código de Concepto " & strCodigoConcepto & " fue agregado a la base de datos con la denominación: " & strDenominacion, IncorporarConceptoSueldo)
                    End If
                    'Pasamos de largo Código Concepto y Denominación Concepto
                    strVarios = Right(strVarios, Len(strVarios) - 29)
                    'Obtenemos y formateamos monto
                    strVarios = Left(strVarios, 12)
                    strVarios = Left(strVarios, 10) & "." & Right(strVarios, 2)
                    'Quitamos los ceros iniciales de la cadena
                    While Left(strVarios, 1) = 0
                        strVarios = Right(strVarios, Len(strVarios) - 1)
                    Wend
                    dblMonto = Val(strVarios)
                   'Verificamos si existe el concepto en la liquidación que se quiere incorporar
                    SQL = "Select * from LIQUIDACIONSUELDOS Where CODIGOLIQUIDACION= '" & CodigoLiquidacion & "' And PUESTOLABORAL= '" & CStr(dblPuestoLaboral) & "' And CODIGOCONCEPTO= '" & strCodigoConcepto & "'"
                    If SQLNoMatch(SQL) = True Then
                        rstRegistroSlave.Open "LIQUIDACIONSUELDOS", dbSlave, adOpenForwardOnly, adLockOptimistic
                        rstRegistroSlave.AddNew
                    Else
                        rstRegistroSlave.Open SQL, dbSlave, adOpenForwardOnly, adLockOptimistic
                    End If
                    rstRegistroSlave!CodigoLiquidacion = CodigoLiquidacion
                    rstRegistroSlave!PuestoLaboral = CStr(dblPuestoLaboral)
                    rstRegistroSlave!CodigoConcepto = strCodigoConcepto
                    rstRegistroSlave!Importe = dblMonto
                    rstRegistroSlave.Update
                    rstRegistroSlave.Close
                    Line Input #intNumeroArchivo, strVarios
                    'Pasamos de largo fecha (8 lugares) y tipo (3 lugares)
                    strVarios = Right(strVarios, Len(strVarios) - 11)
                    'Obtenemos DNI (8 lugares)
                    strDNI = Left(strVarios, 8)
                    'Pasamos de largo DNI (8 lugares)
                    strVarios = Right(strVarios, Len(strVarios) - 8)
                Wend
            End If
        Loop
        Call InfoGeneral("", ImportacionLiquidacionSueldo)
        Call InfoGeneral("- IMPORTACIÓN FINALIZADA -", ImportacionLiquidacionSueldo)
        Close #intNumeroArchivo
        Set rstRegistroSlave = Nothing
    End If
    
End Sub


Public Sub IncorporarConceptoPorArchivo(CodigoLiquidacion As String)

    Dim strDireccion As String
    Dim intNumeroArchivo As Integer
    Dim strVarios As String
    Dim dblPuestoLaboral As Double
    Dim strDNI As String
    Dim strCodigoConcepto As String
    Dim strDenominacion As String
    Dim dblMonto As Double
    Dim SQL As String
    
    With IncorporarConceptoSueldo
        .txtInforme.Text = ""
        Call InfoGeneral("- INICIANDO INCORPORACIÓN CONCEPTO -", IncorporarConceptoSueldo)
        Call InfoGeneral("", IncorporarConceptoSueldo)
        '.dlgMultifuncion.Filter = "Todos los Excel(*.csv)|*.csv|"
        .dlgMultifuncion.FileName = ""
        .dlgMultifuncion.ShowOpen
        strDireccion = .dlgMultifuncion.FileName
    End With
    If strDireccion = "" Then 'En caso de apretar cancelar
        Call InfoGeneral("INCORPORACIÓN CANCELADA", IncorporarConceptoSueldo)
        Exit Sub
    Else
        intNumeroArchivo = FreeFile
        Open strDireccion For Input As #intNumeroArchivo
        strDireccion = ""
        Set rstRegistroSlave = New ADODB.Recordset
        Line Input #intNumeroArchivo, strVarios
        'Pasamos de largo fecha (8 lugares) y tipo (3 lugares)
        strVarios = Right(strVarios, Len(strVarios) - 11)
        'Obtenemos DNI (8 lugares)
        strDNI = Left(strVarios, 8)
        'Pasamos de largo DNI (8 lugares)
        strVarios = Right(strVarios, Len(strVarios) - 8)
        Do Until EOF(intNumeroArchivo)
            'Obtenemos y formateamos el Puesto Laboral
            dblPuestoLaboral = Val(Left(strVarios, 6))
            'Verificar que exista el puesto
            SQL = "Select * from AGENTES Where PUESTOLABORAL= '" & CStr(dblPuestoLaboral) & "'"
            If SQLNoMatch(SQL) = True Then
                'Si no existe el puesto laboral, verificamos si existe el DNI
                SQL = "Select MID(CUIL, 3, 8) As DNI From AGENTES Where MID(CUIL,3,8) = '" & strDNI & "'"
                If SQLNoMatch(SQL) = True Then
                    'Si no existe el DNI, quiere decir que hay que agregar el agente
                    'Obtenemos Nombre del Agente
                    strVarios = Right(strVarios, Len(strVarios) - 12)
                    strVarios = Left(strVarios, 40)
                    Call InfoGeneral("El Agente " & RTrim(strVarios) & ", con DNI " & strDNI & " y P.L. " & CStr(dblPuestoLaboral) & ", no existe en la base de datos principal, por favor verificar", IncorporarConceptoSueldo)
                    'Pasamos de largo todos los registros asociados con el Puesto Laboral
                    Line Input #intNumeroArchivo, strVarios
                    'Pasamos de largo fecha, tipo y número de documento
                    strVarios = Right(strVarios, Len(strVarios) - 19)
                    While Val(Left(strVarios, 6)) = dblPuestoLaboral And EOF(intNumeroArchivo) = False
                        Line Input #intNumeroArchivo, strVarios
                        'Pasamos de largo fecha (8 lugares) y tipo (3 lugares)
                        strVarios = Right(strVarios, Len(strVarios) - 11)
                        'Obtenemos DNI (8 lugares)
                        strDNI = Left(strVarios, 8)
                        'Pasamos de largo DNI (8 lugares)
                        strVarios = Right(strVarios, Len(strVarios) - 8)
                    Wend
                Else
                    'Si existe el DNI, cambiamos el Puesto Laboral del Agente
                    rstRegistroSlave.Open SQL, dbSlave, adOpenForwardOnly, adLockOptimistic
                    rstRegistroSlave!PuestoLaboral = CStr(dblPuestoLaboral)
                    rstRegistroSlave.Update
                    rstRegistroSlave.Close
                    GoTo Principio_Incoporacion
                End If
            Else
Principio_Incoporacion:
                While Val(Left(strVarios, 6)) = dblPuestoLaboral And EOF(intNumeroArchivo) = False
                    'Pasamos de largo Puesto Laboral, Categoría, Clase, Permanente y Nombre del Agente
                    strVarios = Right(strVarios, Len(strVarios) - 52)
                    'Obtenemos y formateamos el Código Concepto
                    strCodigoConcepto = Format(Left(strVarios, 4), "0000")
                    'Verificar que exita el Código de Concepto
                    SQL = "Select * from CONCEPTOS Where CODIGO= '" & strCodigoConcepto & "'"
                    If SQLNoMatch(SQL) = True Then
                        'Obtenemos y formateamos denominación concepto para agregar
                        strDenominacion = Left(strVarios, 29)
                        strDenominacion = Right(strDenominacion, Len(strDenominacion) - 4)
                        strDenominacion = RTrim(strDenominacion)
                        rstRegistroSlave.Open "CONCEPTOS", dbSlave, adOpenForwardOnly, adLockOptimistic
                        rstRegistroSlave.AddNew
                        rstRegistroSlave!Codigo = strCodigoConcepto
                        rstRegistroSlave!Denominacion = strDenominacion
                        rstRegistroSlave.Update
                        rstRegistroSlave.Close
                        Call InfoGeneral("El Código de Concepto " & strCodigoConcepto & " fue agregado a la base de datos con la denominación: " & strDenominacion, IncorporarConceptoSueldo)
                    End If
                    'Pasamos de largo Código Concepto y Denominación Concepto
                    strVarios = Right(strVarios, Len(strVarios) - 29)
                    'Obtenemos y formateamos monto
                    strVarios = Left(strVarios, 12)
                    strVarios = Left(strVarios, 10) & "." & Right(strVarios, 2)
                    dblMonto = Val(strVarios)
                    'Verificamos si existe el concepto en la liquidación que se quiere incorporar
                    SQL = "Select * from LIQUIDACIONSUELDOS Where CODIGOLIQUIDACION= '" & CodigoLiquidacion & "' And PUESTOLABORAL= '" & CStr(dblPuestoLaboral) & "' And CODIGOCONCEPTO= '" & strCodigoConcepto & "'"
                    If SQLNoMatch(SQL) = True Then
                        'Si no existe, lo agregamos
                        rstRegistroSlave.Open "LIQUIDACIONSUELDOS", dbSlave, adOpenForwardOnly, adLockOptimistic
                        rstRegistroSlave.AddNew
                    Else
                        'Si existe, buscamos el importe y lo sumamos al nuevo
                        rstRegistroSlave.Open SQL, dbSlave, adOpenForwardOnly, adLockOptimistic
                        dblMonto = dblMonto + rstRegistroSlave!Importe
                    End If
                    rstRegistroSlave!CodigoLiquidacion = CodigoLiquidacion
                    rstRegistroSlave!PuestoLaboral = CStr(dblPuestoLaboral)
                    rstRegistroSlave!CodigoConcepto = strCodigoConcepto
                    rstRegistroSlave!Importe = dblMonto
                    rstRegistroSlave.Update
                    rstRegistroSlave.Close
                    Line Input #intNumeroArchivo, strVarios
                    'Pasamos de largo fecha (8 lugares) y tipo (3 lugares)
                    strVarios = Right(strVarios, Len(strVarios) - 11)
                    'Obtenemos DNI (8 lugares)
                    strDNI = Left(strVarios, 8)
                    'Pasamos de largo DNI (8 lugares)
                    strVarios = Right(strVarios, Len(strVarios) - 8)
                Wend
            End If
        Loop
        Call InfoGeneral("", IncorporarConceptoSueldo)
        Call InfoGeneral("- INCORPORACIÓN FINALIZADA -", IncorporarConceptoSueldo)
        Close #intNumeroArchivo
        Set rstRegistroSlave = Nothing
    End If
    
End Sub


Public Sub IncorporarConceptoPorArchivoSchema(CodigoLiquidacion As String)

    Dim strDireccion As String
    Dim strNombreArchivo As String
    Dim intNumeroArchivo As Integer
    Dim SQL As String
    Dim lngRows As Long
    Dim lngRows2 As Long
    
    With IncorporarConceptoSueldo
        .txtInforme.Text = ""
        Call InfoGeneral("- INICIANDO INCORPORACIÓN CONCEPTO -", IncorporarConceptoSueldo)
        Call InfoGeneral("", IncorporarConceptoSueldo)
        .dlgMultifuncion.Filter = "Todos los archivos de Texto(*.txt)|*.txt|"
        .dlgMultifuncion.FileName = ""
        .dlgMultifuncion.ShowOpen
        'Cargamos en el nombre del archivo y la ubicación del mismo
        strNombreArchivo = .dlgMultifuncion.FileTitle
        strDireccion = Replace(.dlgMultifuncion.FileName, "\" & strNombreArchivo, "")
    End With
    If strDireccion = "" Then 'En caso de apretar cancelar
        Call InfoGeneral("IMPORTACION CANCELADA", IncorporarConceptoSueldo)
        Exit Sub
'    ElseIf strNombreArchivo <> "ImportarSueldo.txt" Then
'        Call InfoGeneral("IMPORTACION CANCELADA debido a que el Nombre del Archivo no es Import.txt", IncorporarConceptoSueldo)
'        Exit Sub
    Else
        'Generamos el archivo Schema.ini que utilizaremos para importar el txt
        intNumeroArchivo = FreeFile
        'Open strDireccion For Input As #intNumeroArchivo
        Open strDireccion & "\schema.ini" For Output As #intNumeroArchivo
        Print #intNumeroArchivo, "[" & strNombreArchivo & "]"
        Print #intNumeroArchivo, "Format=FixedLength"
        Print #intNumeroArchivo, "TextDelimiter=none"
        Print #intNumeroArchivo, "ColNameHeader = False"
        Print #intNumeroArchivo, "MaxScanRows = 0"
        Print #intNumeroArchivo, "CharacterSet = ANSI"
        'Print #intNumeroArchivo, "DecimalSymbol = ,"
        Print #intNumeroArchivo, "Col1=Entidad Text Width 2"
        Print #intNumeroArchivo, "Col2=Mes Text Width 2"
        Print #intNumeroArchivo, "Col3=Año Text Width 4"
        Print #intNumeroArchivo, "Col4=TipoLiquidacion Text Width 1"
        Print #intNumeroArchivo, "Col5=NoSe Text Width 2"
        Print #intNumeroArchivo, "Col6=DNI Text Width 8"
        Print #intNumeroArchivo, "Col7=PuestoLaboral Text Width 6"
        Print #intNumeroArchivo, "Col8=Categoria Text Width 3"
        Print #intNumeroArchivo, "Col9=Clase Text Width 2"
        Print #intNumeroArchivo, "Col10=Planta Text Width 1"
        Print #intNumeroArchivo, "Col11=NombreCompleto Text Width 40"
        Print #intNumeroArchivo, "Col12=CodigoConcepto Text Width 4"
        Print #intNumeroArchivo, "Col13=DenominacionConcepto Text Width 25"
        Print #intNumeroArchivo, "Col14=Montotxt Text Width 12"
        Print #intNumeroArchivo, "Col15=Restante Text Width 44"
        Close #intNumeroArchivo
        
        'Creamos una tabla auxiliar (utilizando Schema.ini) para trabajar en ella
        SQL = "SELECT * INTO AUXILIARIMPORTACION " _
        & "FROM [" & strNombreArchivo & "] IN '" & strDireccion & "' 'TEXT;'"
        dbSlave.BeginTrans
        dbSlave.Execute SQL, lngRows, adCmdText Or adExecuteNoRecords
        dbSlave.CommitTrans
        Call InfoGeneral("1) Creamos tabla AUXILIARIMPORTACION con " & lngRows & " registros", IncorporarConceptoSueldo)
        Call InfoGeneral("", IncorporarConceptoSueldo)
        
        'Formateamos el Puesto Laboral de la tabla auxiliar
        Call InfoGeneral("2) Formateamos los campos de la tabla auxiliar", IncorporarConceptoSueldo)
        Call InfoGeneral("", IncorporarConceptoSueldo)
        'Formateamos el Puesto Laboral
        SQL = "Update AUXILIARIMPORTACION " _
        & "Set PUESTOLABORAL = CStr(Val(PUESTOLABORAL))"
        dbSlave.BeginTrans
        dbSlave.Execute SQL
        dbSlave.CommitTrans
        'Formateamos el Nombre Completo del Agente
        SQL = "Update AUXILIARIMPORTACION " _
        & "Set NOMBRECOMPLETO = RTrim(NOMBRECOMPLETO)"
        dbSlave.BeginTrans
        dbSlave.Execute SQL
        dbSlave.CommitTrans
        'Formateamos el Código Concepto
        SQL = "Update AUXILIARIMPORTACION " _
        & "Set CODIGOCONCEPTO = Format(CODIGOCONCEPTO, '0000')"
        dbSlave.BeginTrans
        dbSlave.Execute SQL
        dbSlave.CommitTrans
        'Formateamos la Denominación de los Conceptos de Sueldo
        SQL = "Update AUXILIARIMPORTACION " _
        & "Set DENOMINACIONCONCEPTO = RTrim(DENOMINACIONCONCEPTO)"
        dbSlave.BeginTrans
        dbSlave.Execute SQL
        dbSlave.CommitTrans
        'Formateamos el Monto (valores positivos)
        SQL = "Update AUXILIARIMPORTACION " _
        & "Set MONTOTXT = CStr(Val(Left(MONTOTXT,Len(MONTOTXT)-2))) & '.' & Right(MONTOTXT,2) " _
        & "Where InStr(1,MONTOTXT,'-') = 0"
        dbSlave.BeginTrans
        dbSlave.Execute SQL
        dbSlave.CommitTrans
        'Formateamos el Monto (valores negativos)
        SQL = "Update AUXILIARIMPORTACION " _
        & "Set MONTOTXT = '-' & Mid(MONTOTXT,InStr(1,MONTOTXT,'-') + 1, " _
        & "Len(MONTOTXT) - InStr(1,MONTOTXT,'-') -2) & '.' & Right(MONTOTXT,2) " _
        & "Where InStr(1,MONTOTXT,'-') <> 0"
        dbSlave.BeginTrans
        dbSlave.Execute SQL
        dbSlave.CommitTrans
        
        'Verificamos si todos los Puestos Laborales importados existen en la Base de Datos de Slave
        Call InfoGeneral("3) Verificamos si existen nuevos Puestos Laborales", IncorporarConceptoSueldo)
        Call InfoGeneral("", IncorporarConceptoSueldo)
        'Verificar que exista el puesto
        SQL = "Select AUXILIARIMPORTACION.PUESTOLABORAL from AUXILIARIMPORTACION " _
        & "Where AUXILIARIMPORTACION.PUESTOLABORAL Not In (Select AGENTES.PUESTOLABORAL From AGENTES)"
        If SQLNoMatch(SQL) = False Then
            'Si no existe el puesto laboral, verificamos si no existe el DNI
            SQL = "Select NOMBRECOMPLETO, PUESTOLABORAL, DNI From AUXILIARIMPORTACION " _
            & "Where DNI Not in (Select MID(CUIL,3,8) As AGENTESDNI From AGENTES) " _
            & "And AUXILIARIMPORTACION.PUESTOLABORAL Not In (Select AGENTES.PUESTOLABORAL From AGENTES) " _
            & "Group By NOMBRECOMPLETO, PUESTOLABORAL, DNI " _
            & "Order By NOMBRECOMPLETO Asc"
            If SQLNoMatch(SQL) = False Then
                'Si no existe el DNI ni el PuestoLaboral, quiere decir que hay que agregar el agente
                Call InfoGeneral(" - Es necesario incorporar los siguientes Agentes:", IncorporarConceptoSueldo)
                'Mostramos los agentes que son necesarios agregar
                Set rstRegistroSlave = New ADODB.Recordset
                rstRegistroSlave.Open SQL, dbSlave, adOpenForwardOnly, adLockOptimistic
                If rstRegistroSlave.BOF = False Then
                    rstRegistroSlave.MoveFirst
                    While rstRegistroSlave.EOF = False
                        Call InfoGeneral("  + " & rstRegistroSlave!NombreCompleto _
                        & ", con DNI " & rstRegistroSlave!DNI _
                        & " y P.L. " & rstRegistroSlave!PuestoLaboral, IncorporarConceptoSueldo)
                        rstRegistroSlave.MoveNext
                    Wend
                End If
                rstRegistroSlave.Close
                Set rstRegistroSlave = Nothing
                Call InfoGeneral("", IncorporarConceptoSueldo)
                'Eliminamos aquellos agentes cuyo DNI ni Puesto Laboral se encuentran en la base principal
                SQL = "Delete From AUXILIARIMPORTACION " _
                & "Where AUXILIARIMPORTACION.DNI Not in (Select MID(CUIL,3,8) As AGENTESDNI From AGENTES) " _
                & "And AUXILIARIMPORTACION.PUESTOLABORAL Not In (Select AGENTES.PUESTOLABORAL From AGENTES)"
                dbSlave.BeginTrans
                dbSlave.Execute SQL
                dbSlave.CommitTrans
            End If
            'Si no existe el Puesto Laboral, verificamos si existe el DNI
            SQL = "Select NOMBRECOMPLETO, PUESTOLABORAL, DNI From AUXILIARIMPORTACION " _
            & "Where DNI In (Select MID(CUIL,3,8) As AGENTESDNI From AGENTES) " _
            & "And AUXILIARIMPORTACION.PUESTOLABORAL Not In (Select AGENTES.PUESTOLABORAL From AGENTES) " _
            & "Group By NOMBRECOMPLETO, PUESTOLABORAL, DNI " _
            & "Order By NOMBRECOMPLETO Asc"
            If SQLNoMatch(SQL) = False Then
                'Si existe el DNI pero con otro PuestoLaboral, quiere decir que hay que modificar el Puesto Laboral
                Call InfoGeneral(" - Es necesario actualizar el Puesto Laboral de los siguientes Agentes:", IncorporarConceptoSueldo)
                'Mostramos los Puestos Laborales a modificar
                Set rstRegistroSlave = New ADODB.Recordset
                rstRegistroSlave.Open SQL, dbSlave, adOpenForwardOnly, adLockOptimistic
                If rstRegistroSlave.BOF = False Then
                    rstRegistroSlave.MoveFirst
                    While rstRegistroSlave.EOF = False
                        Call InfoGeneral("  + " & rstRegistroSlave!NombreCompleto _
                        & ", con DNI " & rstRegistroSlave!DNI _
                        & " y Nuevo P.L. " & rstRegistroSlave!PuestoLaboral, IncorporarConceptoSueldo)
                        rstRegistroSlave.MoveNext
                    Wend
                End If
                rstRegistroSlave.Close
                Set rstRegistroSlave = Nothing
                Call InfoGeneral("", IncorporarConceptoSueldo)
                'Eliminamos aquellos agentes cuyo Puesto Laboral no se encuentran en la base principal
                SQL = "Delete From AUXILIARIMPORTACION " _
                & "AUXILIARIMPORTACION.PUESTOLABORAL Not In (Select AGENTES.PUESTOLABORAL From AGENTES)"
                dbSlave.BeginTrans
                dbSlave.Execute SQL
                dbSlave.CommitTrans
            End If
        End If

        'Verificar que exita el Código de Concepto
        Call InfoGeneral("4) Verificamos si existen nuevos Conceptos de Sueldos", IncorporarConceptoSueldo)
        Call InfoGeneral("", IncorporarConceptoSueldo)
        SQL = "Select CODIGOCONCEPTO, DENOMINACIONCONCEPTO from AUXILIARIMPORTACION " _
        & "Where CODIGOCONCEPTO Not In (Select CODIGO From CONCEPTOS) " _
        & "Group by CODIGOCONCEPTO, DENOMINACIONCONCEPTO " _
        & "Order By CODIGOCONCEPTO Asc"
        If SQLNoMatch(SQL) = False Then
            Call InfoGeneral(" - Los siguientes Conceptos de Sueldo deben ser agregados:", IncorporarConceptoSueldo)
            Set rstRegistroSlave = New ADODB.Recordset
            rstRegistroSlave.Open SQL, dbSlave, adOpenForwardOnly, adLockOptimistic
            If rstRegistroSlave.BOF = False Then
                rstRegistroSlave.MoveFirst
                While rstRegistroSlave.EOF = False
                    Call InfoGeneral("  + " & rstRegistroSlave!DenominacionConcepto _
                    & ", con Codigo: " & rstRegistroSlave!CodigoConcepto, IncorporarConceptoSueldo)
                    rstRegistroSlave.MoveNext
                Wend
            End If
            rstRegistroSlave.Close
            Set rstRegistroSlave = Nothing
            Call InfoGeneral("", IncorporarConceptoSueldo)
        End If
        
        'Insertamos los registros de la Tabla Auxiliar Importación en la Tabla Liquidación Sueldo
        'Actualizamos los registros de aquellos conceptos que ya se hayan liquidados en la tabla sueldo
        SQL = "Update LIQUIDACIONSUELDOS Inner Join " _
        & "(Select PUESTOLABORAL, CODIGOCONCEPTO, Val(Montotxt) As Monto From AUXILIARIMPORTACION) As AUXILIARIMPORTACION " _
        & "On LIQUIDACIONSUELDOS.PUESTOLABORAL = AUXILIARIMPORTACION.PUESTOLABORAL " _
        & "And LIQUIDACIONSUELDOS.CODIGOCONCEPTO = AUXILIARIMPORTACION.CODIGOCONCEPTO " _
        & "Set LIQUIDACIONSUELDOS.Importe = AUXILIARIMPORTACION.Monto + LIQUIDACIONSUELDOS.IMPORTE " _
        & "Where LIQUIDACIONSUELDOS.CodigoLiquidacion = '" & CodigoLiquidacion & "'"
        dbSlave.BeginTrans
        dbSlave.Execute SQL, lngRows
        dbSlave.CommitTrans
        'Actualizamos el Monto de la tabla auxiliar asignando valor 0 de todos los registros ya importados en paso anterior
        SQL = "Update AUXILIARIMPORTACION Inner Join LIQUIDACIONSUELDOS " _
        & "On AUXILIARIMPORTACION.PUESTOLABORAL = LIQUIDACIONSUELDOS.PUESTOLABORAL " _
        & "And AUXILIARIMPORTACION.CODIGOCONCEPTO = LIQUIDACIONSUELDOS.CODIGOCONCEPTO " _
        & "Set AUXILIARIMPORTACION.Montotxt = '0' " _
        & "Where LIQUIDACIONSUELDOS.CodigoLiquidacion = '" & CodigoLiquidacion & "'"
        dbSlave.BeginTrans
        dbSlave.Execute SQL
        dbSlave.CommitTrans
        'Insertamos los restantes registros de la Tabla Auxiliar Importación en la Tabla Liquidación Sueldo
        SQL = "Insert Into LIQUIDACIONSUELDOS (CodigoLiquidacion, PuestoLaboral, CodigoConcepto, Importe)" & _
        " Select '" & CodigoLiquidacion & "', PuestoLaboral, CodigoConcepto, Val(Montotxt)" & _
        " From AUXILIARIMPORTACION Where Val(AUXILIARIMPORTACION.Montotxt) <> 0"
        dbSlave.BeginTrans
        dbSlave.Execute SQL, lngRows2
        dbSlave.CommitTrans
        Call InfoGeneral("5) Insertamos " & lngRows + lngRows2 & " registros de los datos importados", IncorporarConceptoSueldo)
        Call InfoGeneral("", IncorporarConceptoSueldo)

        'Borramos la tabla creada
        Call InfoGeneral("6) Borramos tabla AUXILIARIMPORTACION", IncorporarConceptoSueldo)
        SQL = "Drop Table AUXILIARIMPORTACION"
        dbSlave.BeginTrans
        dbSlave.Execute SQL
        dbSlave.CommitTrans
'        'Borramos schema.ini
'        Kill strDireccion & "\schema.ini"
        
    End If
            
End Sub



Public Sub IncorporarConceptoMontoFijo(CodigoLiquidacion As String, _
CodigoConcepto As String, Monto As Double, EsRemunerativo As Boolean)

    Dim SQL As String
    
    'Verificamos si la liquidación no está vacía
    SQL = "Select * From LIQUIDACIONSUELDOS Where CODIGOLIQUIDACION = " & "'" & CodigoLiquidacion & "'"
    If SQLNoMatch(SQL) = True Then
        MsgBox "La liquidación que desea modificar debe contener registros", vbCritical + vbOKOnly, "LIQUIDACIÓN VACÍA"
        IncorporarConceptoSueldo.cmbCodigoLiquidacion.SetFocus
        Exit Sub
    End If
    
    'Inicio
    IncorporarConceptoSueldo.txtInforme.Text = ""
    Call InfoGeneral("- INICIANDO INCORPORACIÓN CONCEPTO -", IncorporarConceptoSueldo)
    Call InfoGeneral("", IncorporarConceptoSueldo)
    'Verificamos si el concepto a incorporar ya esta liquidado
    SQL = "Select * From LIQUIDACIONSUELDOS " _
    & "Where CODIGOLIQUIDACION = " & "'" & CodigoLiquidacion & "' " _
    & "And CODIGOCONCEPTO = " & "'" & CodigoConcepto & "'"
    If SQLNoMatch(SQL) = True Then
        'Si no está liquidado el concepto, lo liquidamos para todos aquellos que cobren básico (CUIDADO CON ESTO)
        SQL = "Insert Into LIQUIDACIONSUELDOS (CodigoLiquidacion, PuestoLaboral, CodigoConcepto, Importe)" & _
        " Select '" & CodigoLiquidacion & "', PuestoLaboral, '" & CodigoConcepto & "', Format((" & De_Num_a_Tx_01(Monto, , 2) & "), '#.00')" & _
        " From LIQUIDACIONSUELDOS Where CodigoLiquidacion = " & "'" & CodigoLiquidacion & "' And CODIGOCONCEPTO = '0001'"
        dbSlave.BeginTrans
        dbSlave.Execute SQL
        dbSlave.CommitTrans
    Else
        'Si el concepto ya está liquidado, lo modificamos
        SQL = "Update LIQUIDACIONSUELDOS" _
        & " Set Importe = Importe + Format((" & De_Num_a_Tx_01(Monto, , 2) & "), '#.00')" _
        & " Where CodigoLiquidacion = " & "'" & CodigoLiquidacion & "'" _
        & " And CODIGOCONCEPTO = " & "'" & CodigoConcepto & "'"
        dbSlave.BeginTrans
        dbSlave.Execute SQL
        dbSlave.CommitTrans
    End If
    'Actualizamos el Importe Haber Óptimo (0115) y Total Bruto (9998)
    'Haber Óptimo (0115)
    SQL = "Update LIQUIDACIONSUELDOS" _
    & " Set Importe = Importe + Format((" & De_Num_a_Tx_01(Monto, , 2) & "), '#.00')" _
    & " Where CodigoLiquidacion = " & "'" & CodigoLiquidacion & "'" _
    & " And CODIGOCONCEPTO = '0115'"
    dbSlave.BeginTrans
    dbSlave.Execute SQL
    dbSlave.CommitTrans
    'Total Bruto (9998)
    SQL = "Update LIQUIDACIONSUELDOS" _
    & " Set Importe = Importe + Format((" & De_Num_a_Tx_01(Monto, , 2) & "), '#.00')" _
    & " Where CodigoLiquidacion = " & "'" & CodigoLiquidacion & "'" _
    & " And CODIGOCONCEPTO = '9998'"
    dbSlave.BeginTrans
    dbSlave.Execute SQL
    dbSlave.CommitTrans
    'Actualizamos el Importe de Jubilación y O. Social de ser necesario
    If EsRemunerativo = True Then
        'Actualizamos el Importe de Jubilación Personal (0208)
        SQL = "Update LIQUIDACIONSUELDOS" _
        & " Set Importe = Importe + Format((" & De_Num_a_Tx_01(Monto * 0.185, , 2) & "), '#.00')" _
        & " Where CodigoLiquidacion = " & "'" & CodigoLiquidacion & "'" _
        & " And CODIGOCONCEPTO = '0208'"
        dbSlave.BeginTrans
        dbSlave.Execute SQL
        dbSlave.CommitTrans
        'Actualizamos el Importe de Jubilación Personal (0209)
        SQL = "Update LIQUIDACIONSUELDOS" _
        & " Set Importe = Importe + Format((" & De_Num_a_Tx_01(Monto * 0.185, , 2) & "), '#.00')" _
        & " Where CodigoLiquidacion = " & "'" & CodigoLiquidacion & "'" _
        & " And CODIGOCONCEPTO = '0209'"
        dbSlave.BeginTrans
        dbSlave.Execute SQL
        dbSlave.CommitTrans
        'Actualizamos el Importe de Obra Social Personal (0212)
        SQL = "Update LIQUIDACIONSUELDOS" _
        & " Set Importe = Importe + Format((" & De_Num_a_Tx_01(Monto * 0.05, , 2) & "), '#.00')" _
        & " Where CodigoLiquidacion = " & "'" & CodigoLiquidacion & "'" _
        & " And CODIGOCONCEPTO = '0212'"
        dbSlave.BeginTrans
        dbSlave.Execute SQL
        dbSlave.CommitTrans
        'Actualizamos el Importe de Obra Social Estatal (0213)
        SQL = "Update LIQUIDACIONSUELDOS" _
        & " Set Importe = Importe + Format((" & De_Num_a_Tx_01(Monto * 0.04, , 2) & "), '#.00')" _
        & " Where CodigoLiquidacion = " & "'" & CodigoLiquidacion & "'" _
        & " And CODIGOCONCEPTO = '0213'"
        dbSlave.BeginTrans
        dbSlave.Execute SQL
        dbSlave.CommitTrans
        'Actualizamos el Importe de ART (0381)
        SQL = "Update LIQUIDACIONSUELDOS" _
        & " Set Importe = Importe + Format((" & De_Num_a_Tx_01(Monto * 0.0173, , 2) & "), '#.00')" _
        & " Where CodigoLiquidacion = " & "'" & CodigoLiquidacion & "'" _
        & " And CODIGOCONCEPTO = '0381'"
        dbSlave.BeginTrans
        dbSlave.Execute SQL
        dbSlave.CommitTrans
    End If
 
'Fin
    Call InfoGeneral("", IncorporarConceptoSueldo)
    Call InfoGeneral("- INCORPORACIÓN FINALIZADA -", IncorporarConceptoSueldo)
    
End Sub

Public Sub ImportarLiquidacionHonorarios()

    Dim strDireccion        As String
    Dim intNumeroArchivo    As Integer
    Dim in_sTableName       As String
    Dim SQL_InsertPrefix   As String
    Dim SQL                As String
    Dim intIndex            As Integer
    Dim strComprobante      As String
    Dim sLine               As String
    Dim vasFields           As Variant
    Dim datFecha            As Date
    Dim dblImporte          As Double
    
    in_sTableName = "LiquidacionHonorarios"
    
    With ImportacionLiquidacionHonorarios
        .txtInforme.Text = ""
        Call InfoGeneral("- INICIANDO IMPORTACIÓN HONORARIOS -", ImportacionLiquidacionHonorarios)
        Call InfoGeneral("", ImportacionLiquidacionHonorarios)
        .dlgMultifuncion.Filter = "Todos los Excel(*.csv)|*.csv|"
        .dlgMultifuncion.FileName = ""
        .dlgMultifuncion.ShowOpen
        strDireccion = .dlgMultifuncion.FileName
    End With
    If strDireccion = "" Then 'En caso de apretar cancelar
        Call InfoGeneral("IMPORTACION CANCELADA", ImportacionLiquidacionHonorarios)
        Exit Sub
    Else
        intNumeroArchivo = FreeFile
        Open strDireccion For Input As #intNumeroArchivo
        strDireccion = ""
        'Verificamos si el archivo esta vacio
        If EOF(intNumeroArchivo) Then
            Close #intNumeroArchivo
            Exit Sub
        End If
        'Buscamos un numero libre de comprobante de honorarios
        For intIndex = 1 To 99
            strComprobante = "NoSIIF" & Format(intIndex, "00")
            SQL = "Select * From LIQUIDACIONHONORARIOS" _
            & " Where COMPROBANTE = '" & strComprobante & "'"
            If SQLNoMatch(SQL) = True Then
                Exit For
            End If
        Next intIndex
        ' Read field names from the top line.
        'Line Input #iFileNo, sLine
    
        ' Note that in a "proper" CSV file, there should not be any trailing spaces, and all strings should be surrounded by double quotes.
        ' However, the top row provided can simply be used "as is" in the SQL string.
        SQL_InsertPrefix = "INSERT INTO " & in_sTableName & "" _
        & " (Comprobante, Tipo, Actividad, Partida, Fecha, Proveedor, MontoBruto, Sellos, LibramientoPago, IIBB, OtraRetencion, Anticipo, Seguro, Descuento)" _
        & " VALUES('" & strComprobante & "', 'O', 0, 0,"
        
        Do Until EOF(intNumeroArchivo)
    
            Line Input #intNumeroArchivo, sLine
    
            ' Initialise SQL string.
            SQL = SQL_InsertPrefix
    
            ' Again, in a proper CSV file, the spaces around the commas shouldn't be there, and the fields should be double quoted.
            ' Since the data is supposed to look like this, then the assumption is that the delimiter is space-comma-space.
            vasFields = Split(sLine, Chr(34) & "," & Chr(34))
    
            ' Build up each value, separated by comma.
            ' It is assumed that all values here are string, so they will be double quoted.
            ' However, if there are non-string values, you will have to ensure they don't get quoted.
            
            'Capturamos la fecha hasta del período y la formateamos
            datFecha = DateTime.DateSerial(Right(vasFields(5), 4), Mid(vasFields(5), 4, 2), Left(vasFields(5), 2))
            SQL = SQL & "#" & Format(datFecha, "MM/DD/YYYY") & "#,"
            'Capturamos el nombre del facturero
            SQL = SQL & "'" & vasFields(28) & "',"
            'Capturamos el importe de Importe Bruto y la formateamos
            dblImporte = De_Txt_a_Num_01(vasFields(31), 2, ".")
            SQL = SQL & De_Num_a_Tx_01(dblImporte) & ","
            'Capturamos el importe de Sellos y la formateamos
            dblImporte = De_Txt_a_Num_01(vasFields(32), 2, ".")
            SQL = SQL & De_Num_a_Tx_01(dblImporte) & ","
            'Capturamos el importe de libramiento de pago y la formateamos
            dblImporte = De_Txt_a_Num_01(vasFields(33), 2, ".")
            SQL = SQL & De_Num_a_Tx_01(dblImporte) & ","
            'Capturamos el importe de IIBB y la formateamos
            dblImporte = De_Txt_a_Num_01(vasFields(34), 2, ".")
            SQL = SQL & De_Num_a_Tx_01(dblImporte) & ","
            'Capturamos el importe de Otras Retenciones y la formateamos
            dblImporte = De_Txt_a_Num_01(vasFields(36), 2, ".")
            SQL = SQL & De_Num_a_Tx_01(dblImporte) & ","
            'Capturamos el importe de Anticipo y la formateamos
            dblImporte = De_Txt_a_Num_01(vasFields(37), 2, ".")
            SQL = SQL & De_Num_a_Tx_01(dblImporte) & ","
            'Capturamos el importe de Seguro y la formateamos
            dblImporte = De_Txt_a_Num_01(vasFields(38), 2, ".")
            SQL = SQL & De_Num_a_Tx_01(dblImporte) & ","
            'Capturamos el importe de Descuento
            dblImporte = De_Txt_a_Num_01(vasFields(39), 2, ".")
            SQL = SQL & De_Num_a_Tx_01(dblImporte) & ")"
            
'            For lIndex = 0 To UBound(vasFields) - 1
'                sSQL = sSQL & "'"
'                sSQL = sSQL & vasFields(lIndex)
'                sSQL = sSQL & "'"
'                sSQL = sSQL & ","
'            Next lIndex
'            ' This chunk of code is for the last item, which does not have a following comma.
'            sSQL = sSQL & "'"
'            sSQL = sSQL & vasFields(lIndex)
'            sSQL = sSQL & "'"
'            sSQL = sSQL & ")"
    
            'Run the SQL command.
            'Debug.Print SQL
            dbSlave.BeginTrans
            dbSlave.Execute SQL
            dbSlave.CommitTrans
    
        Loop
        'Verificamos si existen agentes sin Estructura
        SQL = "Select PROVEEDOR From LIQUIDACIONHONORARIOS" _
        & " Where COMPROBANTE = '" & strComprobante & "'" _
        & " And PROVEEDOR Not In (Select AGENTES As PROVEEDOR From PRECARIZADOS)"
        If SQLNoMatch(SQL) = False Then
            Call InfoGeneral("", ImportacionLiquidacionHonorarios)
            Call InfoGeneral(" - Los siguientes AGENTES no cuentan con ESTRUCTURA PRESUPUESTARIA:", ImportacionLiquidacionHonorarios)
            'Mostramos los agentes que no tienen estructura presupuestaria
            Set rstRegistroSlave = New ADODB.Recordset
            rstRegistroSlave.Open SQL, dbSlave, adOpenForwardOnly, adLockOptimistic
            If rstRegistroSlave.BOF = False Then
                rstRegistroSlave.MoveFirst
                While rstRegistroSlave.EOF = False
                    Call InfoGeneral("- " & rstRegistroSlave!Proveedor, ImportacionLiquidacionHonorarios)
                    rstRegistroSlave.MoveNext
                Wend
            End If
            rstRegistroSlave.Close
            Set rstRegistroSlave = Nothing
            'Insertamos los agentes que no tienen estructura en la tabla Precarizados
            SQL = "Insert Into PRECARIZADOS (Agentes, Actividad, Partida)" _
            & " Select Proveedor, '" & "00-00-00" & "', '" & "000" & "'" _
            & " From (" & SQL & ")"
            dbSlave.BeginTrans
            dbSlave.Execute SQL
            dbSlave.CommitTrans
        End If
        Call InfoGeneral("", ImportacionLiquidacionHonorarios)
        Call InfoGeneral("- IMPORTACION FINALIZADA -", ImportacionLiquidacionHonorarios)
        Close #intNumeroArchivo
        SQL = ""
        SQL_InsertPrefix = ""
        sLine = ""
        dblImporte = 0
        datFecha = 0
    End If
    
End Sub


Public Sub IncorporarMontoFijo(CodigoLiquidacion As String)

    Dim strDireccion As String
    Dim intNumeroArchivo As Integer
    Dim strVarios As String
    Dim dblPuestoLaboral As Double
    Dim strCodigoConcepto As String
    Dim dblMonto As Double
    Dim SQL As String
    
    With ImportacionLiquidacionSueldo
        Call InfoGeneral("- INICIANDO INCORPORACIÓN CONCEPTO -", IncorporarConceptoSueldo)
        Set rstRegistroSlave = New ADODB.Recordset
    
    End With
    SQL = "Select * from CONCEPTOS Where CODIGO= '" & strCodigoConcepto & "'"
    If SQLNoMatch(SQL) = False Then
        SQL = "Select * from LIQUIDACIONSUELDOS Where CODIGOLIQUIDACION= '" & CodigoLiquidacion & "' And PUESTOLABORAL= '" & CStr(dblPuestoLaboral) & "' And CODIGOCONCEPTO= '" & strCodigoConcepto & "'"
        If SQLNoMatch(SQL) = True Then
            rstRegistroSlave.Open "LIQUIDACIONSUELDOS", dbSlave, adOpenForwardOnly, adLockOptimistic
            rstRegistroSlave.AddNew
        Else
            rstRegistroSlave.Open SQL, dbSlave, adOpenForwardOnly, adLockOptimistic
        End If
        rstRegistroSlave!CodigoLiquidacion = CodigoLiquidacion
        rstRegistroSlave!PuestoLaboral = CStr(dblPuestoLaboral)
        rstRegistroSlave!CodigoConcepto = strCodigoConcepto
        rstRegistroSlave!Importe = dblMonto
        rstRegistroSlave.Update
        rstRegistroSlave.Close
    Else
        Call InfoGeneral("El Código de Concepto " & strCodigoConcepto & " no existe en la base de datos principal, por favor verificar", ImportacionLiquidacionSueldo)
    End If
    
    rstRegistroSlave.Close
    Set rstRegistroSlave = Nothing
    
End Sub

'This function reads a file into a string.                        '
'I found this in the book Programming Excel with VBA and .NET.    '
Public Function QuickRead(FName As String) As String
    Dim i As Integer
    Dim res As String
    Dim l As Long

    i = FreeFile
    l = FileLen(FName)
    res = Space(l)
    Open FName For Binary Access Read As #i
    Get #i, , res
    Close i
    QuickRead = res
End Function

Public Sub ImportarLiquidacionSueldosSchema(CodigoLiquidacion As String)

    Dim strDireccion As String
    Dim strNombreArchivo As String
    Dim intNumeroArchivo As Integer
    Dim SQL As String
    Dim lngRows As Long
    
    
    SQL = "Select * From LIQUIDACIONSUELDOS Where CODIGOLIQUIDACION = " & "'" & CodigoLiquidacion & "'"
    If SQLNoMatch(SQL) = False Then
        MsgBox "Proceda a borrar los registros de la liquidación destino antes de incorporar datos a la misma", vbCritical + vbOKOnly, "LIQUIDACIÓN A IMPORTAR LLENA"
        ImportacionLiquidacionSueldo.cmbCodigoLiquidacion.SetFocus
        Exit Sub
    End If
    
    With ImportacionLiquidacionSueldo
        .txtInforme.Text = ""
        Call InfoGeneral("- INICIANDO IMPORTACIÓN SUELDOS -", ImportacionLiquidacionSueldo)
        Call InfoGeneral("", ImportacionLiquidacionSueldo)
        .dlgMultifuncion.Filter = "Todos los archivos de Texto(*.txt)|*.txt|"
        .dlgMultifuncion.FileName = ""
        .dlgMultifuncion.ShowOpen
        'Cargamos en el nombre del archivo y la ubicación del mismo
        strNombreArchivo = .dlgMultifuncion.FileTitle
        strDireccion = Replace(.dlgMultifuncion.FileName, "\" & strNombreArchivo, "")
    End With
    If strDireccion = "" Then 'En caso de apretar cancelar
        Call InfoGeneral("IMPORTACION CANCELADA", ImportacionLiquidacionSueldo)
        Exit Sub
'    ElseIf strNombreArchivo <> "ImportarSueldo.txt" Then
'        Call InfoGeneral("IMPORTACION CANCELADA debido a que el Nombre del Archivo no es Import.txt", ImportacionLiquidacionSueldo)
'        Exit Sub
    Else
        'Generamos el archivo Schema.ini que utilizaremos para importar el txt
        intNumeroArchivo = FreeFile
        'Open strDireccion For Input As #intNumeroArchivo
        Open strDireccion & "\schema.ini" For Output As #intNumeroArchivo
        Print #intNumeroArchivo, "[" & strNombreArchivo & "]"
        Print #intNumeroArchivo, "Format=FixedLength"
        Print #intNumeroArchivo, "TextDelimiter=none"
        Print #intNumeroArchivo, "ColNameHeader = False"
        Print #intNumeroArchivo, "MaxScanRows = 0"
        Print #intNumeroArchivo, "CharacterSet = ANSI"
        'Print #intNumeroArchivo, "DecimalSymbol = ,"
        Print #intNumeroArchivo, "Col1=Entidad Text Width 2"
        Print #intNumeroArchivo, "Col2=Mes Text Width 2"
        Print #intNumeroArchivo, "Col3=Año Text Width 4"
        Print #intNumeroArchivo, "Col4=TipoLiquidacion Text Width 1"
        Print #intNumeroArchivo, "Col5=NoSe Text Width 2"
        Print #intNumeroArchivo, "Col6=DNI Text Width 8"
        Print #intNumeroArchivo, "Col7=PuestoLaboral Text Width 6"
        Print #intNumeroArchivo, "Col8=Categoria Text Width 3"
        Print #intNumeroArchivo, "Col9=Clase Text Width 2"
        Print #intNumeroArchivo, "Col10=Planta Text Width 1"
        Print #intNumeroArchivo, "Col11=NombreCompleto Text Width 40"
        Print #intNumeroArchivo, "Col12=CodigoConcepto Text Width 4"
        Print #intNumeroArchivo, "Col13=DenominacionConcepto Text Width 25"
        Print #intNumeroArchivo, "Col14=Montotxt Text Width 12"
        Print #intNumeroArchivo, "Col15=Restante Text Width 44"
        Close #intNumeroArchivo
        
        'Creamos una tabla auxiliar (utilizando Schema.ini) para trabajar en ella
        SQL = "SELECT * INTO AUXILIARIMPORTACION " _
        & "FROM [" & strNombreArchivo & "] IN '" & strDireccion & "' 'TEXT;'"
        dbSlave.BeginTrans
        dbSlave.Execute SQL, lngRows, adCmdText Or adExecuteNoRecords
        dbSlave.CommitTrans
        Call InfoGeneral("1) Creamos tabla AUXILIARIMPORTACION con " & lngRows & " registros", ImportacionLiquidacionSueldo)
        Call InfoGeneral("", ImportacionLiquidacionSueldo)
        
        'Formateamos el Puesto Laboral de la tabla auxiliar
        Call InfoGeneral("2) Formateamos los campos de la tabla auxiliar", ImportacionLiquidacionSueldo)
        Call InfoGeneral("", ImportacionLiquidacionSueldo)
        'Formateamos el Puesto Laboral
        SQL = "Update AUXILIARIMPORTACION " _
        & "Set PUESTOLABORAL = CStr(Val(PUESTOLABORAL))"
        dbSlave.BeginTrans
        dbSlave.Execute SQL
        dbSlave.CommitTrans
        'Formateamos el Nombre Completo del Agente
        SQL = "Update AUXILIARIMPORTACION " _
        & "Set NOMBRECOMPLETO = RTrim(NOMBRECOMPLETO)"
        dbSlave.BeginTrans
        dbSlave.Execute SQL
        dbSlave.CommitTrans
        'Formateamos el Código Concepto
        SQL = "Update AUXILIARIMPORTACION " _
        & "Set CODIGOCONCEPTO = Format(CODIGOCONCEPTO, '0000')"
        dbSlave.BeginTrans
        dbSlave.Execute SQL
        dbSlave.CommitTrans
        'Formateamos la Denominación de los Conceptos de Sueldo
        SQL = "Update AUXILIARIMPORTACION " _
        & "Set DENOMINACIONCONCEPTO = RTrim(DENOMINACIONCONCEPTO)"
        dbSlave.BeginTrans
        dbSlave.Execute SQL
        dbSlave.CommitTrans
        'Formateamos el Monto (valores positivos)
        SQL = "Update AUXILIARIMPORTACION " _
        & "Set MONTOTXT = CStr(Val(Left(MONTOTXT,Len(MONTOTXT)-2))) & '.' & Right(MONTOTXT,2) " _
        & "Where InStr(1,MONTOTXT,'-') = 0"
        dbSlave.BeginTrans
        dbSlave.Execute SQL
        dbSlave.CommitTrans
        'Formateamos el Monto (valores negativos)
        SQL = "Update AUXILIARIMPORTACION " _
        & "Set MONTOTXT = '-' & Mid(MONTOTXT,InStr(1,MONTOTXT,'-') + 1, " _
        & "Len(MONTOTXT) - InStr(1,MONTOTXT,'-') -2) & '.' & Right(MONTOTXT,2) " _
        & "Where InStr(1,MONTOTXT,'-') <> 0"
        dbSlave.BeginTrans
        dbSlave.Execute SQL
        dbSlave.CommitTrans
        
        'Verificamos si todos los Puestos Laborales importados existen en la Base de Datos de Slave
        Call InfoGeneral("3) Verificamos si existen nuevos Puestos Laborales", ImportacionLiquidacionSueldo)
        Call InfoGeneral("", ImportacionLiquidacionSueldo)
        'Verificar que exista el puesto
        SQL = "Select AUXILIARIMPORTACION.PUESTOLABORAL from AUXILIARIMPORTACION " _
        & "Where AUXILIARIMPORTACION.PUESTOLABORAL Not In (Select AGENTES.PUESTOLABORAL From AGENTES)"
        If SQLNoMatch(SQL) = False Then
            'Si no existe el puesto laboral, verificamos si no existe el DNI
            SQL = "Select NOMBRECOMPLETO, PUESTOLABORAL, DNI From AUXILIARIMPORTACION " _
            & "Where DNI Not in (Select MID(CUIL,3,8) As AGENTESDNI From AGENTES) " _
            & "And AUXILIARIMPORTACION.PUESTOLABORAL Not In (Select AGENTES.PUESTOLABORAL From AGENTES) " _
            & "Group By NOMBRECOMPLETO, PUESTOLABORAL, DNI " _
            & "Order By NOMBRECOMPLETO Asc"
            If SQLNoMatch(SQL) = False Then
                'Si no existe el DNI ni el PuestoLaboral, quiere decir que hay que agregar el agente
                Call InfoGeneral(" - Es necesario incorporar los siguientes Agentes:", ImportacionLiquidacionSueldo)
                'Mostramos los agentes que son necesarios agregar
                Set rstRegistroSlave = New ADODB.Recordset
                rstRegistroSlave.Open SQL, dbSlave, adOpenForwardOnly, adLockOptimistic
                If rstRegistroSlave.BOF = False Then
                    rstRegistroSlave.MoveFirst
                    While rstRegistroSlave.EOF = False
                        Call InfoGeneral("  + " & rstRegistroSlave!NombreCompleto _
                        & ", con DNI " & rstRegistroSlave!DNI _
                        & " y P.L. " & rstRegistroSlave!PuestoLaboral, ImportacionLiquidacionSueldo)
                        rstRegistroSlave.MoveNext
                    Wend
                End If
                rstRegistroSlave.Close
                Set rstRegistroSlave = Nothing
                Call InfoGeneral("", ImportacionLiquidacionSueldo)
                'Eliminamos aquellos agentes cuyo DNI ni Puesto Laboral se encuentran en la base principal
                SQL = "Delete From AUXILIARIMPORTACION " _
                & "Where AUXILIARIMPORTACION.DNI Not in (Select MID(CUIL,3,8) As AGENTESDNI From AGENTES) " _
                & "And AUXILIARIMPORTACION.PUESTOLABORAL Not In (Select AGENTES.PUESTOLABORAL From AGENTES)"
                dbSlave.BeginTrans
                dbSlave.Execute SQL
                dbSlave.CommitTrans
            End If
            'Si no existe el Puesto Laboral, verificamos si existe el DNI
            SQL = "Select NOMBRECOMPLETO, PUESTOLABORAL, DNI From AUXILIARIMPORTACION " _
            & "Where DNI In (Select MID(CUIL,3,8) As AGENTESDNI From AGENTES) " _
            & "And AUXILIARIMPORTACION.PUESTOLABORAL Not In (Select AGENTES.PUESTOLABORAL From AGENTES) " _
            & "Group By NOMBRECOMPLETO, PUESTOLABORAL, DNI " _
            & "Order By NOMBRECOMPLETO Asc"
            If SQLNoMatch(SQL) = False Then
                'Si existe el DNI pero con otro PuestoLaboral, quiere decir que hay que modificar el Puesto Laboral
                Call InfoGeneral(" - Es necesario actualizar el Puesto Laboral de los siguientes Agentes:", ImportacionLiquidacionSueldo)
                'Mostramos los Puestos Laborales a modificar
                Set rstRegistroSlave = New ADODB.Recordset
                rstRegistroSlave.Open SQL, dbSlave, adOpenForwardOnly, adLockOptimistic
                If rstRegistroSlave.BOF = False Then
                    rstRegistroSlave.MoveFirst
                    While rstRegistroSlave.EOF = False
                        Call InfoGeneral("  + " & rstRegistroSlave!NombreCompleto _
                        & ", con DNI " & rstRegistroSlave!DNI _
                        & " y Nuevo P.L. " & rstRegistroSlave!PuestoLaboral, ImportacionLiquidacionSueldo)
                        rstRegistroSlave.MoveNext
                    Wend
                End If
                rstRegistroSlave.Close
                Set rstRegistroSlave = Nothing
                Call InfoGeneral("", ImportacionLiquidacionSueldo)
                'Eliminamos aquellos agentes cuyo Puesto Laboral no se encuentran en la base principal
                SQL = "Delete From AUXILIARIMPORTACION " _
                & "Where AUXILIARIMPORTACION.PUESTOLABORAL Not In (Select AGENTES.PUESTOLABORAL From AGENTES)"
                dbSlave.BeginTrans
                dbSlave.Execute SQL
                dbSlave.CommitTrans
            End If
        End If

        'Verificar que exita el Código de Concepto
        Call InfoGeneral("4) Verificamos si existen nuevos Conceptos de Sueldos", ImportacionLiquidacionSueldo)
        Call InfoGeneral("", ImportacionLiquidacionSueldo)
        SQL = "Select CODIGOCONCEPTO, DENOMINACIONCONCEPTO from AUXILIARIMPORTACION " _
        & "Where CODIGOCONCEPTO Not In (Select CODIGO From CONCEPTOS) " _
        & "Group by CODIGOCONCEPTO, DENOMINACIONCONCEPTO " _
        & "Order By CODIGOCONCEPTO Asc"
        If SQLNoMatch(SQL) = False Then
            Call InfoGeneral(" - Los siguientes Conceptos de Sueldo deben ser agregados:", ImportacionLiquidacionSueldo)
            Set rstRegistroSlave = New ADODB.Recordset
            rstRegistroSlave.Open SQL, dbSlave, adOpenForwardOnly, adLockOptimistic
            If rstRegistroSlave.BOF = False Then
                rstRegistroSlave.MoveFirst
                While rstRegistroSlave.EOF = False
                    Call InfoGeneral("  + " & rstRegistroSlave!DenominacionConcepto _
                    & ", con Codigo: " & rstRegistroSlave!CodigoConcepto, ImportacionLiquidacionSueldo)
                    rstRegistroSlave.MoveNext
                Wend
            End If
            rstRegistroSlave.Close
            Set rstRegistroSlave = Nothing
            Call InfoGeneral("", ImportacionLiquidacionSueldo)
        End If
        
        'Insertamos los registros de la Tabla Auxiliar Importación en la Tabla Liquidación Sueldo
        SQL = "Insert Into LIQUIDACIONSUELDOS (CodigoLiquidacion, PuestoLaboral, CodigoConcepto, Importe)" & _
        " Select '" & CodigoLiquidacion & "', PuestoLaboral, CodigoConcepto, Val(Montotxt)" & _
        " From AUXILIARIMPORTACION"
        dbSlave.BeginTrans
        dbSlave.Execute SQL, lngRows
        dbSlave.CommitTrans
        Call InfoGeneral("5) Insertamos " & lngRows & " registros de los datos importados", ImportacionLiquidacionSueldo)
        Call InfoGeneral("", ImportacionLiquidacionSueldo)

        'Borramos la tabla creada
        Call InfoGeneral("6) Borramos tabla AUXILIARIMPORTACION", ImportacionLiquidacionSueldo)
        SQL = "Drop Table AUXILIARIMPORTACION"
        dbSlave.BeginTrans
        dbSlave.Execute SQL
        dbSlave.CommitTrans
'        'Borramos schema.ini
'        Kill strDireccion & "\schema.ini"
        
    End If
    
End Sub

Public Sub ImportarF572Web()
 
    Dim strDireccion As String
    Dim strNombreArchivo As String
    Dim objDOM As MSXML2.DOMDocument
    Dim success As Boolean
    Dim SQL As String
    Dim SQL2 As String

    Dim strTexto As String
    
    With ImportacionF572Web
        .txtInforme.Text = ""
        Call InfoGeneral("- INICIANDO IMPORTACIÓN F. 572 Web -", ImportacionF572Web)
        Call InfoGeneral("", ImportacionF572Web)
        .dlgMultifuncion.Filter = "Todos los XML (*.xml)|*.xml|"
        .dlgMultifuncion.FileName = ""
        .dlgMultifuncion.ShowOpen
        strNombreArchivo = .dlgMultifuncion.FileTitle
        strDireccion = .dlgMultifuncion.FileName
        strDireccion = Replace(strDireccion, "\" & strNombreArchivo, "")
    End With
    
    If strDireccion = "" Then 'En caso de apretar cancelar
        Call InfoGeneral("IMPORTACION CANCELADA", ImportacionF572Web)
        Exit Sub
    Else
        strNombreArchivo = Dir(strDireccion & "\*.xml", vbDirectory)
        Do While strNombreArchivo <> ""   ' Start the loop.
           ' Ignore the current directory and the encompassing directory.
           If strNombreArchivo <> "." And strNombreArchivo <> ".." Then
                'Actualizamos la dirección del archivo
                strDireccion = strDireccion & "\" & strNombreArchivo
                
                'Seteamos la variable
                Set objDOM = New MSXML2.DOMDocument
                
                objDOM.resolveExternals = True
                
                'Para que valide el documento xml
                objDOM.validateOnParse = True
                
                'Carga el documento
                objDOM.async = False
                Call objDOM.Load(strDireccion)
        
                success = objDOM.Load(strDireccion)
                If success = False Then
                   MsgBox objDOM.parseError.reason
                Else
                    'The parser contains two functions to do this: .selectNodes() and .selectSingleNode(), _
                    which retrieve a selection of nodes and the first-found node that satisfies the XPath query, _
                    relative to the node in which the function is called.
                    
                    Dim Nodos As MSXML2.IXMLDOMNodeList
                    Dim oNodo As MSXML2.IXMLDOMNode
                    
                    'Cargamos los datos base de la presentación
                    Dim strPeriodo As String
                    Dim strNroPresentacion As String
                    Dim datFecha As Date
                    Dim strCUIL As String
                    
                    Call InfoGeneral("-Datos Base de la Presentación-", ImportacionF572Web)
                    strTexto = objDOM.selectSingleNode("/presentacion/empleado/apellido").Text _
                    & ", " & objDOM.selectSingleNode("/presentacion/empleado/nombre").Text
                    Call InfoGeneral("Nombre Completo: " & strTexto, ImportacionF572Web)
                    
                    strTexto = objDOM.selectSingleNode("/presentacion/empleado/cuit").Text
                    strCUIL = strTexto
                    Call InfoGeneral("CUIT: " & strTexto, ImportacionF572Web)
                    
                    strTexto = objDOM.selectSingleNode("/presentacion/periodo").Text
                    strPeriodo = strTexto
                    Call InfoGeneral("Período: " & strTexto, ImportacionF572Web)
                    
                    strTexto = objDOM.selectSingleNode("/presentacion/nroPresentacion").Text
                    strNroPresentacion = strTexto
                    Call InfoGeneral("Nro. Presentación: " & strTexto, ImportacionF572Web)
                    
                    strTexto = objDOM.selectSingleNode("/presentacion/fechaPresentacion").Text
                    datFecha = DateTime.DateSerial(Left(strTexto, 4), Mid(strTexto, 6, 2), Right(strTexto, 2))
                    Call InfoGeneral("Fecha Presentación: " & strTexto, ImportacionF572Web)
                    
                    'Verificamos si la ya se Registró la DDJJ
                    SQL = "SELECT * FROM PresentacionSIRADIG " _
                        & "WHERE CUIL = '" & strCUIL & "' " _
                        & "AND RIGHT(ID,2) = '" & Right(strPeriodo, 2) & "' " _
                        & "AND NroPresentacion = '" & strNroPresentacion & "'"
                    If SQLNoMatch(SQL) = True Then
                        Call InfoGeneral("PRESENTACION NUEVA", ImportacionF572Web)
                                
                        Dim strID As String
                        strID = ""
                                
                        'Buscamos el último número de ID utilizado en el período
                        SQL = "SELECT MAX(Left(ID,5)) AS LastID FROM PresentacionSIRADIG " _
                            & "WHERE RIGHT(ID,2) = '" & Right(strPeriodo, 2) & "'"
                        If SQLNoMatch(SQL) = True Then
                            strID = "00001/" & Right(strPeriodo, 2)
                        Else
                            Set rstBuscarSlave = New ADODB.Recordset
                            rstBuscarSlave.Open SQL, dbSlave, adOpenForwardOnly, adLockOptimistic
                            strID = CStr(Val(rstBuscarSlave!LastID) + 1)
                            strID = Format(strID, "00000") & "/" & Right(strPeriodo, 2)
                            rstBuscarSlave.Close
                            Set rstBuscarSlave = Nothing
                        End If
                        'Registramos la carátula de la Presentacion
                        SQL = "INSERT INTO PresentacionSIRADIG " _
                            & "(ID, CUIL, Fecha, NroPresentacion) " _
                            & "VALUES( '" & strID & "' , '" & strCUIL _
                            & "' , #" & Format(datFecha, "MM/DD/YYYY") _
                            & "# , '" & strNroPresentacion & "')"
                        dbSlave.BeginTrans
                        dbSlave.Execute SQL
                        dbSlave.CommitTrans
                        
                        strPeriodo = ""
                        strNroPresentacion = ""
                        datFecha = 0
                        strCUIL = ""
                        SQL = ""
                        Call InfoGeneral("", ImportacionF572Web)
                        
                        'Verificamos las Cargas de Familia de la DDJJ que se está registrando
                        Set Nodos = objDOM.selectNodes("/presentacion/cargasFamilia/cargaFamilia")
                        Call InfoGeneral("-Cargas de Familia (" & Nodos.Length & ")-", ImportacionF572Web)
                        If Nodos.Length > 0 Then
                            Dim strCUILFamiliar As String
                            Dim strMesDesdeCF As String
                            Dim strMesHastaCF As String
                            Dim strCodigoParentesco As String
                            Dim intProximoPeriodo As Integer
                            Dim datFechaNacimiento As Date
                            For Each oNodo In Nodos
                                'Datos del Familiar
                                strTexto = oNodo.selectSingleNode("apellido").Text _
                                & " " & oNodo.selectSingleNode("nombre").Text _
                                & " (" & oNodo.selectSingleNode("nroDoc").Text & ") - Próximo:" _
                                & " " & oNodo.selectSingleNode("vigenteProximosPeriodos").Text
                                Call InfoGeneral(" + Nombre Completo: " & strTexto, ImportacionF572Web)
                                'Preseteamos las variables
                                strCUILFamiliar = oNodo.selectSingleNode("nroDoc").Text
                                If Len(strCUILFamiliar) < 11 Then
                                    strCUILFamiliar = "00" & strCUILFamiliar & "0"
                                End If
                                strMesDesde = oNodo.selectSingleNode("mesDesde").Text
                                strMesHasta = oNodo.selectSingleNode("mesHasta").Text
                                strCodigoParentesco = oNodo.selectSingleNode("parentesco").Text
                                strTexto = oNodo.selectSingleNode("fechaNac").Text
                                datFechaNacimiento = DateTime.DateSerial(Left(strTexto, 4), Mid(strTexto, 6, 2), Right(strTexto, 2))
                                strTexto = oNodo.selectSingleNode("vigenteProximosPeriodos").Text
                                If strTexto = "S" Then
                                    intProximoPeriodo = -1
                                Else
                                    intProximoPeriodo = 0
                                End If
                                'Insertamos el familiar asociado a la Presentacion F572
                                SQL = "INSERT INTO CargasFamiliaSIRADIG " _
                                    & "(ID, CUIL, MesDesde, MesHasta, CodigoParentesco, ProximoPeriodo, FechaNacimiento) " _
                                    & "VALUES( '" & strID & "' , '" & strCUILFamiliar _
                                    & "' , '" & strMesDesde & "' , '" & strMesHasta _
                                    & "' , '" & strCodigoParentesco & "' , " & intProximoPeriodo _
                                    & ", #" & Format(datFechaNacimiento, "MM/DD/YYYY") & "#)"
                                dbSlave.BeginTrans
                                dbSlave.Execute SQL
                                dbSlave.CommitTrans
                            Next oNodo
                        End If
                        strCUILFamiliar = ""
                        strMesDesdeCF = ""
                        strMesHastaCF = ""
                        strCodigoParentesco = ""
                        intProximoPeriodo = 0
                        datFechaNacimiento = 0
                        SQL = ""
                        Call InfoGeneral("", ImportacionF572Web)
                        
                        'Verificamos las Ganancias liquidadas por otros empleadores de la DDJJ que se está registrando
                        Set Nodos = objDOM.selectNodes("/presentacion/ganLiqOtrosEmpEnt/empEnt")
                        If Nodos.Length > 0 Then
                            strTexto = "Sí (CASO NO CONTEMPLADO)"
                        Else
                            strTexto = "No"
                        End If
                        Call InfoGeneral("-¿Tiene Ganancias liquidadas por otros empleadores?- " & strTexto, ImportacionF572Web)
'                        For Each oNodo In Nodos
'                            'Datos de las deducciones
'                            strTexto = oNodo.selectSingleNode("apellido").Text _
'                            & oNodo.selectSingleNode("nombre").Text _
'                            & " (" & oNodo.selectSingleNode("nroDoc").Text & ")"
'                            Call InfoGeneral(" + Nombre Completo: " & strTexto, ImportacionF572Web)
'                        Next oNodo
                        Call InfoGeneral("", ImportacionF572Web)
                        
                        'Verificamos las Deducciones Generales de la DDJJ que se está registrando
                        Set Nodos = objDOM.selectNodes("/presentacion/deducciones/deduccion")
                        If Nodos.Length > 0 Then
                            strTexto = "Sí, " & Nodos.Length
                            Dim NodosDeNodos As MSXML2.IXMLDOMNodeList
                            Dim oNodo2 As MSXML2.IXMLDOMNode
                            Dim strControlAtributo As String
                            strControlAtributo = "0"
                        Else
                            strTexto = "No"
                        End If
                        Call InfoGeneral("-¿Informa Deducciones Generales?- " & strTexto, ImportacionF572Web)
                        For Each oNodo In Nodos
                            'Datos Deducciones Generales
                            Dim strMesDesdeDG As String
                            Dim strMesHastaDG As String
                            Dim dblImporte As Double
                            Dim arrImporte(1 To 12) As Double
                            
                            strTexto = oNodo.Attributes.getNamedItem("tipo").Text
                            If strControlAtributo <> strTexto Then
                                'Preseteamos el array de importes
                                For i = 1 To 12
                                    arrImporte(i) = 0
                                Next i
                                Call InfoGeneral(" + Tipo: " & strTexto, ImportacionF572Web)
                                strControlAtributo = strTexto
                                Set NodosDeNodos = objDOM.selectNodes("/presentacion/deducciones/deduccion[@tipo='" & strTexto & "']/periodos/periodo")
                                For Each oNodo2 In NodosDeNodos
                                    'Preseteamos las variables
                                    strMesDesdeDG = oNodo2.Attributes.getNamedItem("mesDesde").Text
                                    strMesHastaDG = oNodo2.Attributes.getNamedItem("mesHasta").Text
                                    dblImporte = Val(oNodo2.Attributes.getNamedItem("montoMensual").Text)
                                    
                                    strTexto = "Desde " & strMesDesdeDG _
                                    & " Hasta " & strMesHastaDG _
                                    & " $ " & CStr(dblImporte)
                                    Call InfoGeneral(" ++ Periodo(mm): " & strTexto, ImportacionF572Web)
                                    
                                    'Asignamos valores al Array
                                    For i = strMesDesdeDG To strMesHastaDG
                                        arrImporte(i) = arrImporte(i) + dblImporte
                                    Next i

                                Next oNodo2
                                'Insertamos en la Tabla DeduccionesGeneralesSIRADIG
                                SQL = ""
                                For i = 1 To 12
                                    SQL = SQL & ", " & De_Num_a_Tx_01(arrImporte(i), False, 2)
                                Next i
                                SQL = "INSERT INTO DeduccionesGeneralesSIRADIG " _
                                    & "(ID, CodigoDeduccion, Mes01, Mes02, Mes03, Mes04, " _
                                    & "Mes05, Mes06, Mes07, Mes08, Mes09, Mes10, Mes11, Mes12) " _
                                    & "VALUES( '" & strID & "' , '" & strControlAtributo & "' " & SQL & ")"
                                'Debug.Print SQL
                                dbSlave.BeginTrans
                                'Debug.Print SQL
                                dbSlave.Execute SQL
                                dbSlave.CommitTrans
                            End If
                        Next oNodo
                        For i = 1 To 12
                            arrImporte(i) = 0
                        Next i
                        dblImporte = 0
                        strMesDesdeDG = ""
                        strMesHastaDG = ""
                        SQL = ""
                        Call InfoGeneral("", ImportacionF572Web)
                        
                        'Verificamos las Retenciones, Percepciones y Pago a Cuenta
                        Set Nodos = objDOM.selectNodes("/presentacion/retPerPagos/retPerPago")
                        If Nodos.Length > 0 Then
                            strTexto = "Sí (CASO NO CONTEMPLADO)"
                        Else
                            strTexto = "No"
                        End If
                        Call InfoGeneral("-¿Tiene Retenciones, Percepciones y Pago a Cuenta?- " & strTexto, ImportacionF572Web)
'                        For Each oNodo In Nodos
'                            'Datos de las deducciones
'                            strTexto = oNodo.selectSingleNode("apellido").Text _
'                            & oNodo.selectSingleNode("nombre").Text _
'                            & " (" & oNodo.selectSingleNode("nroDoc").Text & ")"
'                            Call InfoGeneral(" + Nombre Completo: " & strTexto, ImportacionF572Web)
'                        Next oNodo
                        Call InfoGeneral("", ImportacionF572Web)
                        
                        'Verificamos los Ajustes
                        Set Nodos = objDOM.selectNodes("/presentacion/ajustes/ajuste")
                        If Nodos.Length > 0 Then
                            strTexto = "Sí (CASO NO CONTEMPLADO)"
                        Else
                            strTexto = "No"
                        End If
                        Call InfoGeneral("-¿Tiene Ajustes?- " & strTexto, ImportacionF572Web)
'                        For Each oNodo In Nodos
'                            'Datos de las deducciones
'                            strTexto = oNodo.selectSingleNode("apellido").Text _
'                            & oNodo.selectSingleNode("nombre").Text _
'                            & " (" & oNodo.selectSingleNode("nroDoc").Text & ")"
'                            Call InfoGeneral(" + Nombre Completo: " & strTexto, ImportacionF572Web)
'                        Next oNodo
                        Call InfoGeneral("", ImportacionF572Web)
                        
                    Else
                        Call InfoGeneral("PRESENTACION YA REGISTRADA", ImportacionF572Web)
                    End If
                    Call InfoGeneral("", ImportacionF572Web)
                
                End If
                strDireccion = Replace(strDireccion, "\" & strNombreArchivo, "")
                Set objDOM = Nothing
    
           End If
           strNombreArchivo = Dir()   ' Get next entry.
        Loop


    End If
    
End Sub

Public Sub ImportarLiquidacionHonorariosSLAVE()

    Dim strDireccion        As String
    Dim intNumeroArchivo    As Integer
    Dim in_sTableName       As String
    Dim SQL_InsertPrefix    As String
    Dim SQL                 As String
    Dim intIndex            As Integer
    Dim strComprobante      As String
    Dim sLine               As String
    Dim vasFields           As Variant
    Dim datFecha            As Date
    Dim dblImporte          As Double
    Dim i                   As Long
    
    in_sTableName = "LiquidacionHonorarios"
    
    With ImportacionLiquidacionHonorarios
        .txtInforme.Text = ""
        Call InfoGeneral("- INICIANDO IMPORTACIÓN HONORARIOS -", ImportacionLiquidacionHonorarios)
        Call InfoGeneral("", ImportacionLiquidacionHonorarios)
        .dlgMultifuncion.Filter = "Todos los Excel(*.csv)|*.csv|"
        .dlgMultifuncion.FileName = ""
        .dlgMultifuncion.ShowOpen
        strDireccion = .dlgMultifuncion.FileName
    End With
    If strDireccion = "" Then 'En caso de apretar cancelar
        Call InfoGeneral("IMPORTACION CANCELADA", ImportacionLiquidacionHonorarios)
        Exit Sub
    Else
        intNumeroArchivo = FreeFile
        Open strDireccion For Input As #intNumeroArchivo
        strDireccion = ""
        'Verificamos si el archivo esta vacio
        If EOF(intNumeroArchivo) Then
            Close #intNumeroArchivo
            Exit Sub
        End If

        ' Read field names from the top line.
        'Line Input #iFileNo, sLine
    
        ' Note that in a "proper" CSV file, there should not be any trailing spaces, and all strings should be surrounded by double quotes.
        ' However, the top row provided can simply be used "as is" in the SQL string.
        SQL_InsertPrefix = "INSERT INTO " & in_sTableName & "" _
        & " (Fecha, Proveedor, Sellos, Seguro, Comprobante, Tipo, MontoBruto, IIBB, LibramientoPago, OtraRetencion, Anticipo, Descuento, Actividad, Partida)" _
        & " VALUES("
        
        i = 1
        
        Do Until EOF(intNumeroArchivo)
            
            Line Input #intNumeroArchivo, sLine
    
            ' Initialise SQL string.
            SQL = SQL_InsertPrefix
    
            ' Again, in a proper CSV file, the spaces around the commas shouldn't be there, and the fields should be double quoted.
            ' Since the data is supposed to look like this, then the assumption is that the delimiter is space-comma-space.
            vasFields = Split(sLine, Chr(34) & "," & Chr(34))
    
            ' Build up each value, separated by comma.
            ' It is assumed that all values here are string, so they will be double quoted.
            ' However, if there are non-string values, you will have to ensure they don't get quoted.
            
            'Capturamos la fecha hasta del período y la formateamos
            vasFields(0) = Right(vasFields(0), Len(vasFields(0)) - 1)
            datFecha = DateTime.DateSerial(Right(vasFields(0), 4), Mid(vasFields(0), 4, 2), Left(vasFields(0), 2))
            SQL = SQL & "#" & Format(datFecha, "MM/DD/YYYY") & "#,"
            'Capturamos el nombre del facturero
            SQL = SQL & "'" & vasFields(1) & "',"
            'Capturamos el importe de Sellos y la formateamos
            dblImporte = De_Txt_a_Num_01(vasFields(2), 2, ".")
            SQL = SQL & De_Num_a_Tx_01(dblImporte) & ","
            'Capturamos el importe de Seguro y la formateamos
            dblImporte = De_Txt_a_Num_01(vasFields(3), 2, ".")
            SQL = SQL & De_Num_a_Tx_01(dblImporte) & ","
            'Capturamos el Nro de Comprobante
            SQL = SQL & "'" & vasFields(4) & "',"
            'Capturamos el Tipo
            SQL = SQL & "'" & vasFields(5) & "',"
            'Capturamos el importe de Monto Bruto y la formateamos
            dblImporte = De_Txt_a_Num_01(vasFields(6), 2, ".")
            SQL = SQL & De_Num_a_Tx_01(dblImporte) & ","
            'Capturamos el importe de IIBB y la formateamos
            dblImporte = De_Txt_a_Num_01(vasFields(7), 2, ".")
            SQL = SQL & De_Num_a_Tx_01(dblImporte) & ","
            'Capturamos el importe Libramiento Pago
            dblImporte = De_Txt_a_Num_01(vasFields(8), 2, ".")
            SQL = SQL & De_Num_a_Tx_01(dblImporte) & ","
            'Capturamos el importe de Otras Retenciones y la formateamos
            dblImporte = De_Txt_a_Num_01(vasFields(9), 2, ".")
            SQL = SQL & De_Num_a_Tx_01(dblImporte) & ","
            'Capturamos el importe de Anticipo y la formateamos
            dblImporte = De_Txt_a_Num_01(vasFields(10), 2, ".")
            SQL = SQL & De_Num_a_Tx_01(dblImporte) & ","
            'Capturamos el importe de Descuento
            dblImporte = De_Txt_a_Num_01(vasFields(11), 2, ".")
            SQL = SQL & De_Num_a_Tx_01(dblImporte) & ","
            'Capturamos la Actividad
            SQL = SQL & "'" & vasFields(12) & "',"
            'Capturamos la Partida
            SQL = SQL & "'" & Left(vasFields(13), Len(vasFields(13)) - 1) & "')"
            
            'Borramos toda los registros del año a importar
            If i = 1 Then
                dbSlave.BeginTrans
                dbSlave.Execute "Delete From LIQUIDACIONHONORARIOS Where Year(Fecha) = " & "'" & Year(datFecha) & "'"
                dbSlave.CommitTrans
            End If
            
            'Importamos un registro
            dbSlave.BeginTrans
            dbSlave.Execute SQL
            dbSlave.CommitTrans
            
            i = i + 1
            
        Loop
        
        Call InfoGeneral("", ImportacionLiquidacionHonorarios)
        Call InfoGeneral("- IMPORTACION FINALIZADA -", ImportacionLiquidacionHonorarios)
        Close #intNumeroArchivo
        SQL = ""
        SQL_InsertPrefix = ""
        sLine = ""
        dblImporte = 0
        datFecha = 0
    End If
    
End Sub

Public Sub ImportarPrecarizadosSLAVE()

    Dim strDireccion        As String
    Dim intNumeroArchivo    As Integer
    Dim in_sTableName       As String
    Dim SQL_InsertPrefix    As String
    Dim SQL                 As String
    Dim intIndex            As Integer
    Dim strComprobante      As String
    Dim sLine               As String
    Dim vasFields           As Variant
    Dim datFecha            As Date
    Dim dblImporte          As Double
    
    in_sTableName = "Precarizados"
    
    With ImportacionLiquidacionHonorarios
        .txtInforme.Text = ""
        Call InfoGeneral("- INICIANDO IMPORTACIÓN HONORARIOS -", ImportacionLiquidacionHonorarios)
        Call InfoGeneral("", ImportacionLiquidacionHonorarios)
        .dlgMultifuncion.Filter = "Todos los Excel(*.csv)|*.csv|"
        .dlgMultifuncion.FileName = ""
        .dlgMultifuncion.ShowOpen
        strDireccion = .dlgMultifuncion.FileName
    End With
    If strDireccion = "" Then 'En caso de apretar cancelar
        Call InfoGeneral("IMPORTACION CANCELADA", ImportacionLiquidacionHonorarios)
        Exit Sub
    Else
        intNumeroArchivo = FreeFile
        Open strDireccion For Input As #intNumeroArchivo
        strDireccion = ""
        'Verificamos si el archivo esta vacio
        If EOF(intNumeroArchivo) Then
            Close #intNumeroArchivo
            Exit Sub
        End If

        'Borramos la base de datos de Precarizados para importar la nueva base
        dbSlave.BeginTrans
        dbSlave.Execute "Delete From PRECARIZADOS"
        dbSlave.CommitTrans


        ' Read field names from the top line.
        'Line Input #iFileNo, sLine
    
        ' Note that in a "proper" CSV file, there should not be any trailing spaces, and all strings should be surrounded by double quotes.
        ' However, the top row provided can simply be used "as is" in the SQL string.
        SQL_InsertPrefix = "INSERT INTO " & in_sTableName & "" _
        & " (Agentes, Actividad, Partida) VALUES("
        
        Do Until EOF(intNumeroArchivo)
            
            Line Input #intNumeroArchivo, sLine
    
            ' Initialise SQL string.
            SQL = SQL_InsertPrefix
    
            ' Again, in a proper CSV file, the spaces around the commas shouldn't be there, and the fields should be double quoted.
            ' Since the data is supposed to look like this, then the assumption is that the delimiter is space-comma-space.
            vasFields = Split(sLine, Chr(34) & "," & Chr(34))
    
            ' Build up each value, separated by comma.
            ' It is assumed that all values here are string, so they will be double quoted.
            ' However, if there are non-string values, you will have to ensure they don't get quoted.
            
            'Capturamos el nombre del facturero
            SQL = SQL & "'" & Right(vasFields(0), Len(vasFields(0)) - 1) & "',"
            'Capturamos la Actividad
            SQL = SQL & "'" & vasFields(1) & "',"
            'Capturamos la Partida
            SQL = SQL & "'" & Left(vasFields(2), Len(vasFields(2)) - 1) & "')"
            
            'Importamos un registro
            dbSlave.BeginTrans
            dbSlave.Execute SQL
            dbSlave.CommitTrans
            
        Loop
        
        Call InfoGeneral("", ImportacionLiquidacionHonorarios)
        Call InfoGeneral("- IMPORTACION FINALIZADA -", ImportacionLiquidacionHonorarios)
        Close #intNumeroArchivo
        SQL = ""
        SQL_InsertPrefix = ""
        sLine = ""
        dblImporte = 0
        datFecha = 0
    End If
    
End Sub
