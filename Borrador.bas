Attribute VB_Name = "Borrador"
'Private Sub BorradorCargardgDeduccionesGeneralesLG4ta()
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
'        If TieneDeduccionGeneral("SERVICIODOMESTICO", LiquidacionGanancia4ta.txtPuestoLaboral.Text, datFecha) = False Then
'            .TextMatrix(1, 1) = De_Num_a_Tx_01(0)
'        Else
'            'Controlamos lo que debió liquidarse en concepto de Servicio Doméstico
'            dblALiquidarse = ImporteDeduccionGeneral(LiquidacionGanancia4ta.txtPuestoLaboral.Text, "SERVICIODOMESTICO", datFecha)
'            dblALiquidarse = dblALiquidarse * Month(datFecha)
'            'Controlamos lo que realmente se liquidó en concepto de Servicio Doméstico
'            SQL = "Select  Sum(LIQUIDACIONGANANCIAS4TACATEGORIA.ServicioDomestico) AS SumaDeServicioDomestico " _
'            & "From CODIGOLIQUIDACIONES Inner Join LIQUIDACIONGANANCIAS4TACATEGORIA On CODIGOLIQUIDACIONES.Codigo = LIQUIDACIONGANANCIAS4TACATEGORIA.CodigoLiquidacion " _
'            & "Where PUESTOLABORAL = '" & LiquidacionGanancia4ta.txtPuestoLaboral.Text & "' " _
'            & "And Right(PERIODO,4) = '" & Right(LiquidacionGanancia4ta.txtPeriodo.Text, 4) & "' And CODIGO < '" & LiquidacionGanancia4ta.txtCodigoLiquidacion.Text & "'"
'            rstListadoSlave.Open SQL, dbSlave, adOpenDynamic, adLockOptimistic
'            If rstListadoSlave.BOF = False And IsNull(rstListadoSlave!SumaDeServicioDomestico) = False Then
'                dblLiquidado = rstListadoSlave!SumaDeServicioDomestico
'                dblAcumulado = dblAcumulado + dblLiquidado
'            Else
'                dblLiquidado = 0
'            End If
'            rstListadoSlave.Close
'            'Cargamos el Importe de la diferencia entre lo que debió liquidarse y lo que se liquidó en concepto de Servicio Doméstico
'            .TextMatrix(1, 1) = De_Num_a_Tx_01(dblALiquidarse - dblLiquidado, , 2)
'            dblTotal = dblTotal + De_Txt_a_Num_01(.TextMatrix(1, 1))
'            dblALiquidarse = 0
'        End If
'
'        .TextMatrix(2, 0) = "Seguro de Vida"
'        If TieneDeduccionGeneral("SEGURODEVIDA", LiquidacionGanancia4ta.txtPuestoLaboral.Text, datFecha) = False Then
'            .TextMatrix(2, 1) = De_Num_a_Tx_01(0)
'            'Controlamos lo que realmente se liquidó en concepto de Seguro de Vida
'            SQL = "Select  Sum(LIQUIDACIONGANANCIAS4TACATEGORIA.SeguroDeVidaOptativo) AS SumaDeSeguroDeVida " _
'            & "From CODIGOLIQUIDACIONES Inner Join LIQUIDACIONGANANCIAS4TACATEGORIA On CODIGOLIQUIDACIONES.Codigo = LIQUIDACIONGANANCIAS4TACATEGORIA.CodigoLiquidacion " _
'            & "Where PUESTOLABORAL = '" & LiquidacionGanancia4ta.txtPuestoLaboral.Text & "' " _
'            & "And Right(PERIODO,4) = '" & Right(LiquidacionGanancia4ta.txtPeriodo.Text, 4) & "' And CODIGO < '" & LiquidacionGanancia4ta.txtCodigoLiquidacion.Text & "'"
'            rstListadoSlave.Open SQL, dbSlave, adOpenDynamic, adLockOptimistic
'            If rstListadoSlave.BOF = False And IsNull(rstListadoSlave!SumaDeSeguroDeVida) = False Then
'                dblAcumulado = dblAcumulado + rstListadoSlave!SumaDeSeguroDeVida
'            End If
'            rstListadoSlave.Close
'        Else
'            'Controlamos lo que debió liquidarse en concepto de Seguro de Vida
'            dblALiquidarse = ImporteDeduccionGeneral(LiquidacionGanancia4ta.txtPuestoLaboral.Text, "SEGURODEVIDA", datFecha, , De_Txt_a_Num_01(LiquidacionGanancia4ta.txtSeguroOptativo.Text))
'            dblALiquidarse = dblALiquidarse * Month(datFecha)
'            'Controlamos lo que realmente se liquidó en concepto de Seguro de Vida
'            SQL = "Select  Sum(LIQUIDACIONGANANCIAS4TACATEGORIA.SeguroDeVidaOptativo) AS SumaDeSeguroDeVida " _
'            & "From CODIGOLIQUIDACIONES Inner Join LIQUIDACIONGANANCIAS4TACATEGORIA On CODIGOLIQUIDACIONES.Codigo = LIQUIDACIONGANANCIAS4TACATEGORIA.CodigoLiquidacion " _
'            & "Where PUESTOLABORAL = '" & LiquidacionGanancia4ta.txtPuestoLaboral.Text & "' " _
'            & "And Right(PERIODO,4) = '" & Right(LiquidacionGanancia4ta.txtPeriodo.Text, 4) & "' And CODIGO < '" & LiquidacionGanancia4ta.txtCodigoLiquidacion.Text & "'"
'            rstListadoSlave.Open SQL, dbSlave, adOpenDynamic, adLockOptimistic
'            If rstListadoSlave.BOF = False And IsNull(rstListadoSlave!SumaDeSeguroDeVida) = False Then
'                dblLiquidado = rstListadoSlave!SumaDeSeguroDeVida
'                dblAcumulado = dblAcumulado + dblLiquidado
'            Else
'                dblLiquidado = 0
'            End If
'            rstListadoSlave.Close
'            'Cargamos el Importe de la diferencia entre lo que debió liquidarse y lo que se liquidó en concepto de Seguro De Vida _
'            teniendo en cuenta lo que se consideró como deducción
'            dblALiquidarse = (dblALiquidarse - dblLiquidado)
'            .TextMatrix(2, 1) = De_Num_a_Tx_01(dblALiquidarse - De_Txt_a_Num_01(LiquidacionGanancia4ta.txtSeguroOptativo.Text), , 2)
'            dblTotal = dblTotal + De_Txt_a_Num_01(.TextMatrix(2, 1))
'            dblALiquidarse = 0
'        End If
'
'        .TextMatrix(3, 0) = "Cuota Médica Asist."
'        If TieneDeduccionGeneral("CUOTAMEDICOASISTENCIAL", LiquidacionGanancia4ta.txtPuestoLaboral.Text, datFecha) = False Then
'            .TextMatrix(3, 1) = De_Num_a_Tx_01(0)
'        Else
'            'Controlamos lo que debió liquidarse en concepto de Cuota Médico Asistencial
'            dblGananciaNeta = (De_Txt_a_Num_01(LiquidacionGanancia4ta.txtGananciaNeta.Text) - dblAcumulado _
'            - De_Txt_a_Num_01(.TextMatrix(2, 1)) - De_Txt_a_Num_01(.TextMatrix(1, 1)) - De_Txt_a_Num_01(LiquidacionGanancia4ta.txtSeguroOptativo.Text))
'            dblALiquidarse = ImporteDeduccionGeneral(LiquidacionGanancia4ta.txtPuestoLaboral.Text, "CUOTAMEDICOASISTENCIAL", datFecha, dblGananciaNeta)
'            'Controlamos lo que realmente se liquidó en concepto de Mínimo no Imponible
'            SQL = "Select  Sum(LIQUIDACIONGANANCIAS4TACATEGORIA.CuotaMedicoAsistencial) AS SumaDeCuotaMedicoAsistencial " _
'            & "From CODIGOLIQUIDACIONES Inner Join LIQUIDACIONGANANCIAS4TACATEGORIA On CODIGOLIQUIDACIONES.Codigo = LIQUIDACIONGANANCIAS4TACATEGORIA.CodigoLiquidacion " _
'            & "Where PUESTOLABORAL = '" & LiquidacionGanancia4ta.txtPuestoLaboral.Text & "' " _
'            & "And Right(PERIODO,4) = '" & Right(LiquidacionGanancia4ta.txtPeriodo.Text, 4) & "' And CODIGO < '" & LiquidacionGanancia4ta.txtCodigoLiquidacion.Text & "'"
'            rstListadoSlave.Open SQL, dbSlave, adOpenDynamic, adLockOptimistic
'            If rstListadoSlave.BOF = False And IsNull(rstListadoSlave!SumaDeCuotaMedicoAsistencial) = False Then
'                dblLiquidado = rstListadoSlave!SumaDeCuotaMedicoAsistencial
'                dblAcumulado = dblAcumulado + dblLiquidado
'            Else
'                dblLiquidado = 0
'            End If
'            rstListadoSlave.Close
'            'Cargamos el Importe de la diferencia entre lo que debió liquidarse y lo que se liquidó en concepto de Cuota Médico Asistencial
'            .TextMatrix(3, 1) = De_Num_a_Tx_01(dblALiquidarse - dblLiquidado, , 2)
'            dblTotal = dblTotal + De_Txt_a_Num_01(.TextMatrix(3, 1))
'            dblALiquidarse = 0
'        End If
'
'        .TextMatrix(4, 0) = "Donaciones"
'        If TieneDeduccionGeneral("DONACIONES", LiquidacionGanancia4ta.txtPuestoLaboral.Text, datFecha) = False Then
'            .TextMatrix(4, 1) = De_Num_a_Tx_01(0)
'        Else
'            'Controlamos lo que debió liquidarse en concepto de Donaciones
'            If dblGananciaNeta = 0 Then
'                dblGananciaNeta = (De_Txt_a_Num_01(LiquidacionGanancia4ta.txtGananciaNeta.Text) - dblAcumulado _
'                - De_Txt_a_Num_01(.TextMatrix(2, 1)) - De_Txt_a_Num_01(.TextMatrix(1, 1)) - De_Txt_a_Num_01(LiquidacionGanancia4ta.txtSeguroOptativo.Text))
'            End If
'            dblALiquidarse = ImporteDeduccionGeneral(LiquidacionGanancia4ta.txtPuestoLaboral.Text, "CUOTAMEDICOASISTENCIAL", datFecha, dblGananciaNeta)
'            'Controlamos lo que realmente se liquidó en concepto de Donaciones
'            SQL = "Select  Sum(LIQUIDACIONGANANCIAS4TACATEGORIA.Donaciones) AS SumaDeDonaciones " _
'            & "From CODIGOLIQUIDACIONES Inner Join LIQUIDACIONGANANCIAS4TACATEGORIA On CODIGOLIQUIDACIONES.Codigo = LIQUIDACIONGANANCIAS4TACATEGORIA.CodigoLiquidacion " _
'            & "Where PUESTOLABORAL = '" & LiquidacionGanancia4ta.txtPuestoLaboral.Text & "' " _
'            & "And Right(PERIODO,4) = '" & Right(LiquidacionGanancia4ta.txtPeriodo.Text, 4) & "' And CODIGO < '" & LiquidacionGanancia4ta.txtCodigoLiquidacion.Text & "'"
'            rstListadoSlave.Open SQL, dbSlave, adOpenDynamic, adLockOptimistic
'            If rstListadoSlave.BOF = False And IsNull(rstListadoSlave!SumaDeDonaciones) = False Then
'                dblLiquidado = rstListadoSlave!SumaDeDonaciones
'                dblAcumulado = dblAcumulado + dblLiquidado
'            Else
'                dblLiquidado = 0
'            End If
'            rstListadoSlave.Close
'            'Cargamos el Importe de la diferencia entre lo que debió liquidarse y lo que se liquidó en concepto de Donaciones
'            .TextMatrix(4, 1) = De_Num_a_Tx_01(dblALiquidarse - dblLiquidado, , 2)
'            dblTotal = dblTotal + De_Txt_a_Num_01(.TextMatrix(4, 1))
'            dblALiquidarse = 0
'        End If
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
'
'End Sub

Private Sub BorradorProcedimientoLiquidacion()

    'Determinamos la Retención Acumulada
    'SQL = "Select Sum(LIQUIDACIONSUELDOS.Importe) AS SumaDeImporte " _
    & "From CODIGOLIQUIDACIONES INNER JOIN LIQUIDACIONSUELDOS ON CODIGOLIQUIDACIONES.Codigo = LIQUIDACIONSUELDOS.CodigoLiquidacion " _
    & "Where PUESTOLABORAL = '" & .txtPuestoLaboral.Text & "' And CODIGOCONCEPTO= '0276' " _
    & "And Right(PERIODO,4) = '" & Right(.txtPeriodo.Text, 4) & "' And CODIGO < '" & .txtCodigoLiquidacion.Text & "'"
    'rstRegistroSlave.Open SQL, dbSlave, adOpenForwardOnly, adLockOptimistic
    'If rstRegistroSlave.BOF = False And IsNull(rstRegistroSlave!SumaDeImporte) = False Then
    '    .txtRetencionAcumulada.Text = De_Num_a_Tx_01(rstRegistroSlave!SumaDeImporte, , 2)
    'Else
    '    .txtRetencionAcumulada.Text = De_Num_a_Tx_01(0)
    'End If
    'rstRegistroSlave.Close

End Sub

'Private Sub ViejoProcedimientoIncrementoDeduccionEspecial()
'
'    'Decreto Especiales del PEN (PENSAR EN UNA ALTERNATIVA MEJOR!!!!!)
'    Set rstProcedimientoSlave = New ADODB.Recordset
'    'Decreto 1006/13
'    If Year(datFecha) = 2013 And Month(datFecha) >= 7 Then
'        SQL = "Select * From LIQUIDACIONSUELDOS" _
'        & " Where CODIGOLIQUIDACION = '0482'" _
'        & " And PUESTOLABORAL = '" & PuestoLaboral & "'"
'        If SQLNoMatch(SQL) = False Then
'            'Buscamos el SAC Bruto
'            SQL = "Select * From LIQUIDACIONSUELDOS Where CODIGOLIQUIDACION = '0482' " _
'            & "And PUESTOLABORAL = '" & PuestoLaboral & "' " _
'            & "And CODIGOCONCEPTO = '0150'"
'            If SQLNoMatch(SQL) = False Then
'                rstProcedimientoSlave.Open SQL, dbSlave, adOpenDynamic, adLockOptimistic
'                dblImporteMensual = rstListadoSlave!Importe
'                rstProcedimientoSlave.Close
'            End If
'            'Buscamos la Retención Jubilación
'            SQL = "Select * From LIQUIDACIONSUELDOS Where CODIGOLIQUIDACION = '0482' " _
'            & "And PUESTOLABORAL = '" & PuestoLaboral & "' " _
'            & "And CODIGOCONCEPTO = '0208'"
'            If SQLNoMatch(SQL) = False Then
'                rstProcedimientoSlave.Open SQL, dbSlave, adOpenDynamic, adLockOptimistic
'                dblImporteMensual = dblImporteMensual - rstListadoSlave!Importe
'                rstProcedimientoSlave.Close
'            End If
'            'Buscamos la Retención Obra Social
'            SQL = "Select * From LIQUIDACIONSUELDOS Where CODIGOLIQUIDACION = '0482' " _
'            & "And PUESTOLABORAL = '" & PuestoLaboral & "' " _
'            & "And CODIGOCONCEPTO = '0212'"
'            If SQLNoMatch(SQL) = False Then
'                rstProcedimientoSlave.Open SQL, dbSlave, adOpenDynamic, adLockOptimistic
'                dblImporteMensual = dblImporteMensual - rstListadoSlave!Importe
'                rstProcedimientoSlave.Close
'            End If
'            'Buscamos la Retención Aporte Sindical
'            SQL = "Select * from LIQUIDACIONSUELDOS Where " _
'            & "((PUESTOLABORAL = '" & PuestoLaboral & "' " _
'            & "And CODIGOLIQUIDACION = '0482' And CODIGOCONCEPTO = '0219') Or " _
'            & "(PUESTOLABORAL = '" & PuestoLaboral & "' " _
'            & "And CODIGOLIQUIDACION = '0482' And CODIGOCONCEPTO = '0227'))"
'            If SQLNoMatch(SQL) = False Then
'                rstProcedimientoSlave.Open SQL, dbSlave, adOpenDynamic, adLockOptimistic
'                dblImporteMensual = dblImporteMensual - rstListadoSlave!Importe
'                rstProcedimientoSlave.Close
'            End If
'            'Buscamos el Pluriempleo
'            SQL = "Select * from LIQUIDACIONGANANCIAS4TACATEGORIA Where " _
'            & "(PUESTOLABORAL = '" & PuestoLaboral & "' " _
'            & "And CODIGOLIQUIDACION = '0482')"
'            If SQLNoMatch(SQL) = False Then
'                rstProcedimientoSlave.Open SQL, dbSlave, adOpenDynamic, adLockOptimistic
'                dblImporteMensual = dblImporteMensual + rstListadoSlave!Pluriempleo
'                rstProcedimientoSlave.Close
'            End If
'            dblImporteCalculado = dblImporteCalculado + dblImporteMensual
'        End If
'    End If
'    'Decreto 2354/14
'    If Year(datFecha) = 2014 And Month(datFecha) >= 12 Then
'        SQL = "Select * From LIQUIDACIONSUELDOS Where CODIGOLIQUIDACION = '0524' " _
'        & "And PUESTOLABORAL = '" & PuestoLaboral & "'"
'        If SQLNoMatch(SQL) = False Then
'            'Buscamos el SAC Bruto
'            SQL = "Select * From LIQUIDACIONSUELDOS Where CODIGOLIQUIDACION = '0524' " _
'            & "And PUESTOLABORAL = '" & PuestoLaboral & "' " _
'            & "And CODIGOCONCEPTO = '0150'"
'            If SQLNoMatch(SQL) = False Then
'                rstProcedimientoSlave.Open SQL, dbSlave, adOpenDynamic, adLockOptimistic
'                dblImporteMensual = rstListadoSlave!Importe
'                rstProcedimientoSlave.Close
'            End If
'            'Buscamos la Retención Jubilación
'            SQL = "Select * From LIQUIDACIONSUELDOS Where CODIGOLIQUIDACION = '0524' " _
'            & "And PUESTOLABORAL = '" & PuestoLaboral & "' " _
'            & "And CODIGOCONCEPTO = '0208'"
'            If SQLNoMatch(SQL) = False Then
'                rstProcedimientoSlave.Open SQL, dbSlave, adOpenDynamic, adLockOptimistic
'                dblImporteMensual = dblImporteMensual - rstListadoSlave!Importe
'                rstProcedimientoSlave.Close
'            End If
'            'Buscamos la Retención Obra Social
'            SQL = "Select * From LIQUIDACIONSUELDOS Where CODIGOLIQUIDACION = '0524' " _
'            & "And PUESTOLABORAL = '" & PuestoLaboral & "' " _
'            & "And CODIGOCONCEPTO = '0212'"
'            If SQLNoMatch(SQL) = False Then
'                rstProcedimientoSlave.Open SQL, dbSlave, adOpenDynamic, adLockOptimistic
'                dblImporteMensual = dblImporteMensual - rstListadoSlave!Importe
'                rstProcedimientoSlave.Close
'            End If
'            'Buscamos la Retención Aporte Sindical
'            SQL = "Select * from LIQUIDACIONSUELDOS Where " _
'            & "((PUESTOLABORAL = '" & PuestoLaboral & "' " _
'            & "And CODIGOLIQUIDACION = '0524' And CODIGOCONCEPTO = '0219') Or " _
'            & "(PUESTOLABORAL = '" & PuestoLaboral & "' " _
'            & "And CODIGOLIQUIDACION = '0524' And CODIGOCONCEPTO = '0227'))"
'            If SQLNoMatch(SQL) = False Then
'                rstProcedimientoSlave.Open SQL, dbSlave, adOpenDynamic, adLockOptimistic
'                dblImporteMensual = dblImporteMensual - rstListadoSlave!Importe
'                rstProcedimientoSlave.Close
'            End If
'            'Buscamos el Pluriempleo
'            SQL = "Select * from LIQUIDACIONGANANCIAS4TACATEGORIA Where " _
'            & "(PUESTOLABORAL = '" & PuestoLaboral & "' " _
'            & "And CODIGOLIQUIDACION = '0524')"
'            If SQLNoMatch(SQL) = False Then
'                rstProcedimientoSlave.Open SQL, dbSlave, adOpenDynamic, adLockOptimistic
'                dblImporteMensual = dblImporteMensual + rstListadoSlave!Pluriempleo
'                rstProcedimientoSlave.Close
'            End If
'            dblImporteCalculado = dblImporteCalculado + dblImporteMensual
'        End If
'    End If
'
'End Sub

Public Function ImportTextFile(cn As Object, _
  ByVal tblName As String, FileFullPath As String, _
  Optional FieldDelimiter As String = ",", _
  Optional RecordDelimiter As String = vbCrLf) As Boolean
 
'PURPOSE: Imports a delimited text file into a database

'PARAMTERS: cn -- an open ado connection
'          : tblName -- import destination table name
'          : FileFullPath -- Full Path of File to import form
'          : FieldDelimiter -- (Optional) String character(s) in
'                              file separating field values
'                              within a record; defaults
'                              to ","
'          : RecordDelimiter -- (Optional) String character(s)
'                                separating records within text
'                                file; defaults to vbcrlf

'RETURNS: True if successful, false otherwise
'EXAMPLE:
'dim cn as new adodb.connection
'cn.connectionstring = _
'   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\db1.mdb"
'cn.open
'ImportTextFile cn, "MyTable", "C:\myCSVFile.csv"

'REQUIRES: VB6

Dim cmd As New ADODB.Command
Dim rs As New ADODB.Recordset
Dim sFileContents As String
Dim iFileNum As Integer
Dim sTableSplit() As String
Dim sRecordSplit() As String
Dim lCtr As Integer
Dim iCtr As Integer
Dim iFieldCtr As Integer
Dim lRecordCount As Long
Dim iFieldsToImport As Integer


'These variables prevent
'having to requery a recordset
'for each record
Dim asFieldNames() As String
Dim abFieldIsString() As Boolean
Dim iFieldCount As Integer
Dim sSQL As String
Dim bQuote As Boolean


On Error GoTo errHandler
If Not TypeOf cn Is ADODB.Connection Then Exit Function
If Dir(FileFullPath) = "" Then Exit Function

If cn.State = 0 Then cn.Open
Set cmd.ActiveConnection = cn
cmd.CommandText = tblName
cmd.CommandType = adCmdTable
Set rs = cmd.Execute
iFieldCount = rs.Fields.Count
rs.Close



ReDim asFieldNames(iFieldCount - 1) As String
ReDim abFieldIsString(iFieldCount - 1) As Boolean

For iCtr = 0 To iFieldCount - 1
    asFieldNames(iCtr) = "[" & rs.Fields(iCtr).Name & "]"
    abFieldIsString(iCtr) = FieldIsString(rs.Fields(iCtr))
Next
    

iFileNum = FreeFile
Open FileFullPath For Input As #iFileNum
sFileContents = Input(LOF(iFileNum), #iFileNum)
Close #iFileNum
'split file contents into rows
sTableSplit = Split(sFileContents, RecordDelimiter)
lRecordCount = UBound(sTableSplit)
'make it "all or nothing: whole text
'file or none of it
cn.BeginTrans

For lCtr = 0 To lRecordCount - 1
        'split record into field values
    
    sRecordSplit = Split(sTableSplit(lCtr), FieldDelimiter)
    iFieldsToImport = IIf(UBound(sRecordSplit) + 1 < _
        iFieldCount, UBound(sRecordSplit) + 1, iFieldCount)
 
   'construct sql
    sSQL = "INSERT INTO " & tblName & " ("
    
    For iCtr = 0 To iFieldsToImport - 1
        bQuote = abFieldIsString(iCtr)
        sSQL = sSQL & asFieldNames(iCtr)
        If iCtr < iFieldsToImport - 1 Then sSQL = sSQL & ","
    Next iCtr
    
    sSQL = sSQL & ") VALUES ("
    
    For iCtr = 0 To iFieldsToImport - 1
        If abFieldIsString(iCtr) Then
             sSQL = sSQL & prepStringForSQL(sRecordSplit(iCtr))
        Else
            sSQL = sSQL & sRecordSplit(iCtr)
        End If
        
        If iCtr < iFieldsToImport - 1 Then sSQL = sSQL & ","
    Next iCtr
    
    sSQL = sSQL & ")"
    cn.Execute sSQL

Next lCtr

cn.CommitTrans
rs.Close
Close #iFileNum
Set rs = Nothing
Set cmd = Nothing

ImportTextFile = True
Exit Function

errHandler:
On Error Resume Next
If cn.State <> 0 Then cn.RollbackTrans
If iFileNum > 0 Then Close #iFileNum
If rs.State <> 0 Then rs.Close
Set rs = Nothing
Set cmd = Nothing


End Function

Private Function FieldIsString(FieldObject As ADODB.Field) _
   As Boolean
    
     Select Case FieldObject.Type
         Case adBSTR, adChar, adVarChar, adWChar, adVarWChar, _
               adLongVarChar, adLongVarWChar
               FieldIsString = True
            Case Else
               FieldIsString = False
        End Select
        
End Function

Private Function prepStringForSQL(ByVal sValue As String) _
   As String

Dim sAns As String
sAns = Replace(sValue, Chr(39), "''")
sAns = "'" & sAns & "'"
prepStringForSQL = sAns

End Function

'Public Sub ImportTextToAccessADO(dbFullPath As String, tableName As String, textFullPath As String, Fields() As String)
'=====================================================================================================
'Does the physical work of inputting the text file into the database using SQL via ADO.
'   Programmatically creates SQL string
'=====================================================================================================
'On Error GoTo errHandler
'
'    Dim cnn As New ADODB.Connection
'    Dim sqlString As String
'    Dim textFilePath As String, textFname As String
'
'    Dim i As Integer 'counter variable
'    Dim intFields As Integer
'
'    textFilePath = textFullPath
'    textFname = getFileName(textFilePath, True)
'
'    cnn.Open _
'    "Provider=Microsoft.Jet.OLEDB.4.0;" & _
'    "Data Source=" & dbFullPath & ";" & _
'    "Jet OLEDB:Engine Type=4;"
'
'    sqlString = "INSERT INTO [" & tableName & "] (" & vbCrLf & vbTab
'
'    intFields = UBound(Fields())
'    For i = 0 To intFields Step 1
'        sqlString = sqlString & "[" & Fields(i) & "]"
'        If Not i = intFields Then
'            sqlString = sqlString & ", " & vbCrLf & vbTab
'        End If
'
'    Next i
'    sqlString = sqlString & vbCrLf & ")" & vbCrLf & " SELECT "
'    For i = 0 To intFields Step 1
'        sqlString = sqlString & "[" & Fields(i) & "]"
'        If Not i = intFields Then
'            sqlString = sqlString & ", " & vbCrLf & vbTab
'        End If
'    Next i
'    sqlString = sqlString & vbCrLf & "FROM [Text;DATABASE=" & textFilePath & "].[" & textFname & "];"
'
''Debug.Print sqlString
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
'    cnn.Execute sqlString
'
'    MsgBox "Operation Completed Successfully!", vbOKOnly, App.Title
'
'    Exit Sub
'
'errHandler:
'Dim prompt As String
'
'prompt = "Error: " & Err.Number & vbCrLf & vbTab & Err.Description & _
'            vbCrLf & vbCrLf & "Please try again.  Data file probably has invalid characters " & _
'            "in the field names.  (ie. '" & Chr(186) & "')."
'
'    MsgBox prompt, vbCritical, App.Title & " - Error"
'
'    BackToMainForm frmImportToDB
'
'End Sub

Sub ImportTextToAccessADO2()

Dim cnn As New ADODB.Connection
Dim sqlString As String

cnn.Open _
"Provider=Microsoft.Jet.OLEDB.4.0;" & _
"Data Source=c:\My Documents\DB1.mdb;" & _
"Jet OLEDB:Engine Type=4;"

sqlString = "SELECT * INTO [tblPeople] FROM [Text;DATABASE=C:\My Documents\TextFiles].[People.txt]"

cnn.Execute sqlString

End Sub

'This function works like the Line Input statement'
Public Sub QRLineInput( _
    ByRef strFileData As String, _
    ByRef lngFilePosition As Long, _
    ByRef strOutputString, _
    ByRef blnEOF As Boolean _
    )
    On Error GoTo LastLine
    strOutputString = Mid$(strFileData, lngFilePosition, _
        InStr(lngFilePosition, strFileData, vbNewLine) - lngFilePosition)
    lngFilePosition = InStr(lngFilePosition, strFileData, vbNewLine) + 2
    Exit Sub
LastLine:
    blnEOF = True
End Sub

Sub Test()
    Dim strFilePathName As String: strFilePathName = "C:\Fld\File.txt"
    Dim strFile As String
    Dim lngPos As Long
    Dim blnEOF As Boolean
    Dim strFileLine As String

    strFile = QuickRead(strFilePathName) & vbNewLine
    lngPos = 1

    Do Until blnEOF
        Call QRLineInput(strFile, lngPos, strFileLine, blnEOF)
    Loop
End Sub

Sub TextFile_PullData()
'PURPOSE: Send All Data From Text File To A String Variable
'SOURCE: www.TheSpreadsheetGuru.com

Dim TextFile As Integer
Dim FilePath As String
Dim FileContent As String

'File Path of Text File
  FilePath = "C:\Users\chris\Desktop\MyFile.txt"

'Determine the next file number available for use by the FileOpen function
  TextFile = FreeFile

'Open the text file
  Open FilePath For Input As TextFile

'Store file content inside a variable
  FileContent = Input(LOF(TextFile), TextFile)

'Report Out Text File Contents
  MsgBox FileContent

'Close Text File
  Close TextFile

End Sub

Sub DelimitedTextFileToArray()
'PURPOSE: Load an Array variable with data from a delimited text file
'SOURCE: www.TheSpreadsheetGuru.com

Dim Delimiter As String
Dim TextFile As Integer
Dim FilePath As String
Dim FileContent As String
Dim LineArray() As String
Dim DataArray() As String
Dim TempArray() As String
Dim rw As Long, col As Long

'Inputs
  Delimiter = ";"
  FilePath = "C:\Users\chris\Desktop\MyFile.txt"
  rw = 0
  
'Open the text file in a Read State
  TextFile = FreeFile
  Open FilePath For Input As TextFile
  
'Store file content inside a variable
  FileContent = Input(LOF(TextFile), TextFile)

'Close Text File
  Close TextFile
  
'Separate Out lines of data
  LineArray() = Split(FileContent, vbCrLf)

'Read Data into an Array Variable
  For x = LBound(LineArray) To UBound(LineArray)
    If Len(Trim(LineArray(x))) <> 0 Then
      'Split up line of text by delimiter
        TempArray = Split(LineArray(x), Delimiter)
      
      'Determine how many columns are needed
        col = UBound(TempArray)
      
      'Re-Adjust Array boundaries
        ReDim Preserve DataArray(col, rw)
      
      'Load line of data into Array variable
        For y = LBound(TempArray) To UBound(TempArray)
          DataArray(y, rw) = TempArray(y)
        Next y
    End If
    
    'Next line
      rw = rw + 1
    
  Next x

End Sub


