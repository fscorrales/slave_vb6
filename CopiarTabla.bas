Attribute VB_Name = "CopiarTabla"
Public Sub CopiarTablaOrigenExterno()

    Dim SQL As String
    'Dim qdf As QueryDef
    
'    SQL = "Select * From LIQUIDACIONSUELDOS Where CODIGOLIQUIDACION = " & "'" & CodigoOrigen & "'"
'    If SQLNoMatch(SQL) = True Then
'        MsgBox "La liquidación que desea copiar debe contener registros", vbCritical + vbOKOnly, "LIQUIDACIÓN DE ORIGEN VACÍA"
'        CopiarLiquidacionSisper.cmbCodigoLiquidacionOrigen.SetFocus
'        Exit Sub
'    End If
'    SQL = "Select * From LIQUIDACIONSUELDOS Where CODIGOLIQUIDACION = " & "'" & CodigoDestino & "'"
'    If SQLNoMatch(SQL) = False Then
'        MsgBox "Proceda a borrar los registros de la liquidación destino antes de incorporar datos a la misma", vbCritical + vbOKOnly, "LIQUIDACIÓN DESTINO LLENA"
'        CopiarLiquidacionSisper.cmbCodigoLiquidacionDestino.SetFocus
'        Exit Sub
'    End If

    ' Select all records in the Employees table
    ' and copy them into a new table, Emp Backup.
    SQL = "SELECT * INTO PRECARIZADOS" _
    & " FROM AGENTE IN 'C:\Users\Toshiba\Dropbox\Programas Creados\SAIHP\SAIHPUltimo.mdb'"
    
'    'SQL INSERT INTO SELECT Statement
'    SQL = "Insert Into LIQUIDACIONSUELDOS (CodigoLiquidacion, PuestoLaboral, CodigoConcepto, Importe) " & _
'    "Select '" & CodigoDestino & "', PuestoLaboral, CodigoConcepto, Importe " & _
'    "From LIQUIDACIONSUELDOS Where CodigoLiquidacion = " & "'" & CodigoOrigen & "'"
'    'Debug.Print SQL
    
    dbSlave.BeginTrans
    dbSlave.Execute SQL
    dbSlave.CommitTrans
    
'    SQL = "DELETE FROM LIQUIDACIONHONORARIOS"
'
'    dbSlave.BeginTrans
'    dbSlave.Execute SQL
'    dbSlave.CommitTrans
    
End Sub

Public Sub SetNullToZero()

    Dim SQL As String

    SQL = "UPDATE Deducciones4taCategoria SET Alquileres = 0 WHERE Alquileres IS NULL"

    dbSlave.BeginTrans
    dbSlave.Execute SQL
    dbSlave.CommitTrans
    
End Sub

