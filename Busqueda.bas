Attribute VB_Name = "Busqueda"
Public Function SeekNoMatch(ValorBuscado As String, NombreTabla As String) As Boolean
    
    SeekNoMatch = False
    Set rstBuscarSlave = New ADODB.Recordset
    rstBuscarSlave.Open NombreTabla, dbSlave, adOpenDynamic, adLockReadOnly
    With rstBuscarSlave
        .Index = rstBuscarSlave.Fields.Item(0).Name
        .Seek "=", ValorBuscado
    End With
    If rstBuscarSlave.EOF = True Then
        SeekNoMatch = True
    End If
    rstBuscarSlave.Close
    Set rstBuscarSlave = Nothing
    
End Function

Public Function FindNoMatch(ValorBuscado As String, SQL As String, CampoBusqueda As String) As Boolean
    
    FindNoMatch = False
    Set rstBuscarSlave = New ADODB.Recordset
    rstBuscarSlave.Open SQL, dbSlave, adOpenDynamic, adLockOptimistic
    rstBuscarSlave.Find CampoBusqueda & "='" & ValorBuscado & "'"
    If rstBuscarSlave.EOF = True Then
        FindNoMatch = True
    End If
    rstBuscarSlave.Close
    Set rstBuscarSlave = Nothing
    
End Function

Public Function SQLNoMatch(SQL As String) As Boolean
    
    SQLNoMatch = False
    Set rstBuscarSlave = New ADODB.Recordset
    rstBuscarSlave.Open SQL, dbSlave, adOpenForwardOnly, adLockOptimistic
    If (rstBuscarSlave.EOF And rstBuscarSlave.BOF) Or IsNull(rstBuscarSlave.Fields(0)) = True Then
        SQLNoMatch = True
    End If
    rstBuscarSlave.Close
    Set rstBuscarSlave = Nothing
    
End Function
