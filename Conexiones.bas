Attribute VB_Name = "Conexiones"
'Public dbSlaveDAO As Database
Public dbSlave As ADODB.Connection
Public rstListadoSlave As ADODB.Recordset
Public rstBuscarSlave As ADODB.Recordset
Public rstRegistroSlave As ADODB.Recordset
Public rstProcedimientoSlave As ADODB.Recordset

Public Sub Conectar()
    
    Set dbSlave = New ADODB.Connection
    dbSlave.CursorLocation = adUseClient
    dbSlave.Mode = adModeUnknown
    dbSlave.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\Slave.mdb"
    dbSlave.Open

End Sub

Public Sub ConectarDAO()
    
    'Set dbSlaveDAO = OpenDatabase(App.Path & "\Slave.mdb")

End Sub

Public Sub Desconectar()
    
    dbSlave.Close
    Set dbSlave = Nothing

End Sub

Public Sub DesconectarDAO()
    
    'dbSlaveDAO.Close
    'Set dbSlaveDAO = Nothing

End Sub

Public Sub VaciarTodasLasVariables()

    'strComprobante = ""
    'strCargaEstructura = ""
    'strEditandoEstructura = ""
    'bolEditandoRetencionGanancias = False
    'ValidarLiquidacionGanancias = False
    'ValidarAgente = False
    'ValidarConcepto = False
    'ValidarEscalaGanancias = False
    'ValidarLimitesDeducciones = False
    'ValidarParentesco = False
    'ValidarDeduccionesGenerales = False
    'ValidarCodigoLiquidacion = False
    'ValidarFamiliar = False
    'ValidarLiquidacionGanancias = False
    'ValidarGenerarF649 = False

End Sub
