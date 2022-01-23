Attribute VB_Name = "Exportar"
Public Function ExportarResumenAnualGcias(strOutputPath As String) As Boolean
    On Error GoTo Error_Handler
  
    Dim o_Excel            As Object
    Dim o_Libro            As Object
    Dim o_Hoja             As Object
    Dim Fila               As Long
    Dim Columna            As Long
    Dim strPuestoLaboral   As String
       
    ' -- Crea el objeto Excel, el objeto workBook y el objeto sheet
    Set o_Excel = CreateObject("Excel.Application")
    Set o_Libro = o_Excel.Workbooks.Add
    
    With ResumenAnualGanancias
'        Debug.Print .cmbAgente.ListCount
'        .cmbAgente.ListIndex = 0
'        Debug.Print .cmbAgente.Text
        For i = 0 To (.cmbAgente.ListCount - 1)
            .cmbAgente.ListIndex = i
            strPuestoLaboral = "Select PUESTOLABORAL From AGENTES" _
            & " Where NOMBRECOMPLETO = '" & .cmbAgente.Text & "'"
            Set rstBuscarSlave = New ADODB.Recordset
            rstBuscarSlave.Open strPuestoLaboral, dbSlave, adOpenDynamic, adLockOptimistic
            strPuestoLaboral = rstBuscarSlave!PuestoLaboral
            rstBuscarSlave.Close
            Set rstBuscarSlave = Nothing
            Call CargardgLiquidacionMensualGanancias(strPuestoLaboral, .cmbPeriodo.Text)
            Set o_Hoja = o_Libro.Worksheets.Add
            o_Hoja.Name = .cmbAgente.Text
            ' -- Bucle para Exportar los datos
            For Fila = 0 To .dgLiquidacionMensualGanancias.Rows - 1
                For Columna = 0 To .dgLiquidacionMensualGanancias.Cols - 1
                    o_Hoja.Cells(Fila + 1, Columna + 1).Value = .dgLiquidacionMensualGanancias.TextMatrix(Fila, Columna)
                Next
            Next
            If Not o_Hoja Is Nothing Then Set o_Hoja = Nothing
        Next i
    End With
    o_Libro.Close True, strOutputPath
    ' -- Cerrar Excel
    o_Excel.Quit
    ' -- Terminar instancias
    Call ReleaseObjects(o_Excel, o_Libro, o_Hoja)
    ExportarResumenAnualGcias = True
Exit Function
  
' -- Controlador de Errores
Error_Handler:
    ' -- Cierra la hoja y el la aplicación Excel
    If Not o_Libro Is Nothing Then: o_Libro.Close False
    If Not o_Excel Is Nothing Then: o_Excel.Quit
    Call ReleaseObjects(o_Excel, o_Libro, o_Hoja)
    If Err.Number <> 1004 Then MsgBox Err.Description, vbCritical
End Function


Public Function Exportar_Excel(strOutputPath As String, FlexGrid As Object, Optional sNameSheet As String = "SinNombre") As Boolean
    On Error GoTo Error_Handler
  
    Dim o_Excel     As Object
    Dim o_Libro     As Object
    Dim o_Hoja      As Object
    Dim Fila        As Long
    Dim Columna     As Long
       
    ' -- Crea el objeto Excel, el objeto workBook y el objeto sheet
    Set o_Excel = CreateObject("Excel.Application")
    Set o_Libro = o_Excel.Workbooks.Add
    Set o_Hoja = o_Libro.Worksheets.Add
    o_Hoja.Name = sNameSheet
    'Set o_Hoja = o_Libro.Worksheets(sNameSheet).Add
       
    ' -- Bucle para Exportar los datos
    With FlexGrid
        For Fila = 0 To .Rows - 1
            For Columna = 0 To .Cols - 1
                o_Hoja.Cells(Fila + 1, Columna + 1).Value = .TextMatrix(Fila, Columna)
            Next
        Next
    End With
    o_Libro.Close True, strOutputPath
    ' -- Cerrar Excel
    o_Excel.Quit
    ' -- Terminar instancias
    Call ReleaseObjects(o_Excel, o_Libro, o_Hoja)
    Exportar_Excel = True
Exit Function
  
' -- Controlador de Errores
Error_Handler:
    ' -- Cierra la hoja y el la aplicación Excel
    If Not o_Libro Is Nothing Then: o_Libro.Close False
    If Not o_Excel Is Nothing Then: o_Excel.Quit
    Call ReleaseObjects(o_Excel, o_Libro, o_Hoja)
    If Err.Number <> 1004 Then MsgBox Err.Description, vbCritical
End Function
' -------------------------------------------------------------------
' \\ -- Eliminar objetos para liberar recursos
' -------------------------------------------------------------------
Private Sub ReleaseObjects(o_Excel As Object, o_Libro As Object, o_Hoja As Object)
    If Not o_Excel Is Nothing Then Set o_Excel = Nothing
    If Not o_Libro Is Nothing Then Set o_Libro = Nothing
    If Not o_Hoja Is Nothing Then Set o_Hoja = Nothing
End Sub

' \\ -- función para psar los datos hacia una hoja de un libro exisitente
' -------------------------------------------------------------------------------------------
Public Function Exportar_Excel2(sBookFileName As String, FlexGrid As Object, Optional sNameSheet As String = vbNullString) As Boolean
  
    On Error GoTo Error_Handler
  
    Dim o_Excel     As Object
    Dim o_Libro     As Object
    Dim o_Hoja      As Object
    Dim Fila        As Long
    Dim Columna     As Long
  
    ' -- Error en la ruta del libro
    If sBookFileName = vbNullString Or Len(Dir(sBookFileName)) = 0 Then
           
        MsgBox " Falta el Path del archivo de Excel o no se ha encontrado el libro en la ruta especificada ", vbCritical
        Exit Function
    End If
       
    ' -- Crea el objeto Excel, el objeto workBook y el objeto sheet
    Set o_Excel = CreateObject("Excel.Application")
    Set o_Libro = o_Excel.Workbooks.Open(sBookFileName)
    ' -- Comprobar si se abre la hoja por defecto, o la indicada en el parámetro de la función
    If Len(sNameSheet) = 0 Then
        Set o_Hoja = o_Libro.Worksheets(1)
    Else
        Set o_Hoja = o_Libro.Worksheets(sNameSheet)
    End If
    ' -- Bucle para Exportar los datos
    With FlexGrid
        For Fila = 1 To .Rows - 1
            For Columna = 0 To .Cols - 1
                o_Hoja.Cells(Fila, Columna + 1).Value = .TextMatrix(Fila, Columna)
            Next
        Next
    End With
    ' -- Cerrar libro y guardar los datos
    o_Libro.Close True
    ' -- Cerrar Excel
    o_Excel.Quit
    ' -- Terminar instancias
    Call ReleaseObjects(o_Excel, o_Libro, o_Hoja)
    Exportar_Excel2 = True
Exit Function
  
' -- Controlador de Errores
Error_Handler:
    ' -- Cierra la hoja y el la aplicación Excel
    If Not o_Libro Is Nothing Then: o_Libro.Close False
    If Not o_Excel Is Nothing Then: o_Excel.Quit
    Call ReleaseObjects(o_Excel, o_Libro, o_Hoja)
    MsgBox Err.Description, vbCritical
End Function


Public Function ExportarEPAMConsolidado(strOutputPath As String, FlexGridGeneral As Object, FlexGridDiscriminado As Object) As Boolean
    On Error GoTo Error_Handler
  
    Dim o_Excel             As Object
    Dim o_Libro             As Object
    Dim o_Hoja              As Object
    Dim FilaGeneral         As Long
    Dim FilaDiscriminada    As Long
    Dim ColumnaDiscriminada As Long
    Dim x                   As Long

       
    ' -- Crea el objeto Excel, el objeto workBook y el objeto sheet
    Set o_Excel = CreateObject("Excel.Application")
    Set o_Libro = o_Excel.Workbooks.Add
    Set o_Hoja = o_Libro.Worksheets.Add
       
    ' -- Bucle para Exportar los datos
    o_Hoja.Cells(1, 1).Value = "PRESUPUESTO DE OBRAS EPAM AL " & Date
    o_Hoja.Range("A1:H1").MergeCells = True
    o_Hoja.Range("A1:H1").Font.Size = 14
    o_Hoja.Range("A1:H1").Font.Bold = True
    o_Hoja.Range("A1:H1").HorizontalAlignment = xlVAlignCenter
    o_Hoja.Cells(3, 1).Value = "Descripción"
    o_Hoja.Cells(3, 3).Value = "Imputación Presupuestaria"
    o_Hoja.Cells(3, 4).Value = "Presupuesto"
    o_Hoja.Cells(3, 5).Value = "Pagado"
    o_Hoja.Cells(3, 6).Value = "Imputado"
    o_Hoja.Cells(3, 7).Value = "Decisión a Tomar"
    o_Hoja.Cells(3, 8).Value = "Límite"
    o_Hoja.Range("C3:H3").HorizontalAlignment = xlVAlignCenter
    o_Hoja.Range("A3:H3").Font.Bold = True
    o_Hoja.Range("A3:H3").Interior.ColorIndex = 48
    o_Hoja.Columns(1).ColumnWidth = 35
    o_Hoja.Columns(2).ColumnWidth = 12
    o_Hoja.Columns(3).ColumnWidth = 25
    o_Hoja.Columns(4).ColumnWidth = 12
    o_Hoja.Columns(5).ColumnWidth = 11
    o_Hoja.Columns(6).ColumnWidth = 11
    o_Hoja.Columns(7).ColumnWidth = 22
    o_Hoja.Columns(8).ColumnWidth = 11
    
    With o_Hoja.Range("A3:H3").Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .Weight = xlMedium
        .ColorIndex = xlAutomatic
    End With
    With o_Hoja.Range("A3:H3").Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Weight = xlMedium
        .ColorIndex = xlAutomatic
    End With
    With o_Hoja.Range("A3:H3").Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlMedium
        .ColorIndex = xlAutomatic
    End With
    With o_Hoja.Range("A3:H3").Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .Weight = xlMedium
        .ColorIndex = xlAutomatic
    End With
    
    With FlexGridGeneral
        x = 3
        For FilaGeneral = 1 To .Rows - 1
            EPAMConsolidado.ConfigurarResolucionesEPAM
            EPAMConsolidado.CargarResolucionesEPAM (FlexGridGeneral.TextMatrix(FilaGeneral, 0))
            x = x + 1
            o_Hoja.Cells(x, 1).Value = .TextMatrix(FilaGeneral, 0)
            o_Hoja.Range("A" & x & ":B" & x).MergeCells = True
            o_Hoja.Range("A" & x & ":B" & x).HorizontalAlignment = xlVAlignCenter
            o_Hoja.Cells(x, 3).Value = .TextMatrix(FilaGeneral, 1)
            o_Hoja.Cells(x, 5).Value = .TextMatrix(FilaGeneral, 2)
            o_Hoja.Cells(x, 6).Value = .TextMatrix(FilaGeneral, 3)
            o_Hoja.Cells(x, 7).Value = .TextMatrix(FilaGeneral, 4)
            With FlexGridDiscriminado
                For FilaDiscriminada = 1 To .Rows - 1
                    x = x + 1
                    For ColumnaDiscriminada = 0 To .Cols - 1
                        o_Hoja.Cells(x, ColumnaDiscriminada + 1).Value = .TextMatrix(FilaDiscriminada, ColumnaDiscriminada)
                    Next
                Next
            End With
        Next
    End With

'Configurando Bordes de la Grilla
    o_Hoja.Range("A4:H" & x).Select
    o_Excel.Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    o_Excel.Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With o_Excel.Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .Weight = xlMedium
        .ColorIndex = xlAutomatic
    End With
    With o_Excel.Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Weight = xlMedium
        .ColorIndex = xlAutomatic
    End With
    With o_Excel.Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlMedium
        .ColorIndex = xlAutomatic
    End With
    With o_Excel.Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .Weight = xlMedium
        .ColorIndex = xlAutomatic
    End With
    With o_Excel.Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With o_Excel.Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    
'Configurando Para Impresión
    o_Hoja.Rows("1:3").Select
    With o_Libro.ActiveSheet.PageSetup
        .PrintTitleRows = "$1:$3"
        .PrintTitleColumns = ""
    End With
    o_Libro.ActiveSheet.PageSetup.PrintArea = ""
    With o_Libro.ActiveSheet.PageSetup
        .LeftHeader = ""
        .CenterHeader = ""
        .RightHeader = ""
        .LeftFooter = ""
        .CenterFooter = ""
        .RightFooter = ""
        .LeftMargin = Application.InchesToPoints(0.393700787401575)
        .RightMargin = Application.InchesToPoints(0.393700787401575)
        .TopMargin = Application.InchesToPoints(0.984251968503937)
        .BottomMargin = Application.InchesToPoints(0.984251968503937)
        .HeaderMargin = Application.InchesToPoints(0)
        .FooterMargin = Application.InchesToPoints(0)
        .PrintHeadings = False
        .PrintGridlines = False
        .PrintComments = xlPrintNoComments
        .PrintQuality = 600
        .CenterHorizontally = True
        .CenterVertically = False
        .Orientation = xlLandscape
        .Draft = False
        .PaperSize = xlPaperLetter
        .FirstPageNumber = xlAutomatic
        .Order = xlDownThenOver
        .BlackAndWhite = False
        .Zoom = 90
        .PrintErrors = xlPrintErrorsDisplayed
    End With
'Inmovilizar Paneles
    o_Hoja.Rows("4:4").Select
    o_Excel.ActiveWindow.FreezePanes = True

    o_Libro.Close True, strOutputPath
    ' -- Cerrar Excel
    o_Excel.Quit
    ' -- Terminar instancias
    Call ReleaseObjects(o_Excel, o_Libro, o_Hoja)
    ExportarEPAMConsolidado = True
Exit Function
  
' -- Controlador de Errores
Error_Handler:
    ' -- Cierra la hoja y el la aplicación Excel
    If Not o_Libro Is Nothing Then: o_Libro.Close False
    If Not o_Excel Is Nothing Then: o_Excel.Quit
    Call ReleaseObjects(o_Excel, o_Libro, o_Hoja)
    If Err.Number <> 1004 Then MsgBox Err.Description, vbCritical
End Function

Public Function ExportarLiquidacionGananciasSISPERViejo(strOutputPath As String, CodigoLiquidacion As String, FlexGridAgentesRetenidos As Object) As Boolean
    On Error GoTo Error_Handler
  
    Dim o_Excel             As Object
    Dim o_Libro             As Object
    Dim o_Hoja              As Object
    Dim FilaFlexGrid        As Integer
    Dim FilaExcel           As Integer
       
    ' -- Crea el objeto Excel, el objeto workBook y el objeto sheet
    Set o_Excel = CreateObject("Excel.Application")
    Set o_Libro = o_Excel.Workbooks.Add
    Set o_Hoja = o_Libro.Worksheets(1)
       
    ' -- Bucle para Exportar los datos
    With FlexGridAgentesRetenidos
        FilaExcel = 0
        For FilaFlexGrid = 1 To .Rows - 1
            If IsNumeric(.TextMatrix(FilaFlexGrid, 2)) = True And .TextMatrix(FilaFlexGrid, 2) <> 0 Then
                FilaExcel = FilaExcel + 1
                o_Hoja.Cells(FilaExcel, 1).Value = 26
                o_Hoja.Cells(FilaExcel, 2).Value = De_Txt_a_Num_01(CodigoLiquidacion, 0)
                o_Hoja.Cells(FilaExcel, 3).Value = .TextMatrix(FilaFlexGrid, 0)
                o_Hoja.Cells(FilaExcel, 4).Value = 8276
                o_Hoja.Cells(FilaExcel, 5).Value = .TextMatrix(FilaFlexGrid, 2)
            End If
        Next
    End With

    o_Libro.Close True, strOutputPath
    ' -- Cerrar Excel
    o_Excel.Quit
    ' -- Terminar instancias
    Call ReleaseObjects(o_Excel, o_Libro, o_Hoja)
    ExportarLiquidacionGananciasSISPERViejo = True
Exit Function
  
' -- Controlador de Errores
Error_Handler:
    ' -- Cierra la hoja y el la aplicación Excel
    If Not o_Libro Is Nothing Then: o_Libro.Close False
    If Not o_Excel Is Nothing Then: o_Excel.Quit
    Call ReleaseObjects(o_Excel, o_Libro, o_Hoja)
    If Err.Number <> 1004 Then MsgBox Err.Description, vbCritical
End Function

Public Function ExportarLiquidacionGananciasSISPER(strOutputPath As String, CodigoLiquidacion As String) As Boolean
    On Error GoTo Error_Handler
  
    Dim o_Excel             As Object
    Dim o_Libro             As Object
    Dim o_Hoja              As Object
    Dim FilaFlexGrid        As Integer
    Dim FilaExcel           As Integer
    Dim SQL                 As String
       
    ' -- Crea el objeto Excel, el objeto workBook y el objeto sheet
    Set o_Excel = CreateObject("Excel.Application")
    Set o_Libro = o_Excel.Workbooks.Add
    Set o_Hoja = o_Libro.Worksheets(1)
       
    ' -- Creamos el recordset
    SQL = "SELECT PuestoLaboral, Retencion " _
        & "FROM LiquidacionGanancias4taCategoria " _
        & "Where CodigoLiquidacion = '" & CodigoLiquidacion & "' " _
        & "And Retencion > 0"
    Set rstListadoSlave = New ADODB.Recordset
    rstListadoSlave.Open SQL, dbSlave, adOpenDynamic, adLockOptimistic
       
    ' -- Bucle para Exportar los datos
    If rstListadoSlave.BOF = False Then
        rstListadoSlave.MoveFirst
        FilaExcel = 0
        While rstListadoSlave.EOF = False
            FilaExcel = FilaExcel + 1
            o_Hoja.Cells(FilaExcel, 1).Value = 26
            o_Hoja.Cells(FilaExcel, 2).Value = De_Txt_a_Num_01(CodigoLiquidacion, 0)
            o_Hoja.Cells(FilaExcel, 3).Value = rstListadoSlave!PuestoLaboral
            o_Hoja.Cells(FilaExcel, 4).Value = 8276
            o_Hoja.Cells(FilaExcel, 5).Value = rstListadoSlave!Retencion
            rstListadoSlave.MoveNext
        Wend
    End If

    o_Libro.Close True, strOutputPath
    ' -- Cerrar Excel
    o_Excel.Quit
    ' -- Terminar instancias
    Call ReleaseObjects(o_Excel, o_Libro, o_Hoja)
    rstListadoSlave.Close
    Set rstListadoSlave = Nothing
    ExportarLiquidacionGananciasSISPER = True
Exit Function
  
' -- Controlador de Errores
Error_Handler:
    ' -- Cierra la hoja y el la aplicación Excel
    If Not o_Libro Is Nothing Then: o_Libro.Close False
    If Not o_Excel Is Nothing Then: o_Excel.Quit
    Call ReleaseObjects(o_Excel, o_Libro, o_Hoja)
    If Err.Number <> 1004 Then MsgBox Err.Description, vbCritical
End Function


Public Function ImprimirComprobanteSIIF(strOutputPath As String, ComprobanteSIIF As String, _
FlexGridImputacion As Object, FlexGridRetencion As Object) As Boolean
    On Error GoTo Error_Handler
  
    Dim o_Excel             As Object
    Dim o_Libro             As Object
    Dim o_Hoja              As Object
    Dim FilaFlexGrid        As Integer
    Dim FilaExcel           As Integer
       
    ' -- Crea el objeto Excel, el objeto workBook y el objeto sheet
    Set o_Excel = CreateObject("Excel.Application")
    Set o_Libro = o_Excel.Workbooks.Add
    Set o_Hoja = o_Libro.Worksheets(1)
       
    ' -- Titulo
    FilaExcel = 1
    o_Hoja.Cells(FilaExcel, 1).Value = "Comprobante SIIF Nro: " & ComprobanteSIIF
       
    ' -- Bucle para Exportar los datos
    With FlexGridImputacion
        FilaExcel = FilaExcel + 2
        o_Hoja.Cells(FilaExcel, 1).Value = "Prog."
        o_Hoja.Cells(FilaExcel, 2).Value = "Proy."
        o_Hoja.Cells(FilaExcel, 3).Value = "Act."
        o_Hoja.Cells(FilaExcel, 4).Value = "Partida"
        o_Hoja.Cells(FilaExcel, 5).Value = "Importe"
        For FilaFlexGrid = 1 To .Rows - 1
            If IsNumeric(.TextMatrix(FilaFlexGrid, 2)) = True Then
                FilaExcel = FilaExcel + 1
                o_Hoja.Cells(FilaExcel, 1).Value = Left(.TextMatrix(FilaFlexGrid, 0), 2)
                o_Hoja.Cells(FilaExcel, 2).Value = Mid(.TextMatrix(FilaFlexGrid, 0), 4, 2)
                o_Hoja.Cells(FilaExcel, 3).Value = Right(.TextMatrix(FilaFlexGrid, 0), 2)
                o_Hoja.Cells(FilaExcel, 4).Value = .TextMatrix(FilaFlexGrid, 1)
                o_Hoja.Cells(FilaExcel, 5).Value = .TextMatrix(FilaFlexGrid, 2)
            End If
        Next
    End With

    With FlexGridRetencion
        FilaExcel = FilaExcel + 2
        o_Hoja.Cells(FilaExcel, 1).Value = "Código"
        o_Hoja.Cells(FilaExcel, 2).Value = "Importe"
        For FilaFlexGrid = 1 To .Rows - 1
            If IsNumeric(.TextMatrix(FilaFlexGrid, 1)) = True Then
                FilaExcel = FilaExcel + 1
                o_Hoja.Cells(FilaExcel, 1).Value = .TextMatrix(FilaFlexGrid, 0)
                o_Hoja.Cells(FilaExcel, 2).Value = .TextMatrix(FilaFlexGrid, 1)
            End If
        Next
    End With

    o_Libro.Close True, strOutputPath
    ' -- Cerrar Excel
    o_Excel.Quit
    ' -- Terminar instancias
    Call ReleaseObjects(o_Excel, o_Libro, o_Hoja)
    ImprimirComprobanteSIIF = True
Exit Function
  
' -- Controlador de Errores
Error_Handler:
    ' -- Cierra la hoja y el la aplicación Excel
    If Not o_Libro Is Nothing Then: o_Libro.Close False
    If Not o_Excel Is Nothing Then: o_Excel.Quit
    Call ReleaseObjects(o_Excel, o_Libro, o_Hoja)
    If Err.Number <> 1004 Then MsgBox Err.Description, vbCritical
End Function

'Public Function GenerarReporteMensualGanancias(strOutputPath As String, CodigoLiquidacion As String) As Boolean
'    On Error GoTo Error_Handler
'
'    Dim o_Excel                 As Object
'    Dim o_Libro                 As Object
'    Dim o_Hoja                  As Object
'    Dim ColumnaExcel            As Integer
'    Dim FilaExcel               As Integer
'    Dim SQL                     As String
'    Dim PeriodoLiquidacion      As String
'    Dim DescripcionLiquidacion  As String
'    Dim ApellidoAgente()        As String
'    Dim SumaTotal1              As Double
'    Dim SumaTotal2              As Double
'
'    ' -- Crea el objeto Excel, el objeto workBook y el objeto sheet
'    Set o_Excel = CreateObject("Excel.Application")
'    Set o_Libro = o_Excel.Workbooks.Add
'    Set o_Hoja = o_Libro.Worksheets(1)
'
'    SQL = "Select * From CODIGOLIQUIDACIONES Where CODIGO = '" & CodigoLiquidacion & "'"
'    Set rstRegistroSlave = New ADODB.Recordset
'    rstRegistroSlave.Open SQL, dbSlave, adOpenForwardOnly, adLockOptimistic
'    DescripcionLiquidacion = rstRegistroSlave!Descripcion
'    PeriodoLiquidacion = rstRegistroSlave!PERIODO
'    rstRegistroSlave.Close
'    Set rstRegistroSlave = Nothing
'
'    ' -- Bucle para Exportar los datos
'    Set rstListadoSlave = New ADODB.Recordset
'    Set rstBuscarSlave = New ADODB.Recordset
'    SQL = "Select * From LIQUIDACIONGANANCIAS4TACATEGORIA Where CODIGOLIQUIDACION = '" & CodigoLiquidacion & "' " _
'    & "And RETENCION <> 0 Order by Retencion Desc"
'    rstListadoSlave.Open SQL, dbSlave, adOpenDynamic, adLockOptimistic
'    FilaExcel = 8
'    rstListadoSlave.MoveFirst
'    With o_Hoja
'        Do
'            ColumnaExcel = 1
'            .Cells(FilaExcel - 6, 1).Value = "RETENCIÓN DE IMPUESTO A LAS GANANCIAS 4TA CATEGORÍA"
'            .Range("A" & FilaExcel - 6 & ":O" & FilaExcel - 6).MergeCells = True
'            .Range("A" & FilaExcel - 6 & ":O" & FilaExcel - 6).Font.Size = 12
'            .Range("A" & FilaExcel - 6 & ":O" & FilaExcel - 6).Font.Bold = True
'            .Range("A" & FilaExcel - 6 & ":O" & FilaExcel - 6).HorizontalAlignment = xlVAlignCenter
'            .Cells(FilaExcel - 4, 1).Value = DescripcionLiquidacion
'            .Range("A" & FilaExcel - 4 & ":O" & FilaExcel - 4).MergeCells = True
'            .Range("A" & FilaExcel - 4 & ":O" & FilaExcel - 4).Font.Size = 11
'            .Range("A" & FilaExcel - 4 & ":O" & FilaExcel - 4).Font.Bold = True
'            .Range("A" & FilaExcel - 4 & ":O" & FilaExcel - 4).HorizontalAlignment = xlVAlignCenter
'
'            'Configurando Bordes de la Grilla
'            .Range("B" & FilaExcel & ":O" & FilaExcel + 36).Select
'            o_Excel.Selection.Borders(xlDiagonalDown).LineStyle = xlNone
'            o_Excel.Selection.Borders(xlDiagonalUp).LineStyle = xlNone
'            With o_Excel.Selection.Borders(xlEdgeLeft)
'                .LineStyle = xlContinuous
'                .Weight = xlMedium
'                .ColorIndex = xlAutomatic
'            End With
'            With o_Excel.Selection.Borders(xlEdgeTop)
'                .LineStyle = xlContinuous
'                .Weight = xlMedium
'                .ColorIndex = xlAutomatic
'            End With
'            With o_Excel.Selection.Borders(xlEdgeBottom)
'                .LineStyle = xlContinuous
'                .Weight = xlMedium
'                .ColorIndex = xlAutomatic
'            End With
'            With o_Excel.Selection.Borders(xlEdgeRight)
'                .LineStyle = xlContinuous
'                .Weight = xlMedium
'                .ColorIndex = xlAutomatic
'            End With
'            With o_Excel.Selection.Borders(xlInsideVertical)
'                .LineStyle = xlContinuous
'                .Weight = xlMedium
'                .ColorIndex = xlAutomatic
'            End With
'            With o_Excel.Selection.Borders(xlInsideHorizontal)
'                .LineStyle = xlContinuous
'                .Weight = xlThin
'                .ColorIndex = xlAutomatic
'            End With
'            .Range("A" & FilaExcel & ":A" & FilaExcel + 38).Select
'            o_Excel.Selection.Borders(xlDiagonalDown).LineStyle = xlNone
'            o_Excel.Selection.Borders(xlDiagonalUp).LineStyle = xlNone
'            With o_Excel.Selection.Borders(xlEdgeLeft)
'                .LineStyle = xlContinuous
'                .Weight = xlMedium
'                .ColorIndex = xlAutomatic
'            End With
'            With o_Excel.Selection.Borders(xlEdgeTop)
'                .LineStyle = xlContinuous
'                .Weight = xlMedium
'                .ColorIndex = xlAutomatic
'            End With
'            With o_Excel.Selection.Borders(xlEdgeBottom)
'                .LineStyle = xlContinuous
'                .Weight = xlMedium
'                .ColorIndex = xlAutomatic
'            End With
'            With o_Excel.Selection.Borders(xlEdgeRight)
'                .LineStyle = xlContinuous
'                .Weight = xlMedium
'                .ColorIndex = xlAutomatic
'            End With
'            .Range("A" & FilaExcel + 37 & ":O" & FilaExcel + 38).Select
'            o_Excel.Selection.Borders(xlDiagonalDown).LineStyle = xlNone
'            o_Excel.Selection.Borders(xlDiagonalUp).LineStyle = xlNone
'            With o_Excel.Selection.Borders(xlEdgeLeft)
'                .LineStyle = xlContinuous
'                .Weight = xlMedium
'                .ColorIndex = xlAutomatic
'            End With
'            With o_Excel.Selection.Borders(xlEdgeTop)
'                .LineStyle = xlContinuous
'                .Weight = xlMedium
'                .ColorIndex = xlAutomatic
'            End With
'            With o_Excel.Selection.Borders(xlEdgeBottom)
'                .LineStyle = xlContinuous
'                .Weight = xlMedium
'                .ColorIndex = xlAutomatic
'            End With
'            With o_Excel.Selection.Borders(xlEdgeRight)
'                .LineStyle = xlContinuous
'                .Weight = xlMedium
'                .ColorIndex = xlAutomatic
'            End With
'            With o_Excel.Selection.Borders(xlInsideVertical)
'                .LineStyle = xlContinuous
'                .Weight = xlMedium
'                .ColorIndex = xlAutomatic
'            End With
'            .Range("B" & FilaExcel - 2 & ":O" & FilaExcel - 1).Select
'            o_Excel.Selection.Borders(xlDiagonalDown).LineStyle = xlNone
'            o_Excel.Selection.Borders(xlDiagonalUp).LineStyle = xlNone
'            With o_Excel.Selection.Borders(xlEdgeLeft)
'                .LineStyle = xlContinuous
'                .Weight = xlMedium
'                .ColorIndex = xlAutomatic
'            End With
'            With o_Excel.Selection.Borders(xlEdgeTop)
'                .LineStyle = xlContinuous
'                .Weight = xlMedium
'                .ColorIndex = xlAutomatic
'            End With
'            With o_Excel.Selection.Borders(xlEdgeBottom)
'                .LineStyle = xlContinuous
'                .Weight = xlMedium
'                .ColorIndex = xlAutomatic
'            End With
'            With o_Excel.Selection.Borders(xlEdgeRight)
'                .LineStyle = xlContinuous
'                .Weight = xlMedium
'                .ColorIndex = xlAutomatic
'            End With
'            With o_Excel.Selection.Borders(xlInsideVertical)
'                .LineStyle = xlContinuous
'                .Weight = xlMedium
'                .ColorIndex = xlAutomatic
'            End With
'            .Range("A" & FilaExcel & ":O" & FilaExcel).Select
'            o_Excel.Selection.Borders(xlDiagonalDown).LineStyle = xlNone
'            o_Excel.Selection.Borders(xlDiagonalUp).LineStyle = xlNone
'            With o_Excel.Selection.Borders(xlEdgeLeft)
'                .LineStyle = xlContinuous
'                .Weight = xlMedium
'                .ColorIndex = xlAutomatic
'            End With
'            With o_Excel.Selection.Borders(xlEdgeTop)
'                .LineStyle = xlContinuous
'                .Weight = xlMedium
'                .ColorIndex = xlAutomatic
'            End With
'            With o_Excel.Selection.Borders(xlEdgeBottom)
'                .LineStyle = xlContinuous
'                .Weight = xlMedium
'                .ColorIndex = xlAutomatic
'            End With
'            With o_Excel.Selection.Borders(xlEdgeRight)
'                .LineStyle = xlContinuous
'                .Weight = xlMedium
'                .ColorIndex = xlAutomatic
'            End With
'
'            'Configurando Tamaño Fuente
'            .Range("B" & FilaExcel - 2 & ":O" & FilaExcel + 38).Select
'            o_Excel.Selection.Font.Size = 10
'
'            'Configurando Formato de Números
'            .Range("B" & FilaExcel + 2 & ":O" & FilaExcel + 38).Select
'            o_Excel.Selection.Style = "Currency"
'
'            .Cells(FilaExcel, 1).Value = "CONCEPTO"
'            .Cells(FilaExcel, 1).Font.Bold = True
'            .Cells(FilaExcel, 1).HorizontalAlignment = xlVAlignCenter
'            .Cells(FilaExcel + 2, 1).Value = "Rentas Habituales"
'            .Cells(FilaExcel + 3, 1).Value = "Pluriempleo (Neto)"
'            .Cells(FilaExcel + 4, 1).Value = "Ajustes"
'            .Cells(FilaExcel + 6, 1).Value = "GCIA. BRUTA"
'            .Range("A" & FilaExcel + 6 & ":O" & FilaExcel + 6).Select
'            o_Excel.Selection.Font.Bold = True
'            .Cells(FilaExcel + 6, 1).HorizontalAlignment = xlVAlignCenter
'            .Cells(FilaExcel + 8, 1).Value = "Jubilación Personal"
'            .Cells(FilaExcel + 9, 1).Value = "O.Social Personal"
'            .Cells(FilaExcel + 10, 1).Value = "Adherente O.Social"
'            .Cells(FilaExcel + 11, 1).Value = "Seguro Obligatorio"
'            .Cells(FilaExcel + 12, 1).Value = "Cuota Sindical"
'            .Cells(FilaExcel + 14, 1).Value = "GCIA. NETA"
'            .Range("A" & FilaExcel + 14 & ":O" & FilaExcel + 14).Select
'            o_Excel.Selection.Font.Bold = True
'            .Cells(FilaExcel + 14, 1).HorizontalAlignment = xlVAlignCenter
'            .Cells(FilaExcel + 16, 1).Value = "Gtos. Sepelio"
'            .Cells(FilaExcel + 17, 1).Value = "Int. Hipotecarios"
'            .Cells(FilaExcel + 18, 1).Value = "Seguro de Vida"
'            .Cells(FilaExcel + 19, 1).Value = "Serv. Doméstico"
'            .Cells(FilaExcel + 20, 1).Value = "Cuota Médico Asist."
'            .Cells(FilaExcel + 21, 1).Value = "Donaciones"
'            .Cells(FilaExcel + 22, 1).Value = "Gcia. No Imponible"
'            .Cells(FilaExcel + 23, 1).Value = "Conyuge"
'            .Cells(FilaExcel + 24, 1).Value = "Hijos"
'            .Cells(FilaExcel + 25, 1).Value = "Otras Cargas"
'            .Cells(FilaExcel + 26, 1).Value = "Deducción Especial"
'            .Cells(FilaExcel + 28, 1).Value = "GCIA. IMPONIBLE"
'            .Range("A" & FilaExcel + 28 & ":O" & FilaExcel + 28).Select
'            o_Excel.Selection.Font.Bold = True
'            .Cells(FilaExcel + 28, 1).HorizontalAlignment = xlVAlignCenter
'            .Cells(FilaExcel + 30, 1).Value = "Porcentaje"
'            .Cells(FilaExcel + 31, 1).Value = "Importe Fijo"
'            .Cells(FilaExcel + 32, 1).Value = "Importe Variable"
'            .Cells(FilaExcel + 33, 1).Value = "IMP. DETERMINADO"
'            .Range("A" & FilaExcel + 33 & ":O" & FilaExcel + 35).Select
'            o_Excel.Selection.Font.Bold = True
'            .Cells(FilaExcel + 33, 1).HorizontalAlignment = xlVAlignCenter
'            .Cells(FilaExcel + 34, 1).Value = "RETENCIÓN ACUM."
'            .Cells(FilaExcel + 34, 1).HorizontalAlignment = xlVAlignCenter
'            .Cells(FilaExcel + 35, 1).Value = "AJUSTES ACUM."
'            .Cells(FilaExcel + 35, 1).HorizontalAlignment = xlVAlignCenter
'            .Cells(FilaExcel + 37, 1).Value = "RETENER"
'            .Range("A" & FilaExcel + 37 & ":O" & FilaExcel + 38).Select
'            o_Excel.Selection.Font.Bold = True
'            .Cells(FilaExcel + 37, 1).HorizontalAlignment = xlVAlignCenter
'            .Cells(FilaExcel + 38, 1).Value = "REINTEGRAR"
'            .Cells(FilaExcel + 38, 1).HorizontalAlignment = xlVAlignCenter
'            .Cells(FilaExcel + 43, 4).Value = "C.P.MIRTA S.SÁNCHEZ GÓMEZ"
'            .Range("D" & FilaExcel + 43 & ":F" & FilaExcel + 43).MergeCells = True
'            .Range("D" & FilaExcel + 43 & ":F" & FilaExcel + 43).Font.Size = 10
'            .Range("D" & FilaExcel + 43 & ":F" & FilaExcel + 43).Font.Bold = True
'            .Range("D" & FilaExcel + 43 & ":F" & FilaExcel + 43).HorizontalAlignment = xlVAlignCenter
'            .Cells(FilaExcel + 44, 4).Value = "JEFE DPTO.CONTABLE"
'            .Range("D" & FilaExcel + 44 & ":F" & FilaExcel + 44).MergeCells = True
'            .Range("D" & FilaExcel + 44 & ":F" & FilaExcel + 44).Font.Size = 10
'            .Range("D" & FilaExcel + 44 & ":F" & FilaExcel + 44).Font.Bold = True
'            .Range("D" & FilaExcel + 44 & ":F" & FilaExcel + 44).HorizontalAlignment = xlVAlignCenter
'            .Cells(FilaExcel + 45, 4).Value = "IN.VI.CO"
'            .Range("D" & FilaExcel + 45 & ":F" & FilaExcel + 45).MergeCells = True
'            .Range("D" & FilaExcel + 45 & ":F" & FilaExcel + 45).Font.Size = 10
'            .Range("D" & FilaExcel + 45 & ":F" & FilaExcel + 45).Font.Bold = True
'            .Range("D" & FilaExcel + 45 & ":F" & FilaExcel + 45).HorizontalAlignment = xlVAlignCenter
'            .Cells(FilaExcel + 43, 10).Value = "JUAN CARLOS CARBALLO"
'            .Range("J" & FilaExcel + 43 & ":L" & FilaExcel + 43).MergeCells = True
'            .Range("J" & FilaExcel + 43 & ":L" & FilaExcel + 43).Font.Size = 10
'            .Range("J" & FilaExcel + 43 & ":L" & FilaExcel + 43).Font.Bold = True
'            .Range("J" & FilaExcel + 43 & ":L" & FilaExcel + 43).HorizontalAlignment = xlVAlignCenter
'            .Cells(FilaExcel + 44, 10).Value = "A/C Gerencia Económico Financiera"
'            .Range("J" & FilaExcel + 44 & ":L" & FilaExcel + 44).MergeCells = True
'            .Range("J" & FilaExcel + 44 & ":L" & FilaExcel + 44).Font.Size = 10
'            .Range("J" & FilaExcel + 44 & ":L" & FilaExcel + 44).Font.Bold = True
'            .Range("J" & FilaExcel + 44 & ":L" & FilaExcel + 44).HorizontalAlignment = xlVAlignCenter
'            .Cells(FilaExcel + 45, 10).Value = "IN.VI.CO"
'            .Range("J" & FilaExcel + 45 & ":L" & FilaExcel + 45).MergeCells = True
'            .Range("J" & FilaExcel + 45 & ":L" & FilaExcel + 45).Font.Size = 10
'            .Range("J" & FilaExcel + 45 & ":L" & FilaExcel + 45).Font.Bold = True
'            .Range("J" & FilaExcel + 45 & ":L" & FilaExcel + 45).HorizontalAlignment = xlVAlignCenter
'            .Cells(FilaExcel + 44, 15).Value = "Generado por"
'            .Range("N" & FilaExcel + 44 & ":O" & FilaExcel + 44).MergeCells = True
'            .Range("N" & FilaExcel + 44 & ":O" & FilaExcel + 44).Font.Size = 8
'            .Range("N" & FilaExcel + 44 & ":O" & FilaExcel + 44).Font.Bold = True
'            .Range("N" & FilaExcel + 44 & ":O" & FilaExcel + 44).HorizontalAlignment = xlVAlignCenter
'            .Cells(FilaExcel + 45, 15).Value = "SLAVE v 1.0"
'            .Range("N" & FilaExcel + 45 & ":O" & FilaExcel + 45).MergeCells = True
'            .Range("N" & FilaExcel + 45 & ":O" & FilaExcel + 45).Font.Size = 8
'            .Range("N" & FilaExcel + 45 & ":O" & FilaExcel + 45).Font.Bold = True
'            .Range("N" & FilaExcel + 45 & ":O" & FilaExcel + 45).HorizontalAlignment = xlVAlignCenter
'
'            While rstListadoSlave.EOF = False And ColumnaExcel <= 14
'                SQL = "Select * From AGENTES Where PUESTOLABORAL = '" & rstListadoSlave!PuestoLaboral & "'"
'                rstBuscarSlave.Open SQL, dbSlave, adOpenDynamic, adLockOptimistic
'                ApellidoAgente = Split(rstBuscarSlave!NombreCompleto, " ")
'                .Cells(FilaExcel - 2, ColumnaExcel + 1).Value = ApellidoAgente(0)
'                .Cells(FilaExcel - 2, ColumnaExcel + 1).Font.Bold = True
'                .Cells(FilaExcel - 2, ColumnaExcel + 1).HorizontalAlignment = xlVAlignCenter
'                .Cells(FilaExcel - 1, ColumnaExcel + 1).Value = Format(rstBuscarSlave!CUIL, "00-00000000-0")
'                .Cells(FilaExcel - 1, ColumnaExcel + 1).HorizontalAlignment = xlVAlignCenter
'                .Cells(FilaExcel, ColumnaExcel + 1).Value = "Acumulado"
'                .Cells(FilaExcel, ColumnaExcel + 1).HorizontalAlignment = xlVAlignCenter
'                rstBuscarSlave.Close
'                SQL = "Select Sum(LIQUIDACIONGANANCIAS4TACATEGORIA.HABEROPTIMO) As SumaHaberOptimo, Sum(LIQUIDACIONGANANCIAS4TACATEGORIA.PLURIEMPLEO) As SumaPluriempleo, " _
'                & "Sum(LIQUIDACIONGANANCIAS4TACATEGORIA.AJUSTE) As SumaAjuste, Sum(LIQUIDACIONGANANCIAS4TACATEGORIA.JUBILACION) As SumaJubilacion, " _
'                & "Sum(LIQUIDACIONGANANCIAS4TACATEGORIA.OBRASOCIAL) As SumaObraSocial, Sum(LIQUIDACIONGANANCIAS4TACATEGORIA.ADHERENTEOBRASOCIAL) As SumaAdherente, " _
'                & "Sum(LIQUIDACIONGANANCIAS4TACATEGORIA.DONACIONES) As SumaDonaciones, Sum(LIQUIDACIONGANANCIAS4TACATEGORIA.HONORARIOSMEDICOS) As SumaHonorariosMedicos, " _
'                & "Sum(LIQUIDACIONGANANCIAS4TACATEGORIA.SEGURODEVIDAOBLIGATORIO) As SumaSeguroObligatorio, Sum(LIQUIDACIONGANANCIAS4TACATEGORIA.CUOTASINDICAL) As SumaCuotaSindical, " _
'                & "Sum(LIQUIDACIONGANANCIAS4TACATEGORIA.SEGURODEVIDAOPTATIVO) As SumaSeguroOptativo, Sum(LIQUIDACIONGANANCIAS4TACATEGORIA.SERVICIODOMESTICO) As SumaServicioDomestico, " _
'                & "Sum(LIQUIDACIONGANANCIAS4TACATEGORIA.CUOTAMEDICOASISTENCIAL) As SumaCuotaMedicoAsistencial, Sum(LIQUIDACIONGANANCIAS4TACATEGORIA.MINIMONOIMPONIBLE) As SumaMinimoNoImponible, " _
'                & "Sum(LIQUIDACIONGANANCIAS4TACATEGORIA.CONYUGE) As SumaConyuge, Sum(LIQUIDACIONGANANCIAS4TACATEGORIA.HIJO) As SumaHijo, " _
'                & "Sum(LIQUIDACIONGANANCIAS4TACATEGORIA.OTRASCARGASDEFAMILIA) As SumaOtrasCargas, Sum(LIQUIDACIONGANANCIAS4TACATEGORIA.DEDUCCIONESPECIAL) As SumaDeduccionEspecial " _
'                & "From CODIGOLIQUIDACIONES INNER JOIN LIQUIDACIONGANANCIAS4TACATEGORIA ON CODIGOLIQUIDACIONES.Codigo = LIQUIDACIONGANANCIAS4TACATEGORIA.CodigoLiquidacion " _
'                & "Where PUESTOLABORAL = '" & rstListadoSlave!PuestoLaboral & "' And CODIGOLIQUIDACION <= '" & CodigoLiquidacion & "' " _
'                & "And Right(PERIODO,4) = '" & Right(PeriodoLiquidacion, 4) & "'"
'                rstBuscarSlave.Open SQL, dbSlave, adOpenDynamic, adLockOptimistic
'                .Cells(FilaExcel + 2, ColumnaExcel + 1).Value = rstBuscarSlave!SumaHaberOptimo
'                SumaTotal1 = rstBuscarSlave!SumaHaberOptimo
'                .Cells(FilaExcel + 3, ColumnaExcel + 1).Value = rstBuscarSlave!SumaPluriempleo
'                SumaTotal1 = SumaTotal1 + rstBuscarSlave!SumaPluriempleo
'                .Cells(FilaExcel + 4, ColumnaExcel + 1).Value = rstBuscarSlave!SumaAjuste
'                SumaTotal1 = SumaTotal1 + rstBuscarSlave!SumaAjuste
'                .Cells(FilaExcel + 6, ColumnaExcel + 1).Value = SumaTotal1
'                .Cells(FilaExcel + 8, ColumnaExcel + 1).Value = rstBuscarSlave!SumaJubilacion
'                SumaTotal2 = SumaTotal2 + rstBuscarSlave!SumaJubilacion
'                .Cells(FilaExcel + 9, ColumnaExcel + 1).Value = rstBuscarSlave!SumaObraSocial
'                SumaTotal2 = SumaTotal2 + rstBuscarSlave!SumaObraSocial
'                .Cells(FilaExcel + 10, ColumnaExcel + 1).Value = rstBuscarSlave!SumaAdherente
'                SumaTotal2 = SumaTotal2 + rstBuscarSlave!SumaAdherente
'                .Cells(FilaExcel + 11, ColumnaExcel + 1).Value = rstBuscarSlave!SumaSeguroObligatorio
'                SumaTotal2 = SumaTotal2 + rstBuscarSlave!SumaSeguroObligatorio
'                .Cells(FilaExcel + 12, ColumnaExcel + 1).Value = rstBuscarSlave!SumaCuotaSindical
'                SumaTotal2 = SumaTotal2 + rstBuscarSlave!SumaCuotaSindical
'                SumaTotal1 = SumaTotal1 - SumaTotal2
'                SumaTotal2 = 0
'                .Cells(FilaExcel + 14, ColumnaExcel + 1).Value = SumaTotal1
'                .Cells(FilaExcel + 16, ColumnaExcel + 1).Value = 0 'Gastos de Sepelio no tenido en cuenta
'                SumaTotal2 = SumaTotal2 + 0
'                .Cells(FilaExcel + 17, ColumnaExcel + 1).Value = 0 'Intereses Hipotecarios no tenido en cuenta
'                SumaTotal2 = SumaTotal2 + 0
'                .Cells(FilaExcel + 18, ColumnaExcel + 1).Value = rstBuscarSlave!SumaSeguroOptativo
'                SumaTotal2 = SumaTotal2 + rstBuscarSlave!SumaSeguroOptativo
'                .Cells(FilaExcel + 19, ColumnaExcel + 1).Value = rstBuscarSlave!SumaServicioDomestico
'                SumaTotal2 = SumaTotal2 + rstBuscarSlave!SumaServicioDomestico
'                .Cells(FilaExcel + 20, ColumnaExcel + 1).Value = rstBuscarSlave!SumaCuotaMedicoAsistencial
'                SumaTotal2 = SumaTotal2 + rstBuscarSlave!SumaCuotaMedicoAsistencial
'                .Cells(FilaExcel + 21, ColumnaExcel + 1).Value = rstBuscarSlave!SumaDonaciones 'Donaciones no tenidas en cuenta
'                SumaTotal2 = SumaTotal2 + rstBuscarSlave!SumaDonaciones
'                .Cells(FilaExcel + 22, ColumnaExcel + 1).Value = rstBuscarSlave!SumaMinimoNoImponible
'                SumaTotal2 = SumaTotal2 + rstBuscarSlave!SumaMinimoNoImponible
'                .Cells(FilaExcel + 23, ColumnaExcel + 1).Value = rstBuscarSlave!SumaConyuge
'                SumaTotal2 = SumaTotal2 + rstBuscarSlave!SumaConyuge
'                .Cells(FilaExcel + 24, ColumnaExcel + 1).Value = rstBuscarSlave!SumaHijo
'                SumaTotal2 = SumaTotal2 + rstBuscarSlave!SumaHijo
'                .Cells(FilaExcel + 25, ColumnaExcel + 1).Value = rstBuscarSlave!SumaOtrasCargas
'                SumaTotal2 = SumaTotal2 + rstBuscarSlave!SumaOtrasCargas
'                .Cells(FilaExcel + 26, ColumnaExcel + 1).Value = rstBuscarSlave!SumaDeduccionEspecial
'                SumaTotal2 = SumaTotal2 + rstBuscarSlave!SumaDeduccionEspecial
'                SumaTotal1 = SumaTotal1 - SumaTotal2
'                SumaTotal2 = 0
'                .Cells(FilaExcel + 28, ColumnaExcel + 1).Value = SumaTotal1
'                rstBuscarSlave.Close
'                If SumaTotal1 > 0 Then
'                    SQL = "Select * From ESCALAAPLICABLEGANANCIAS Order by IMPORTEMAXIMO Asc"
'                    rstBuscarSlave.Open SQL, dbSlave, adOpenDynamic, adLockOptimistic
'                    If rstBuscarSlave.BOF = False Then
'                        rstBuscarSlave.MoveFirst
'                        Do While rstBuscarSlave.EOF = False
'                            If (rstBuscarSlave!ImporteMaximo / 12 * Left(PeriodoLiquidacion, 2)) > SumaTotal1 Then
'                                .Cells(FilaExcel + 30, ColumnaExcel + 1).Value = rstBuscarSlave!ImporteVariable * 100 & " %"
'                                Exit Do
'                            End If
'                            rstBuscarSlave.MoveNext
'                        Loop
'                        If rstBuscarSlave!ImporteFijo = 0 Then
'                            .Cells(FilaExcel + 31, ColumnaExcel + 1).Value = 0
'                            .Cells(FilaExcel + 32, ColumnaExcel + 1).Value = SumaTotal1 * rstBuscarSlave!ImporteVariable
'                        Else
'                            .Cells(FilaExcel + 31, ColumnaExcel + 1).Value = (rstBuscarSlave!ImporteFijo / 12 * Left(PeriodoLiquidacion, 2))
'                            SumaTotal2 = rstBuscarSlave!ImporteVariable
'                            rstBuscarSlave.MovePrevious
'                            .Cells(FilaExcel + 32, ColumnaExcel + 1).Value = (SumaTotal1 - (rstBuscarSlave!ImporteMaximo / 12 * Left(PeriodoLiquidacion, 2))) * SumaTotal2
'                        End If
'                    End If
'                    .Cells(FilaExcel + 33, ColumnaExcel + 1).Value = CDbl(.Cells(FilaExcel + 31, ColumnaExcel + 1).Value) + CDbl(.Cells(FilaExcel + 32, ColumnaExcel + 1).Value)
'                    rstBuscarSlave.Close
'                Else
'                    .Cells(FilaExcel + 30, ColumnaExcel + 1).Value = 0
'                    .Cells(FilaExcel + 31, ColumnaExcel + 1).Value = 0
'                    .Cells(FilaExcel + 32, ColumnaExcel + 1).Value = 0
'                    .Cells(FilaExcel + 33, ColumnaExcel + 1).Value = 0
'                End If
'                SQL = "Select Sum(LIQUIDACIONSUELDOS.Importe) AS SumaDeImporte " _
'                & "From CODIGOLIQUIDACIONES INNER JOIN LIQUIDACIONSUELDOS ON CODIGOLIQUIDACIONES.Codigo = LIQUIDACIONSUELDOS.CodigoLiquidacion " _
'                & "Where PUESTOLABORAL = '" & rstListadoSlave!PuestoLaboral & "' And CODIGOCONCEPTO= '0276' " _
'                & "And Right(PERIODO,4) = '" & Right(PeriodoLiquidacion, 4) & "' And CODIGO < '" & CodigoLiquidacion & "'"
'                rstBuscarSlave.Open SQL, dbSlave, adOpenForwardOnly, adLockOptimistic
'                If rstBuscarSlave.BOF = False And IsNull(rstBuscarSlave!SumaDeImporte) = False Then
'                    .Cells(FilaExcel + 34, ColumnaExcel + 1).Value = CDbl(-rstBuscarSlave!SumaDeImporte)
'                Else
'                    .Cells(FilaExcel + 34, ColumnaExcel + 1).Value = 0
'                End If
'                rstBuscarSlave.Close
'                .Cells(FilaExcel + 35, ColumnaExcel + 1).Value = rstListadoSlave!AjusteRetencion
'                If rstListadoSlave!Retencion > 0 Then
'                    .Cells(FilaExcel + 37, ColumnaExcel + 1).Value = rstListadoSlave!Retencion
'                Else
'                    .Cells(FilaExcel + 38, ColumnaExcel + 1).Value = rstListadoSlave!Retencion
'                End If
'                SumaTotal1 = 0
'                SumaTotal2 = 0
'                rstListadoSlave.MoveNext
'                ColumnaExcel = ColumnaExcel + 1
'            Wend
'
'            FilaExcel = FilaExcel + 53
'        Loop Until rstListadoSlave.EOF = True
'        .Columns(1).ColumnWidth = 19
'        For ColumnaExcel = 2 To 15
'            .Columns(ColumnaExcel).ColumnWidth = 15
'        Next
'    End With
'    rstListadoSlave.Close
'    Set rstListadoSlave = Nothing
'    Set rstBuscarSlave = Nothing
'
''Configurando Para Impresión
'    With o_Libro.ActiveSheet.PageSetup
'        .LeftHeader = ""
'        .CenterHeader = ""
'        .RightHeader = ""
'        .LeftFooter = ""
'        .CenterFooter = ""
'        .RightFooter = ""
'        .LeftMargin = Application.InchesToPoints(0.393700787401575)
'        .RightMargin = Application.InchesToPoints(0.393700787401575)
'        .TopMargin = Application.InchesToPoints(0.984251968503937)
'        .BottomMargin = Application.InchesToPoints(0.984251968503937)
'        .HeaderMargin = Application.InchesToPoints(0)
'        .FooterMargin = Application.InchesToPoints(0)
'        .PrintHeadings = False
'        .PrintGridlines = False
'        .PrintComments = xlPrintNoComments
'        .PrintQuality = 600
'        .CenterHorizontally = True
'        .CenterVertically = False
'        .Orientation = xlLandscape
'        .Draft = False
'        .PaperSize = xlPaperLegal
'        .FirstPageNumber = xlAutomatic
'        .Order = xlDownThenOver
'        .BlackAndWhite = False
'        .Zoom = 70
'        .PrintErrors = xlPrintErrorsDisplayed
'    End With
'
'
'    o_Libro.Close True, strOutputPath
'    ' -- Cerrar Excel
'    o_Excel.Quit
'    ' -- Terminar instancias
'    Call ReleaseObjects(o_Excel, o_Libro, o_Hoja)
'    GenerarReporteMensualGanancias = True
'Exit Function
'
'' -- Controlador de Errores
'Error_Handler:
'    ' -- Cierra la hoja y el la aplicación Excel
'    If Not o_Libro Is Nothing Then: o_Libro.Close False
'    If Not o_Excel Is Nothing Then: o_Excel.Quit
'    Call ReleaseObjects(o_Excel, o_Libro, o_Hoja)
'    If Err.Number <> 1004 Then MsgBox Err.Description, vbCritical
'End Function

'Public Function GenerarF649Viejo(strOutputPath As String) As Boolean
'    On Error GoTo Error_Handler
'
'    Dim o_Excel                 As Object
'    Dim o_Libro                 As Object
'    Dim o_Hoja                  As Object
'    Dim ColumnaExcel            As Integer
'    Dim FilaExcel               As Integer
'    Dim SQL                     As String
'    Dim datFechaInicio          As Date
'    Dim datFechaFin             As Date
'    Dim dblImporteMensual       As Double
'    Dim intMesesLiquidados      As Integer
'    Dim intFamiliares           As Integer
'    Dim i                       As Integer
'    Dim dblImporteControl       As Double
'    Dim dblGananciaNeta         As Double
'    Dim dblPorcentajeAplicable  As Double
'
'    If ValidarGenerarF649 = True Then
'        ' -- Crea el objeto Excel, el objeto workBook y el objeto sheet
'        Set o_Excel = CreateObject("Excel.Application")
'        Set o_Libro = o_Excel.Workbooks.Add
'        Set o_Hoja = o_Libro.Worksheets(1)
'
'
'        'Configurando Para Impresión
'        With o_Libro.ActiveSheet.PageSetup
'            .LeftHeader = ""
'            .CenterHeader = ""
'            .RightHeader = ""
'            .LeftFooter = ""
'            .CenterFooter = ""
'            .RightFooter = ""
'            .LeftMargin = Application.InchesToPoints(0.196850393700787)
'            .RightMargin = Application.InchesToPoints(0.196850393700787)
'            .TopMargin = Application.InchesToPoints(0.78740157480315)
'            .BottomMargin = Application.InchesToPoints(0.78740157480315)
'            .HeaderMargin = Application.InchesToPoints(0)
'            .FooterMargin = Application.InchesToPoints(0)
'            .PrintHeadings = False
'            .PrintGridlines = False
'            .PrintComments = xlPrintNoComments
'            .PrintQuality = 600
'            .CenterHorizontally = True
'            .CenterVertically = False
'            .Orientation = xlPortrait
'            .Draft = False
'            .PaperSize = xlPaperA4
'            .FirstPageNumber = xlAutomatic
'            .Order = xlDownThenOver
'            .BlackAndWhite = False
'            .Zoom = 80
'            .PrintErrors = xlPrintErrorsDisplayed
'        End With
'
'
'        With o_Hoja
'            'Configurando Ancho de Columnas
'            For ColumnaExcel = 1 To 29
'                .Columns(ColumnaExcel).ColumnWidth = 3.57
'            Next
'
'            'Configurando Ancho de Filas
'            For FilaExcel = 1 To 140
'                .Rows(FilaExcel).RowHeight = 14.25
'            Next
'
'            'Configurando Fuente
'            .Range("A1:AC138").Select
'            o_Excel.Selection.Font.Size = 9
'
'            'Configurar Formulario
'            .Cells(1, 1).Value = "AFIP"
'            .Range("A1:E4").Select
'            With o_Excel.Selection
'                .MergeCells = True
'                .WrapText = True
'                .Font.Size = 42
'                .Font.Bold = True
'                .HorizontalAlignment = xlVAlignCenter
'                .VerticalAlignment = xlVAlignCenter
'            End With
'            With o_Excel.Selection.Borders(xlEdgeLeft)
'                .LineStyle = xlContinuous
'                .Weight = xlMedium
'                .ColorIndex = xlAutomatic
'            End With
'            With o_Excel.Selection.Borders(xlEdgeTop)
'                .LineStyle = xlContinuous
'                .Weight = xlMedium
'                .ColorIndex = xlAutomatic
'            End With
'            With o_Excel.Selection.Borders(xlEdgeBottom)
'                .LineStyle = xlContinuous
'                .Weight = xlMedium
'                .ColorIndex = xlAutomatic
'            End With
'            With o_Excel.Selection.Borders(xlEdgeRight)
'                .LineStyle = xlContinuous
'                .Weight = xlMedium
'                .ColorIndex = xlAutomatic
'            End With
'
'            .Cells(5, 1).Value = "IMPUESTO A LAS GANANCIAS REGIMEN DE RETENCIÓN"
'            .Range("A5:E7").Select
'            With o_Excel.Selection
'                .MergeCells = True
'                .WrapText = True
'                .Font.Size = 10
'                .Font.Bold = True
'                .HorizontalAlignment = xlVAlignCenter
'                .VerticalAlignment = xlVAlignCenter
'            End With
'            .Cells(8, 1).Value = "Sueldos, Jubilaciones,etc."
'            .Range("A8:E8").Select
'            With o_Excel.Selection
'                .MergeCells = True
'                .WrapText = True
'                .Font.Size = 8
'                .Font.Bold = True
'                .HorizontalAlignment = xlVAlignCenter
'                .VerticalAlignment = xlVAlignCenter
'            End With
'            .Range("A5:E8").Select
'            With o_Excel.Selection.Borders(xlEdgeLeft)
'                .LineStyle = xlContinuous
'                .Weight = xlMedium
'                .ColorIndex = xlAutomatic
'            End With
'            With o_Excel.Selection.Borders(xlEdgeTop)
'                .LineStyle = xlContinuous
'                .Weight = xlMedium
'                .ColorIndex = xlAutomatic
'            End With
'            With o_Excel.Selection.Borders(xlEdgeBottom)
'                .LineStyle = xlContinuous
'                .Weight = xlMedium
'                .ColorIndex = xlAutomatic
'            End With
'            With o_Excel.Selection.Borders(xlEdgeRight)
'                .LineStyle = xlContinuous
'                .Weight = xlMedium
'                .ColorIndex = xlAutomatic
'            End With
'
'            .Cells(9, 1).Value = "DECLARACIÓN JURADA"
'            .Range("A9:E9").Select
'            With o_Excel.Selection
'                .MergeCells = True
'                .WrapText = True
'                .HorizontalAlignment = xlVAlignCenter
'                .VerticalAlignment = xlVAlignCenter
'            End With
'            .Cells(10, 1).Value = "En pesos con ctvs."
'            .Range("A10:E10").Select
'            With o_Excel.Selection
'                .MergeCells = True
'                .WrapText = True
'                .HorizontalAlignment = xlVAlignCenter
'                .VerticalAlignment = xlVAlignCenter
'            End With
'            .Range("A9:E10").Select
'            With o_Excel.Selection.Borders(xlEdgeLeft)
'                .LineStyle = xlContinuous
'                .Weight = xlMedium
'                .ColorIndex = xlAutomatic
'            End With
'            With o_Excel.Selection.Borders(xlEdgeTop)
'                .LineStyle = xlContinuous
'                .Weight = xlMedium
'                .ColorIndex = xlAutomatic
'            End With
'            With o_Excel.Selection.Borders(xlEdgeBottom)
'                .LineStyle = xlContinuous
'                .Weight = xlMedium
'                .ColorIndex = xlAutomatic
'            End With
'            With o_Excel.Selection.Borders(xlEdgeRight)
'                .LineStyle = xlContinuous
'                .Weight = xlMedium
'                .ColorIndex = xlAutomatic
'            End With
'
'            .Cells(1, 6).Value = "Sello fechador de recepción"
'            .Range("F1:M10").Select
'            With o_Excel.Selection
'                .MergeCells = True
'                .Font.Bold = True
'                .HorizontalAlignment = xlVAlignCenter
'            End With
'            With o_Excel.Selection.Borders(xlEdgeLeft)
'                .LineStyle = xlContinuous
'                .Weight = xlMedium
'                .ColorIndex = xlAutomatic
'            End With
'            With o_Excel.Selection.Borders(xlEdgeTop)
'                .LineStyle = xlContinuous
'                .Weight = xlMedium
'                .ColorIndex = xlAutomatic
'            End With
'            With o_Excel.Selection.Borders(xlEdgeBottom)
'                .LineStyle = xlContinuous
'                .Weight = xlMedium
'                .ColorIndex = xlAutomatic
'            End With
'            With o_Excel.Selection.Borders(xlEdgeRight)
'                .LineStyle = xlContinuous
'                .Weight = xlMedium
'                .ColorIndex = xlAutomatic
'            End With
'
'            .Cells(1, 14).Value = "F.649"
'            .Range("N1:O2").Select
'            With o_Excel.Selection
'                .MergeCells = True
'                .WrapText = True
'                .Font.Bold = True
'                .HorizontalAlignment = xlVAlignCenter
'                .VerticalAlignment = xlVAlignCenter
'            End With
'            With o_Excel.Selection.Borders(xlEdgeLeft)
'                .LineStyle = xlContinuous
'                .Weight = xlMedium
'                .ColorIndex = xlAutomatic
'            End With
'            With o_Excel.Selection.Borders(xlEdgeTop)
'                .LineStyle = xlContinuous
'                .Weight = xlMedium
'                .ColorIndex = xlAutomatic
'            End With
'            With o_Excel.Selection.Borders(xlEdgeBottom)
'                .LineStyle = xlContinuous
'                .Weight = xlMedium
'                .ColorIndex = xlAutomatic
'            End With
'            With o_Excel.Selection.Borders(xlEdgeRight)
'                .LineStyle = xlContinuous
'                .Weight = xlMedium
'                .ColorIndex = xlAutomatic
'            End With
'
'            .Cells(1, 16).Value = ""
'            .Range("P1:Q2").Select
'            With o_Excel.Selection
'                .MergeCells = True
'                .Font.Size = 12
'                .Font.Bold = True
'                .Interior.ColorIndex = 15
'                .Interior.Pattern = xlSolid
'                .HorizontalAlignment = xlVAlignCenter
'                .VerticalAlignment = xlVAlignCenter
'            End With
'            With o_Excel.Selection.Borders(xlEdgeLeft)
'                .LineStyle = xlContinuous
'                .Weight = xlMedium
'                .ColorIndex = xlAutomatic
'            End With
'            With o_Excel.Selection.Borders(xlEdgeTop)
'                .LineStyle = xlContinuous
'                .Weight = xlMedium
'                .ColorIndex = xlAutomatic
'            End With
'            With o_Excel.Selection.Borders(xlEdgeBottom)
'                .LineStyle = xlContinuous
'                .Weight = xlMedium
'                .ColorIndex = xlAutomatic
'            End With
'            With o_Excel.Selection.Borders(xlEdgeRight)
'                .LineStyle = xlContinuous
'                .Weight = xlMedium
'                .ColorIndex = xlAutomatic
'            End With
'
'            .Cells(1, 19).Value = "ORIGINAL"
'            .Range("S1:T1").Select
'            With o_Excel.Selection
'                .MergeCells = True
'                .Font.Bold = True
'                .HorizontalAlignment = xlVAlignCenter
'                .VerticalAlignment = xlVAlignCenter
'            End With
'            .Cells(1, 22).Select
'            With o_Excel.Selection
'                .Value = ""
'                .Font.Size = 12
'                .Font.Bold = True
'                .HorizontalAlignment = xlVAlignCenter
'                .VerticalAlignment = xlVAlignCenter
'                .Interior.ColorIndex = 15
'                .Interior.Pattern = xlSolid
'            End With
'            With o_Excel.Selection.Borders(xlEdgeLeft)
'                .LineStyle = xlContinuous
'                .Weight = xlMedium
'                .ColorIndex = xlAutomatic
'            End With
'            With o_Excel.Selection.Borders(xlEdgeTop)
'                .LineStyle = xlContinuous
'                .Weight = xlMedium
'                .ColorIndex = xlAutomatic
'            End With
'            With o_Excel.Selection.Borders(xlEdgeBottom)
'                .LineStyle = xlContinuous
'                .Weight = xlMedium
'                .ColorIndex = xlAutomatic
'            End With
'            With o_Excel.Selection.Borders(xlEdgeRight)
'                .LineStyle = xlContinuous
'                .Weight = xlMedium
'                .ColorIndex = xlAutomatic
'            End With
'
'            .Range("V1").Select
'            o_Excel.Selection.Copy
'            .Range("AB1").Select
'            .Paste
'
'            .Cells(1, 24).Value = "RECTIFICATIVA"
'            .Range("X1:Z1").Select
'            With o_Excel.Selection
'                .MergeCells = True
'                .Font.Bold = True
'                .HorizontalAlignment = xlVAlignCenter
'                .VerticalAlignment = xlVAlignCenter
'            End With
'
'            .Cells(2, 18).Value = "(Marcar con X el cuadro correspondiente)"
'            .Range("R2:AC2").Select
'            With o_Excel.Selection
'                .MergeCells = True
'                .WrapText = True
'                .HorizontalAlignment = xlVAlignCenter
'                .VerticalAlignment = xlVAlignCenter
'            End With
'            .Range("R1:AC2").Select
'            With o_Excel.Selection.Borders(xlEdgeLeft)
'                .LineStyle = xlContinuous
'                .Weight = xlMedium
'                .ColorIndex = xlAutomatic
'            End With
'            With o_Excel.Selection.Borders(xlEdgeTop)
'                .LineStyle = xlContinuous
'                .Weight = xlMedium
'                .ColorIndex = xlAutomatic
'            End With
'            With o_Excel.Selection.Borders(xlEdgeBottom)
'                .LineStyle = xlContinuous
'                .Weight = xlMedium
'                .ColorIndex = xlAutomatic
'            End With
'            With o_Excel.Selection.Borders(xlEdgeRight)
'                .LineStyle = xlContinuous
'                .Weight = xlMedium
'                .ColorIndex = xlAutomatic
'            End With
'
'            .Cells(3, 14).Value = "Clave Única de Identificación Tributaria:"
'            .Range("N3:U3").Select
'            With o_Excel.Selection
'                .MergeCells = True
'                .WrapText = True
'                .HorizontalAlignment = xlLeft
'                .VerticalAlignment = xlVAlignCenter
'            End With
'            .Cells(4, 14).Value = ""
'            .Range("N4:U4").Select
'            With o_Excel.Selection
'                .MergeCells = True
'                .Font.Size = 12
'                .Font.Bold = True
'                .Interior.ColorIndex = 15
'                .Interior.Pattern = xlSolid
'                .HorizontalAlignment = xlLeft
'                .VerticalAlignment = xlVAlignCenter
'            End With
'            .Range("N3:U4").Select
'            With o_Excel.Selection.Borders(xlEdgeLeft)
'                .LineStyle = xlContinuous
'                .Weight = xlMedium
'                .ColorIndex = xlAutomatic
'            End With
'            With o_Excel.Selection.Borders(xlEdgeTop)
'                .LineStyle = xlContinuous
'                .Weight = xlMedium
'                .ColorIndex = xlAutomatic
'            End With
'            With o_Excel.Selection.Borders(xlEdgeBottom)
'                .LineStyle = xlContinuous
'                .Weight = xlMedium
'                .ColorIndex = xlAutomatic
'            End With
'            With o_Excel.Selection.Borders(xlEdgeRight)
'                .LineStyle = xlContinuous
'                .Weight = xlMedium
'                .ColorIndex = xlAutomatic
'            End With
'
'            .Range("N3:U4").Select
'            o_Excel.Selection.Copy
'            .Range("V3:AC4").Select
'            .Paste
'            .Cells(3, 22).Value = "Código Único de Identificación Laboral:"
'
'            .Cells(5, 14).Value = "Apellido y Nombres del Beneficiario:"
'            .Range("N5:AC5").Select
'            With o_Excel.Selection
'                .MergeCells = True
'                .WrapText = True
'                .HorizontalAlignment = xlLeft
'                .VerticalAlignment = xlVAlignCenter
'            End With
'            .Cells(6, 14).Value = ""
'            .Range("N6:AC6").Select
'            With o_Excel.Selection
'                .MergeCells = True
'                .Font.Size = 12
'                .Font.Bold = True
'                .Interior.ColorIndex = 15
'                .Interior.Pattern = xlSolid
'                .HorizontalAlignment = xlLeft
'                .VerticalAlignment = xlVAlignCenter
'            End With
'            .Range("N5:AC6").Select
'            With o_Excel.Selection.Borders(xlEdgeLeft)
'                .LineStyle = xlContinuous
'                .Weight = xlMedium
'                .ColorIndex = xlAutomatic
'            End With
'            With o_Excel.Selection.Borders(xlEdgeTop)
'                .LineStyle = xlContinuous
'                .Weight = xlMedium
'                .ColorIndex = xlAutomatic
'            End With
'            With o_Excel.Selection.Borders(xlEdgeBottom)
'                .LineStyle = xlContinuous
'                .Weight = xlMedium
'                .ColorIndex = xlAutomatic
'            End With
'            With o_Excel.Selection.Borders(xlEdgeRight)
'                .LineStyle = xlContinuous
'                .Weight = xlMedium
'                .ColorIndex = xlAutomatic
'            End With
'
'            .Cells(7, 14).Value = "Domicilio - Calle:"
'            .Range("N7:W7").Select
'            With o_Excel.Selection
'                .MergeCells = True
'                .WrapText = True
'                .HorizontalAlignment = xlLeft
'                .VerticalAlignment = xlVAlignCenter
'            End With
'            .Cells(8, 14).Value = ""
'            .Range("N8:W8").Select
'            With o_Excel.Selection
'                .MergeCells = True
'                .Font.Size = 12
'                .Font.Bold = True
'                .Interior.ColorIndex = 15
'                .Interior.Pattern = xlSolid
'                .HorizontalAlignment = xlLeft
'                .VerticalAlignment = xlVAlignCenter
'            End With
'            .Range("N7:W8").Select
'            With o_Excel.Selection.Borders(xlEdgeLeft)
'                .LineStyle = xlContinuous
'                .Weight = xlMedium
'                .ColorIndex = xlAutomatic
'            End With
'            With o_Excel.Selection.Borders(xlEdgeTop)
'                .LineStyle = xlContinuous
'                .Weight = xlMedium
'                .ColorIndex = xlAutomatic
'            End With
'            With o_Excel.Selection.Borders(xlEdgeBottom)
'                .LineStyle = xlContinuous
'                .Weight = xlMedium
'                .ColorIndex = xlAutomatic
'            End With
'            With o_Excel.Selection.Borders(xlEdgeRight)
'                .LineStyle = xlContinuous
'                .Weight = xlMedium
'                .ColorIndex = xlAutomatic
'            End With
'
'            .Cells(7, 24).Value = "Número"
'            .Range("X7:Y7").Select
'            With o_Excel.Selection
'                .MergeCells = True
'                .WrapText = True
'                .HorizontalAlignment = xlVAlignCenter
'                .VerticalAlignment = xlVAlignCenter
'            End With
'            .Cells(8, 24).Value = ""
'            .Range("X8:Y8").Select
'            With o_Excel.Selection
'                .MergeCells = True
'                .Font.Size = 12
'                .Font.Bold = True
'                .Interior.ColorIndex = 15
'                .Interior.Pattern = xlSolid
'                .HorizontalAlignment = xlVAlignCenter
'                .VerticalAlignment = xlVAlignCenter
'            End With
'            .Range("X7:Y8").Select
'            With o_Excel.Selection.Borders(xlEdgeLeft)
'                .LineStyle = xlContinuous
'                .Weight = xlMedium
'                .ColorIndex = xlAutomatic
'            End With
'            With o_Excel.Selection.Borders(xlEdgeTop)
'                .LineStyle = xlContinuous
'                .Weight = xlMedium
'                .ColorIndex = xlAutomatic
'            End With
'            With o_Excel.Selection.Borders(xlEdgeBottom)
'                .LineStyle = xlContinuous
'                .Weight = xlMedium
'                .ColorIndex = xlAutomatic
'            End With
'            With o_Excel.Selection.Borders(xlEdgeRight)
'                .LineStyle = xlContinuous
'                .Weight = xlMedium
'                .ColorIndex = xlAutomatic
'            End With
'
'            .Range("X7:Y8").Select
'            o_Excel.Selection.Copy
'            .Range("Z7:AA8").Select
'            .Paste
'            .Range("AB7:AC8").Select
'            .Paste
'            .Range("AB9:AC10").Select
'            .Paste
'            .Cells(7, 26).Value = "Piso"
'            .Cells(7, 28).Value = "Dpto."
'            .Cells(9, 28).Value = "C.P."
'
'            .Cells(9, 14).Value = "Localidad:"
'            .Range("N9:T9").Select
'            With o_Excel.Selection
'                .MergeCells = True
'                .WrapText = True
'                .HorizontalAlignment = xlLeft
'                .VerticalAlignment = xlVAlignCenter
'            End With
'            .Cells(10, 14).Value = ""
'            .Range("N10:T10").Select
'            With o_Excel.Selection
'                .MergeCells = True
'                .Font.Size = 12
'                .Font.Bold = True
'                .Interior.ColorIndex = 15
'                .Interior.Pattern = xlSolid
'                .HorizontalAlignment = xlLeft
'                .VerticalAlignment = xlVAlignCenter
'            End With
'            .Range("N9:T10").Select
'            With o_Excel.Selection.Borders(xlEdgeLeft)
'                .LineStyle = xlContinuous
'                .Weight = xlMedium
'                .ColorIndex = xlAutomatic
'            End With
'            With o_Excel.Selection.Borders(xlEdgeTop)
'                .LineStyle = xlContinuous
'                .Weight = xlMedium
'                .ColorIndex = xlAutomatic
'            End With
'            With o_Excel.Selection.Borders(xlEdgeBottom)
'                .LineStyle = xlContinuous
'                .Weight = xlMedium
'                .ColorIndex = xlAutomatic
'            End With
'            With o_Excel.Selection.Borders(xlEdgeRight)
'                .LineStyle = xlContinuous
'                .Weight = xlMedium
'                .ColorIndex = xlAutomatic
'            End With
'
'            .Range("N9:T10").Select
'            o_Excel.Selection.Copy
'            .Range("U9:AB10").Select
'            .Paste
'            .Cells(9, 21).Value = "Provincia:"
'
'            .Cells(12, 1).Value = "Dependencia DGI en la que se haya inscripto:"
'            .Range("A12:Y12").Select
'            With o_Excel.Selection
'                .MergeCells = True
'                .WrapText = True
'                .HorizontalAlignment = xlLeft
'                .VerticalAlignment = xlVAlignCenter
'            End With
'            .Cells(13, 1).Value = ""
'            .Range("A13:Y13").Select
'            With o_Excel.Selection
'                .MergeCells = True
'                .Font.Size = 12
'                .Font.Bold = True
'                .Interior.ColorIndex = 15
'                .Interior.Pattern = xlSolid
'                .HorizontalAlignment = xlLeft
'                .VerticalAlignment = xlVAlignCenter
'            End With
'            .Range("A12:Y13").Select
'            With o_Excel.Selection.Borders(xlEdgeLeft)
'                .LineStyle = xlContinuous
'                .Weight = xlMedium
'                .ColorIndex = xlAutomatic
'            End With
'            With o_Excel.Selection.Borders(xlEdgeTop)
'                .LineStyle = xlContinuous
'                .Weight = xlMedium
'                .ColorIndex = xlAutomatic
'            End With
'            With o_Excel.Selection.Borders(xlEdgeBottom)
'                .LineStyle = xlContinuous
'                .Weight = xlMedium
'                .ColorIndex = xlAutomatic
'            End With
'            With o_Excel.Selection.Borders(xlEdgeRight)
'                .LineStyle = xlContinuous
'                .Weight = xlMedium
'                .ColorIndex = xlAutomatic
'            End With
'
'            .Range("N1:O2").Select
'            o_Excel.Selection.Copy
'            .Range("Z12:AA13").Select
'            .Paste
'            .Cells(12, 26).Value = "USO DGI"
'
'            .Range("AB9:AC10").Select
'            o_Excel.Selection.Copy
'            .Range("AB12:AC13").Select
'            .Paste
'            .Cells(12, 28).Value = "CÓDIGO"
'
'            .Cells(15, 1).Value = "DATOS DEL AGENTE DE RETENCIÓN"
'            .Range("A15:AC15").Select
'            With o_Excel.Selection
'                .MergeCells = True
'                .WrapText = True
'                .Font.Bold = True
'                .HorizontalAlignment = xlVAlignCenter
'                .VerticalAlignment = xlVAlignCenter
'            End With
'            With o_Excel.Selection.Borders(xlEdgeLeft)
'                .LineStyle = xlContinuous
'                .Weight = xlMedium
'                .ColorIndex = xlAutomatic
'            End With
'            With o_Excel.Selection.Borders(xlEdgeTop)
'                .LineStyle = xlContinuous
'                .Weight = xlMedium
'                .ColorIndex = xlAutomatic
'            End With
'            With o_Excel.Selection.Borders(xlEdgeBottom)
'                .LineStyle = xlContinuous
'                .Weight = xlMedium
'                .ColorIndex = xlAutomatic
'            End With
'            With o_Excel.Selection.Borders(xlEdgeRight)
'                .LineStyle = xlContinuous
'                .Weight = xlMedium
'                .ColorIndex = xlAutomatic
'            End With
'
'            .Cells(16, 1).Value = "Apellido y Nombres o Razón Social:"
'            .Range("A16:M16").Select
'            With o_Excel.Selection
'                .MergeCells = True
'                .WrapText = True
'                .HorizontalAlignment = xlLeft
'                .VerticalAlignment = xlVAlignCenter
'            End With
'            .Cells(17, 1).Value = ""
'            .Range("A17:M19").Select
'            With o_Excel.Selection
'                .MergeCells = True
'                .WrapText = True
'                .Font.Size = 12
'                .Font.Bold = True
'                .Interior.ColorIndex = 15
'                .Interior.Pattern = xlSolid
'                .HorizontalAlignment = xlGeneral
'                .VerticalAlignment = xlTop
'            End With
'            .Range("A16:M19").Select
'            With o_Excel.Selection.Borders(xlEdgeLeft)
'                .LineStyle = xlContinuous
'                .Weight = xlMedium
'                .ColorIndex = xlAutomatic
'            End With
'            With o_Excel.Selection.Borders(xlEdgeTop)
'                .LineStyle = xlContinuous
'                .Weight = xlMedium
'                .ColorIndex = xlAutomatic
'            End With
'            With o_Excel.Selection.Borders(xlEdgeBottom)
'                .LineStyle = xlContinuous
'                .Weight = xlMedium
'                .ColorIndex = xlAutomatic
'            End With
'            With o_Excel.Selection.Borders(xlEdgeRight)
'                .LineStyle = xlContinuous
'                .Weight = xlMedium
'                .ColorIndex = xlAutomatic
'            End With
'
'            .Cells(16, 14).Value = "Clave Única de Identificación Tributaria:"
'            .Range("N16:Y16").Select
'            With o_Excel.Selection
'                .MergeCells = True
'                .WrapText = True
'                .HorizontalAlignment = xlLeft
'                .VerticalAlignment = xlVAlignCenter
'            End With
'            .Cells(17, 14).Value = ""
'            .Range("N17:Y17").Select
'            With o_Excel.Selection
'                .MergeCells = True
'                .WrapText = True
'                .Font.Size = 12
'                .Font.Bold = True
'                .Interior.ColorIndex = 15
'                .Interior.Pattern = xlSolid
'                .HorizontalAlignment = xlGeneral
'                .VerticalAlignment = xlTop
'            End With
'            .Range("N16:Y17").Select
'            With o_Excel.Selection.Borders(xlEdgeLeft)
'                .LineStyle = xlContinuous
'                .Weight = xlMedium
'                .ColorIndex = xlAutomatic
'            End With
'            With o_Excel.Selection.Borders(xlEdgeTop)
'                .LineStyle = xlContinuous
'                .Weight = xlMedium
'                .ColorIndex = xlAutomatic
'            End With
'            With o_Excel.Selection.Borders(xlEdgeBottom)
'                .LineStyle = xlContinuous
'                .Weight = xlMedium
'                .ColorIndex = xlAutomatic
'            End With
'            With o_Excel.Selection.Borders(xlEdgeRight)
'                .LineStyle = xlContinuous
'                .Weight = xlMedium
'                .ColorIndex = xlAutomatic
'            End With
'
'            .Range("N16:Y17").Select
'            o_Excel.Selection.Copy
'            .Range("N18:Y19").Select
'            .Paste
'            .Cells(18, 14).Value = "Dependencia DGI en la que se haya inscripto:"
'
'            .Cells(16, 26).Value = "Pagos Extraord. (4)"
'            .Range("Z16:AC16").Select
'            With o_Excel.Selection
'                .MergeCells = True
'                .HorizontalAlignment = xlVAlignCenter
'                .VerticalAlignment = xlVAlignCenter
'            End With
'            .Range("AB13:AC13").Select
'            o_Excel.Selection.Copy
'            .Range("Z17:AA17").Select
'            .Paste
'            .Range("AB17:AC17").Select
'            .Paste
'            .Cells(17, 26).Value = "SI"
'            .Cells(17, 28).Value = "NO"
'            .Range("Z16:AC17").Select
'            With o_Excel.Selection.Borders(xlEdgeLeft)
'                .LineStyle = xlContinuous
'                .Weight = xlMedium
'                .ColorIndex = xlAutomatic
'            End With
'            With o_Excel.Selection.Borders(xlEdgeTop)
'                .LineStyle = xlContinuous
'                .Weight = xlMedium
'                .ColorIndex = xlAutomatic
'            End With
'            With o_Excel.Selection.Borders(xlEdgeBottom)
'                .LineStyle = xlContinuous
'                .Weight = xlMedium
'                .ColorIndex = xlAutomatic
'            End With
'            With o_Excel.Selection.Borders(xlEdgeRight)
'                .LineStyle = xlContinuous
'                .Weight = xlMedium
'                .ColorIndex = xlAutomatic
'            End With
'
'            .Range("Z12:AC13").Select
'            o_Excel.Selection.Copy
'            .Range("Z18:AC19").Select
'            .Paste
'
'            .Cells(21, 1).Value = "ESTA DECLARACION JURADA DEBERA SER CONFECCIONADA" _
'            & " POR EL AGENTE DE RETENCION, CONFORME LO DISPUESTO POR EL ARTICULO 18" _
'            & " DE LA RESOLUCION GENERAL NRO. 4139 Y DEBERA SER PRESENTADA CUANDO EL" _
'            & " IMPORTE DEL RUBRO 3 DE ESTE FORMULARIO SEA IGUAL O SUPERIOR AL IMPORTE" _
'            & " QUE A DICHOS EFECTOS, ESTABLECE EL ART.21 DE LA MISMA."
'            .Range("A21:AC23").Select
'            With o_Excel.Selection
'                .MergeCells = True
'                .WrapText = True
'                .Font.Bold = True
'                .HorizontalAlignment = xlLeft
'                .VerticalAlignment = xlVAlignCenter
'            End With
'            With o_Excel.Selection.Borders(xlEdgeLeft)
'                .LineStyle = xlContinuous
'                .Weight = xlMedium
'                .ColorIndex = xlAutomatic
'            End With
'            With o_Excel.Selection.Borders(xlEdgeTop)
'                .LineStyle = xlContinuous
'                .Weight = xlMedium
'                .ColorIndex = xlAutomatic
'            End With
'            With o_Excel.Selection.Borders(xlEdgeBottom)
'                .LineStyle = xlContinuous
'                .Weight = xlMedium
'                .ColorIndex = xlAutomatic
'            End With
'            With o_Excel.Selection.Borders(xlEdgeRight)
'                .LineStyle = xlContinuous
'                .Weight = xlMedium
'                .ColorIndex = xlAutomatic
'            End With
'
'            .Range("X1:Z1").Select
'            o_Excel.Selection.Copy
'            .Range("A25:C25").Select
'            .Paste
'            .Range("AB13:AC13").Select
'            o_Excel.Selection.Copy
'            .Range("D25:E25").Select
'            .Paste
'            .Range("F25:G25").Select
'            .Paste
'            .Cells(25, 1).Value = "LIQUIDACION:"
'            .Cells(25, 4).Value = "Anual"
'            .Cells(25, 6).Value = "Final"
'            .Cells(25, 8).NumberFormat = "@"
'            .Cells(25, 8).Value = "(1)"
'            .Cells(25, 9).Value = "Comprendidos entre el"
'            .Cells(25, 17).Value = "y el"
'
'            .Cells(25, 16).Value = ""
'            .Range("N25:P25").Select
'            With o_Excel.Selection
'                .MergeCells = True
'                .Font.Bold = True
'                .WrapText = True
'                .Font.Size = 12
'                .Interior.ColorIndex = 15
'                .Interior.Pattern = xlSolid
'                .HorizontalAlignment = xlVAlignCenter
'                .VerticalAlignment = xlVAlignCenter
'            End With
'            With o_Excel.Selection.Borders(xlEdgeLeft)
'                .LineStyle = xlContinuous
'                .Weight = xlMedium
'                .ColorIndex = xlAutomatic
'            End With
'            With o_Excel.Selection.Borders(xlEdgeTop)
'                .LineStyle = xlContinuous
'                .Weight = xlMedium
'                .ColorIndex = xlAutomatic
'            End With
'            With o_Excel.Selection.Borders(xlEdgeBottom)
'                .LineStyle = xlContinuous
'                .Weight = xlMedium
'                .ColorIndex = xlAutomatic
'            End With
'            With o_Excel.Selection.Borders(xlEdgeRight)
'                .LineStyle = xlContinuous
'                .Weight = xlMedium
'                .ColorIndex = xlAutomatic
'            End With
'
'            .Range("N25:P25").Select
'            o_Excel.Selection.Copy
'            .Range("R25:T25").Select
'            .Paste
'
'            .Range("A25:AC25").Select
'            With o_Excel.Selection.Borders(xlEdgeLeft)
'                .LineStyle = xlContinuous
'                .Weight = xlMedium
'                .ColorIndex = xlAutomatic
'            End With
'            With o_Excel.Selection.Borders(xlEdgeTop)
'                .LineStyle = xlContinuous
'                .Weight = xlMedium
'                .ColorIndex = xlAutomatic
'            End With
'            With o_Excel.Selection.Borders(xlEdgeBottom)
'                .LineStyle = xlContinuous
'                .Weight = xlMedium
'                .ColorIndex = xlAutomatic
'            End With
'            With o_Excel.Selection.Borders(xlEdgeRight)
'                .LineStyle = xlContinuous
'                .Weight = xlMedium
'                .ColorIndex = xlAutomatic
'            End With
'
'            .Cells(26, 1).Value = "Rub."
'            .Cells(26, 1).Select
'            With o_Excel.Selection
'                .MergeCells = True
'                .WrapText = True
'                .Font.Bold = True
'                .HorizontalAlignment = xlVAlignCenter
'                .VerticalAlignment = xlVAlignCenter
'            End With
'            With o_Excel.Selection.Borders(xlEdgeLeft)
'                .LineStyle = xlContinuous
'                .Weight = xlMedium
'                .ColorIndex = xlAutomatic
'            End With
'            With o_Excel.Selection.Borders(xlEdgeTop)
'                .LineStyle = xlContinuous
'                .Weight = xlMedium
'                .ColorIndex = xlAutomatic
'            End With
'            With o_Excel.Selection.Borders(xlEdgeBottom)
'                .LineStyle = xlContinuous
'                .Weight = xlMedium
'                .ColorIndex = xlAutomatic
'            End With
'            With o_Excel.Selection.Borders(xlEdgeRight)
'                .LineStyle = xlContinuous
'                .Weight = xlMedium
'                .ColorIndex = xlAutomatic
'            End With
'
'            .Cells(26, 1).Select
'            o_Excel.Selection.Copy
'            .Cells(26, 2).Select
'            .Paste
'            .Cells(26, 25).Select
'            .Paste
'            .Cells(26, 2).Value = "Ins."
'            .Cells(26, 25).Value = "COD"
'
'            .Cells(26, 3).Value = "DETERMINACION DE LA GANANCIA NETA Y LIQUIDACION DEL IMPUESTO"
'            .Range("C26:X26").Select
'            With o_Excel.Selection
'                .MergeCells = True
'                .WrapText = True
'                .Font.Bold = True
'                .HorizontalAlignment = xlLeft
'                .VerticalAlignment = xlVAlignCenter
'            End With
'            .Cells(26, 26).Value = "IMPORTES"
'            .Range("Z26:AC26").Select
'            With o_Excel.Selection
'                .MergeCells = True
'                .WrapText = True
'                .Font.Bold = True
'                .HorizontalAlignment = xlVAlignCenter
'                .VerticalAlignment = xlVAlignCenter
'            End With
'            .Range("A26:AC26").Select
'            With o_Excel.Selection.Borders(xlEdgeLeft)
'                .LineStyle = xlContinuous
'                .Weight = xlMedium
'                .ColorIndex = xlAutomatic
'            End With
'            With o_Excel.Selection.Borders(xlEdgeTop)
'                .LineStyle = xlContinuous
'                .Weight = xlMedium
'                .ColorIndex = xlAutomatic
'            End With
'            With o_Excel.Selection.Borders(xlEdgeBottom)
'                .LineStyle = xlContinuous
'                .Weight = xlMedium
'                .ColorIndex = xlAutomatic
'            End With
'            With o_Excel.Selection.Borders(xlEdgeRight)
'                .LineStyle = xlContinuous
'                .Weight = xlMedium
'                .ColorIndex = xlAutomatic
'            End With
'
'            .Range("A27:A64").Select
'            With o_Excel.Selection
'                .NumberFormat = "@"
'                .HorizontalAlignment = xlVAlignCenter
'                .VerticalAlignment = xlVAlignCenter
'            End With
'            With o_Excel.Selection.Borders(xlEdgeLeft)
'                .LineStyle = xlContinuous
'                .Weight = xlMedium
'                .ColorIndex = xlAutomatic
'            End With
'            With o_Excel.Selection.Borders(xlEdgeTop)
'                .LineStyle = xlContinuous
'                .Weight = xlMedium
'                .ColorIndex = xlAutomatic
'            End With
'            With o_Excel.Selection.Borders(xlEdgeBottom)
'                .LineStyle = xlContinuous
'                .Weight = xlMedium
'                .ColorIndex = xlAutomatic
'            End With
'            With o_Excel.Selection.Borders(xlEdgeRight)
'                .LineStyle = xlContinuous
'                .Weight = xlMedium
'                .ColorIndex = xlAutomatic
'            End With
'
'            .Range("A27:A64").Select
'            o_Excel.Selection.Copy
'            .Range("B27:B64").Select
'            .Paste
'            .Range("Y27:Y64").Select
'            .Paste
'
'            .Cells(27, 1).Value = "1"
'            .Cells(36, 1).Value = "2"
'            .Cells(44, 1).Value = "3"
'            .Cells(45, 1).Value = "4"
'            .Cells(46, 1).Value = "5"
'            .Cells(47, 1).Value = "6"
'            .Cells(55, 1).Value = "7"
'            .Cells(56, 1).Value = "8"
'            .Cells(57, 1).Value = "9"
'            .Cells(61, 1).Value = "10"
'
'            .Cells(28, 2).Value = "a"
'            .Cells(29, 2).Value = "b"
'            .Cells(37, 2).Value = "a"
'            .Cells(38, 2).Value = "b"
'            .Cells(39, 2).Value = "c"
'            .Cells(40, 2).Value = "d"
'            .Cells(41, 2).Value = "e"
'            .Cells(42, 2).Value = "f"
'            .Cells(48, 2).Value = "a"
'            .Cells(49, 2).Value = "b"
'            .Cells(50, 2).Value = "c"
'            .Cells(58, 2).Value = "a"
'            .Cells(59, 2).Value = "b"
'            .Cells(62, 2).Value = "a"
'            .Cells(63, 2).Value = "b"
'
'            .Cells(28, 25).Value = "019"
'            .Cells(31, 25).Value = "027"
'            .Cells(32, 25).Value = "035"
'            .Cells(33, 25).Value = "043"
'            .Cells(34, 25).Value = "078"
'            .Cells(35, 25).Value = "094"
'            .Cells(37, 25).Value = "116"
'            .Cells(38, 25).Value = "124"
'            .Cells(39, 25).Value = "132"
'            .Cells(40, 25).Value = "140"
'            .Cells(41, 25).Value = "159"
'            .Cells(42, 25).Value = "167"
'            .Cells(43, 25).Value = "175"
'            .Cells(44, 25).Value = "183"
'            .Cells(45, 25).Value = "191"
'            .Cells(46, 25).Value = "205"
'            .Cells(48, 25).Value = "213"
'            .Cells(49, 25).Value = "221"
'            .Cells(51, 25).Value = "256"
'            .Cells(52, 25).Value = "264"
'            .Cells(53, 25).Value = "272"
'            .Cells(54, 25).Value = "302"
'            .Cells(55, 25).Value = "310"
'            .Cells(56, 25).Value = "329"
'            .Cells(58, 25).Value = "345"
'            .Cells(59, 25).Value = "353"
'            .Cells(60, 25).Value = "361"
'            .Cells(62, 25).Value = "388"
'            .Cells(63, 25).Value = "393"
'
'            .Cells(27, 3).Value = "IMPORTE BRUTO DE LAS GANANCIAS"
'            .Cells(27, 3).Font.Bold = True
'            .Cells(28, 4).Value = "Liquidadas por la entidad que actúa como agente de retención"
'            .Cells(29, 4).Value = "Liquidadas por otras personas o entidades"
'
'            .Cells(30, 3).Value = "Apellido y Nombres o denominación y domicilio"
'            .Range("C30:T30").Select
'            With o_Excel.Selection
'                .MergeCells = True
'                .WrapText = True
'                .HorizontalAlignment = xlVAlignCenter
'                .VerticalAlignment = xlVAlignCenter
'            End With
'            .Range("C31:T31").Select
'            With o_Excel.Selection
'                .MergeCells = True
'                .WrapText = True
'                .Font.Size = 12
'                .Interior.ColorIndex = 15
'                .Interior.Pattern = xlSolid
'                .Font.Bold = True
'                .HorizontalAlignment = xlLeft
'                .VerticalAlignment = xlVAlignCenter
'            End With
'            .Range("C31:T31").Select
'            o_Excel.Selection.Copy
'            .Range("C32:T32").Select
'            .Paste
'            .Range("C33:T33").Select
'            .Paste
'            .Range("C34:T34").Select
'            .Paste
'
'            .Cells(30, 21).Value = "Nro. de C.U.I.T."
'            .Range("U30:X30").Select
'            With o_Excel.Selection
'                .MergeCells = True
'                .WrapText = True
'                .HorizontalAlignment = xlVAlignCenter
'                .VerticalAlignment = xlVAlignCenter
'            End With
'            .Range("U31:X31").Select
'            With o_Excel.Selection
'                .MergeCells = True
'                .WrapText = True
'                .Font.Size = 12
'                .Interior.ColorIndex = 15
'                .Interior.Pattern = xlSolid
'                .Font.Bold = True
'                .HorizontalAlignment = xlVAlignCenter
'                .VerticalAlignment = xlVAlignCenter
'            End With
'            .Range("U31:X31").Select
'            o_Excel.Selection.Copy
'            .Range("U32:X32").Select
'            .Paste
'            .Range("U33:X33").Select
'            .Paste
'            .Range("U34:X34").Select
'            .Paste
'            .Range("C30:T34").Select
'            With o_Excel.Selection.Borders(xlEdgeLeft)
'                .LineStyle = xlContinuous
'                .Weight = xlMedium
'                .ColorIndex = xlAutomatic
'            End With
'            With o_Excel.Selection.Borders(xlEdgeTop)
'                .LineStyle = xlContinuous
'                .Weight = xlMedium
'                .ColorIndex = xlAutomatic
'            End With
'            With o_Excel.Selection.Borders(xlEdgeBottom)
'                .LineStyle = xlContinuous
'                .Weight = xlMedium
'                .ColorIndex = xlAutomatic
'            End With
'            With o_Excel.Selection.Borders(xlEdgeRight)
'                .LineStyle = xlContinuous
'                .Weight = xlMedium
'                .ColorIndex = xlAutomatic
'            End With
'            .Range("U30:X34").Select
'            With o_Excel.Selection.Borders(xlEdgeLeft)
'                .LineStyle = xlContinuous
'                .Weight = xlMedium
'                .ColorIndex = xlAutomatic
'            End With
'            With o_Excel.Selection.Borders(xlEdgeTop)
'                .LineStyle = xlContinuous
'                .Weight = xlMedium
'                .ColorIndex = xlAutomatic
'            End With
'            With o_Excel.Selection.Borders(xlEdgeBottom)
'                .LineStyle = xlContinuous
'                .Weight = xlMedium
'                .ColorIndex = xlAutomatic
'            End With
'            With o_Excel.Selection.Borders(xlEdgeRight)
'                .LineStyle = xlContinuous
'                .Weight = xlMedium
'                .ColorIndex = xlAutomatic
'            End With
'
'            .Cells(35, 3).Value = "IMPORTE BRUTO DE LAS GANANCIAS"
'            .Cells(35, 3).Font.Bold = True
'
'            .Cells(36, 3).Value = "DEDUCCIONES Y DESGRAVACIONES"
'            .Cells(36, 3).Font.Bold = True
'            .Cells(37, 4).Value = "Aportes Jubilatorios"
'            .Cells(38, 4).Value = "Aportes para obras sociales y cuotas médico asistenciales (total del rubro 11)"
'            .Cells(39, 4).Value = "Primas de seguro para el caso de muerte (total del rubro 12)"
'            .Cells(40, 4).Value = "Gastos de sepelio (total del rubro 13)"
'            .Cells(41, 4).Value = "Gastos estimativos de corredores y viajantes de comercio (movilidad, etc.)"
'            .Cells(42, 4).Value = "Otras deducciones (total del rubro 15)"
'            .Cells(43, 3).Value = "TOTAL DEL RUBRO 2 (suma de los incisos a) al f)"
'            .Cells(43, 3).Font.Bold = True
'
'            .Cells(44, 3).Value = "RESULTADO NETO (Diferencia entre el rubro 1 y el rubro 2)"
'            .Cells(44, 3).Font.Bold = True
'
'            .Cells(45, 3).Value = "DONACIONES (Hasta el límite del 5% del rubro 3)"
'            .Cells(45, 3).Font.Bold = True
'
'            .Cells(46, 3).Value = "DIFERENCIA (Rubro 3 menos rubro 4)"
'            .Cells(46, 3).Font.Bold = True
'
'            .Cells(47, 3).Value = "DEDUCCION ESPECIAL, GANANCIAS NO IMPONIBLES Y CARGAS DE FAMILIA"
'            .Cells(47, 3).Font.Bold = True
'            .Cells(48, 4).Value = "Deducción especial"
'            .Cells(49, 4).Value = "Ganancia no imponible"
'            .Cells(50, 4).Value = "Cargas de familia (6)"
'            .Cells(51, 5).Value = "Cónyuge"
'            .Cells(52, 5).Value = "Hijos"
'            .Cells(53, 5).Value = "Otras cargas"
'            .Cells(54, 3).Value = "TOTALES DEL RUBRO 6 (suma de los incisos a), b) y c))"
'            .Cells(54, 3).Font.Bold = True
'
'            .Cells(55, 3).Value = "GANANCIAS NETAS SUJETAS A IMPUESTO (diferencia entre el rubro 5 y 6)"
'            .Cells(55, 3).Font.Bold = True
'
'            .Cells(56, 3).Value = "TOTAL DEL IMPUESTO DETERMINADO"
'            .Cells(56, 3).Font.Bold = True
'
'            .Cells(57, 3).Value = "MONTOS COMPUTABLES"
'            .Cells(57, 3).Font.Bold = True
'            .Cells(58, 4).Value = "Retenciones efectuadas en el período fiscal que se liquida"
'            .Cells(59, 4).Value = "Regímenes de promoción (Rebaja de Impuesto, Diferimiento u otros)"
'            .Cells(60, 3).Value = "TOTALES DEL RUBRO 9 (suma de los incisos a) y b))"
'            .Cells(60, 3).Font.Bold = True
'
'            .Cells(61, 3).Value = "SALDO DEL IMPUESTO (Diferencia entre el rubro 8 y rubro 9)"
'            .Cells(61, 3).Font.Bold = True
'            .Cells(62, 4).Value = "A favor D.G.I."
'            .Cells(63, 4).Value = "A favor Beneficiario"
'            .Cells(64, 4).Value = "O sea pesos"
'            .Range("G64:X64").Select
'            With o_Excel.Selection
'                .MergeCells = True
'                .WrapText = True
'                .Font.Size = 12
'                .Interior.ColorIndex = 15
'                .Interior.Pattern = xlSolid
'                .Font.Bold = True
'                .HorizontalAlignment = xlLeft
'                .VerticalAlignment = xlVAlignCenter
'            End With
'
'            .Range("Z28:AC28").Select
'            With o_Excel.Selection
'                .MergeCells = True
'                .WrapText = True
'                .Font.Size = 12
'                .NumberFormat = "#,##0.00"
'                .Interior.ColorIndex = 15
'                .Interior.Pattern = xlSolid
'                .Font.Bold = True
'                .HorizontalAlignment = xlRight
'                .VerticalAlignment = xlVAlignCenter
'            End With
'            .Range("Z28:AC28").Select
'            o_Excel.Selection.Copy
'            .Range("Z31:AC31").Select
'            .Paste
'            .Range("Z32:AC32").Select
'            .Paste
'            .Range("Z33:AC33").Select
'            .Paste
'            .Range("Z34:AC34").Select
'            .Paste
'            .Range("Z35:AC35").Select
'            .Paste
'            .Range("Z37:AC37").Select
'            .Paste
'            .Range("Z38:AC38").Select
'            .Paste
'            .Range("Z39:AC39").Select
'            .Paste
'            .Range("Z40:AC40").Select
'            .Paste
'            .Range("Z41:AC41").Select
'            .Paste
'            .Range("Z42:AC42").Select
'            .Paste
'            .Range("Z43:AC43").Select
'            .Paste
'            .Range("Z44:AC44").Select
'            .Paste
'            .Range("Z45:AC45").Select
'            .Paste
'            .Range("Z46:AC46").Select
'            .Paste
'            .Range("Z48:AC48").Select
'            .Paste
'            .Range("Z49:AC49").Select
'            .Paste
'            .Range("Z51:AC51").Select
'            .Paste
'            .Range("Z51:AC51").Select
'            .Paste
'            .Range("Z52:AC52").Select
'            .Paste
'            .Range("Z53:AC53").Select
'            .Paste
'            .Range("Z54:AC54").Select
'            .Paste
'            .Range("Z55:AC55").Select
'            .Paste
'            .Range("Z56:AC56").Select
'            .Paste
'            .Range("Z58:AC58").Select
'            .Paste
'            .Range("Z59:AC59").Select
'            .Paste
'            .Range("Z60:AC60").Select
'            .Paste
'            .Range("Z62:AC62").Select
'            .Paste
'            .Range("Z63:AC63").Select
'            .Paste
'
'            .Range("Z27:AC64").Select
'            With o_Excel.Selection.Borders(xlEdgeLeft)
'                .LineStyle = xlContinuous
'                .Weight = xlMedium
'                .ColorIndex = xlAutomatic
'            End With
'            With o_Excel.Selection.Borders(xlEdgeTop)
'                .LineStyle = xlContinuous
'                .Weight = xlMedium
'                .ColorIndex = xlAutomatic
'            End With
'            With o_Excel.Selection.Borders(xlEdgeBottom)
'                .LineStyle = xlContinuous
'                .Weight = xlMedium
'                .ColorIndex = xlAutomatic
'            End With
'            With o_Excel.Selection.Borders(xlEdgeRight)
'                .LineStyle = xlContinuous
'                .Weight = xlMedium
'                .ColorIndex = xlAutomatic
'            End With
'
'            .Range("C27:AC35").Select
'            With o_Excel.Selection.Borders(xlEdgeLeft)
'                .LineStyle = xlContinuous
'                .Weight = xlMedium
'                .ColorIndex = xlAutomatic
'            End With
'            With o_Excel.Selection.Borders(xlEdgeTop)
'                .LineStyle = xlContinuous
'                .Weight = xlMedium
'                .ColorIndex = xlAutomatic
'            End With
'            With o_Excel.Selection.Borders(xlEdgeBottom)
'                .LineStyle = xlContinuous
'                .Weight = xlMedium
'                .ColorIndex = xlAutomatic
'            End With
'            With o_Excel.Selection.Borders(xlEdgeRight)
'                .LineStyle = xlContinuous
'                .Weight = xlMedium
'                .ColorIndex = xlAutomatic
'            End With
'
'            .Range("C36:AC43").Select
'            With o_Excel.Selection.Borders(xlEdgeLeft)
'                .LineStyle = xlContinuous
'                .Weight = xlMedium
'                .ColorIndex = xlAutomatic
'            End With
'            With o_Excel.Selection.Borders(xlEdgeTop)
'                .LineStyle = xlContinuous
'                .Weight = xlMedium
'                .ColorIndex = xlAutomatic
'            End With
'            With o_Excel.Selection.Borders(xlEdgeBottom)
'                .LineStyle = xlContinuous
'                .Weight = xlMedium
'                .ColorIndex = xlAutomatic
'            End With
'            With o_Excel.Selection.Borders(xlEdgeRight)
'                .LineStyle = xlContinuous
'                .Weight = xlMedium
'                .ColorIndex = xlAutomatic
'            End With
'
'            .Range("C45:AC45").Select
'            With o_Excel.Selection.Borders(xlEdgeLeft)
'                .LineStyle = xlContinuous
'                .Weight = xlMedium
'                .ColorIndex = xlAutomatic
'            End With
'            With o_Excel.Selection.Borders(xlEdgeTop)
'                .LineStyle = xlContinuous
'                .Weight = xlMedium
'                .ColorIndex = xlAutomatic
'            End With
'            With o_Excel.Selection.Borders(xlEdgeBottom)
'                .LineStyle = xlContinuous
'                .Weight = xlMedium
'                .ColorIndex = xlAutomatic
'            End With
'            With o_Excel.Selection.Borders(xlEdgeRight)
'                .LineStyle = xlContinuous
'                .Weight = xlMedium
'                .ColorIndex = xlAutomatic
'            End With
'
'            .Range("C47:AC54").Select
'            With o_Excel.Selection.Borders(xlEdgeLeft)
'                .LineStyle = xlContinuous
'                .Weight = xlMedium
'                .ColorIndex = xlAutomatic
'            End With
'            With o_Excel.Selection.Borders(xlEdgeTop)
'                .LineStyle = xlContinuous
'                .Weight = xlMedium
'                .ColorIndex = xlAutomatic
'            End With
'            With o_Excel.Selection.Borders(xlEdgeBottom)
'                .LineStyle = xlContinuous
'                .Weight = xlMedium
'                .ColorIndex = xlAutomatic
'            End With
'            With o_Excel.Selection.Borders(xlEdgeRight)
'                .LineStyle = xlContinuous
'                .Weight = xlMedium
'                .ColorIndex = xlAutomatic
'            End With
'
'            .Range("C56:AC56").Select
'            With o_Excel.Selection.Borders(xlEdgeLeft)
'                .LineStyle = xlContinuous
'                .Weight = xlMedium
'                .ColorIndex = xlAutomatic
'            End With
'            With o_Excel.Selection.Borders(xlEdgeTop)
'                .LineStyle = xlContinuous
'                .Weight = xlMedium
'                .ColorIndex = xlAutomatic
'            End With
'            With o_Excel.Selection.Borders(xlEdgeBottom)
'                .LineStyle = xlContinuous
'                .Weight = xlMedium
'                .ColorIndex = xlAutomatic
'            End With
'            With o_Excel.Selection.Borders(xlEdgeRight)
'                .LineStyle = xlContinuous
'                .Weight = xlMedium
'                .ColorIndex = xlAutomatic
'            End With
'
'            .Range("C61:AC64").Select
'            With o_Excel.Selection.Borders(xlEdgeLeft)
'                .LineStyle = xlContinuous
'                .Weight = xlMedium
'                .ColorIndex = xlAutomatic
'            End With
'            With o_Excel.Selection.Borders(xlEdgeTop)
'                .LineStyle = xlContinuous
'                .Weight = xlMedium
'                .ColorIndex = xlAutomatic
'            End With
'            With o_Excel.Selection.Borders(xlEdgeBottom)
'                .LineStyle = xlContinuous
'                .Weight = xlMedium
'                .ColorIndex = xlAutomatic
'            End With
'            With o_Excel.Selection.Borders(xlEdgeRight)
'                .LineStyle = xlContinuous
'                .Weight = xlMedium
'                .ColorIndex = xlAutomatic
'            End With
'
'            .Range("A26:AC26").Select
'            o_Excel.Selection.Copy
'            .Range("A66:AC66").Select
'            .Paste
'
'            .Range("Z63:AC63").Select
'            o_Excel.Selection.Copy
'            .Range("Z69:AC69").Select
'            .Paste
'            .Range("Z70:AC70").Select
'            .Paste
'            .Range("Z71:AC71").Select
'            .Paste
'            .Range("Z74:AC74").Select
'            .Paste
'            .Range("Z75:AC75").Select
'            .Paste
'            .Range("Z78:AC78").Select
'            .Paste
'            .Range("Z79:AC79").Select
'            .Paste
'            .Range("Z80:AC80").Select
'            .Paste
'            .Range("Z83:AC83").Select
'            .Paste
'            .Range("Z84:AC84").Select
'            .Paste
'            .Range("Z85:AC85").Select
'            .Paste
'            .Range("Z88:AC88").Select
'            .Paste
'            .Range("Z89:AC89").Select
'            .Paste
'            .Range("Z90:AC90").Select
'            .Paste
'            .Range("Z91:AC91").Select
'            .Paste
'
'            .Range("A67:A91").Select
'            With o_Excel.Selection
'                .NumberFormat = "@"
'                .HorizontalAlignment = xlVAlignCenter
'                .VerticalAlignment = xlVAlignCenter
'            End With
'            With o_Excel.Selection.Borders(xlEdgeLeft)
'                .LineStyle = xlContinuous
'                .Weight = xlMedium
'                .ColorIndex = xlAutomatic
'            End With
'            With o_Excel.Selection.Borders(xlEdgeTop)
'                .LineStyle = xlContinuous
'                .Weight = xlMedium
'                .ColorIndex = xlAutomatic
'            End With
'            With o_Excel.Selection.Borders(xlEdgeBottom)
'                .LineStyle = xlContinuous
'                .Weight = xlMedium
'                .ColorIndex = xlAutomatic
'            End With
'            With o_Excel.Selection.Borders(xlEdgeRight)
'                .LineStyle = xlContinuous
'                .Weight = xlMedium
'                .ColorIndex = xlAutomatic
'            End With
'
'            .Range("A67:A91").Select
'            o_Excel.Selection.Copy
'            .Range("B67:B91").Select
'            .Paste
'            .Range("Y67:Y91").Select
'            .Paste
'
'            .Cells(67, 1).Value = "11"
'            .Cells(72, 1).Value = "12"
'            .Cells(76, 1).Value = "13"
'            .Cells(81, 1).Value = "14"
'            .Cells(86, 1).Value = "15"
'
'            .Cells(69, 2).Value = "a"
'            .Cells(70, 2).Value = "b"
'            .Cells(74, 2).Value = "a"
'            .Cells(78, 2).Value = "a"
'            .Cells(79, 2).Value = "b"
'            .Cells(83, 2).Value = "a"
'            .Cells(84, 2).Value = "b"
'            .Cells(88, 2).Value = "a"
'            .Cells(89, 2).Value = "b"
'            .Cells(90, 2).Value = "c"
'
'            .Cells(69, 25).Value = "418"
'            .Cells(70, 25).Value = "426"
'            .Cells(71, 25).Value = "434"
'            .Cells(74, 25).Value = "507"
'            .Cells(75, 25).Value = "515"
'            .Cells(78, 25).Value = "604"
'            .Cells(79, 25).Value = "612"
'            .Cells(80, 25).Value = "620"
'            .Cells(83, 25).Value = "701"
'            .Cells(84, 25).Value = "728"
'            .Cells(85, 25).Value = "736"
'            .Cells(88, 25).Value = "809"
'            .Cells(89, 25).Value = "817"
'            .Cells(90, 25).Value = "825"
'            .Cells(91, 25).Value = "833"
'
'            .Cells(67, 3).Value = "CUOTAS MEDICO ASISTENCIALES"
'            .Cells(67, 3).Font.Bold = True
'            .Range("C30:X31").Select
'            o_Excel.Selection.Copy
'            .Range("C68:X69").Select
'            .Paste
'            .Cells(68, 3).Value = "Denominación y domicilio de la empresa asistencial"
'            .Range("C34:X34").Select
'            o_Excel.Selection.Copy
'            .Range("C70:X70").Select
'            .Paste
'            .Cells(71, 3).Value = "TOTALES DEL RUBRO 11"
'            .Cells(71, 3).Font.Bold = True
'
'            .Range("C67:X68").Select
'            o_Excel.Selection.Copy
'            .Range("C72:X73").Select
'            .Paste
'            .Cells(72, 3).Value = "PRIMAS DE SEGURO"
'            .Cells(73, 3).Value = "Denominación y domicilio de la Cia.Aseguradora"
'            .Range("C70:X71").Select
'            o_Excel.Selection.Copy
'            .Range("C74:X75").Select
'            .Paste
'            .Cells(75, 3).Value = "TOTALES DEL RUBRO 12"
'
'            .Cells(76, 3).Value = "GASTOS DE SEPELIO"
'            .Cells(76, 3).Font.Bold = True
'            .Range("U68:X70").Select
'            o_Excel.Selection.Copy
'            .Range("M77:P79").Select
'            .Paste
'            .Range("Q77:T79").Select
'            .Paste
'            .Range("U77:X79").Select
'            .Paste
'            .Range("Q78:X79").Select
'            With o_Excel.Selection
'                .NumberFormat = "#,##0.00"
'                .HorizontalAlignment = xlRight
'            End With
'            .Cells(77, 17).Value = "Gasto Total"
'            .Cells(77, 21).Value = "Importe Diferido"
'            .Cells(77, 3).Value = "Denominación y domicilio de la Empresa"
'            .Range("C77:L77").Select
'            With o_Excel.Selection
'                .MergeCells = True
'                .WrapText = True
'                .HorizontalAlignment = xlVAlignCenter
'                .VerticalAlignment = xlVAlignCenter
'            End With
'            .Range("C78:L78").Select
'            With o_Excel.Selection
'                .MergeCells = True
'                .WrapText = True
'                .Font.Size = 12
'                .Interior.ColorIndex = 15
'                .Interior.Pattern = xlSolid
'                .Font.Bold = True
'                .HorizontalAlignment = xlLeft
'                .VerticalAlignment = xlVAlignCenter
'            End With
'            .Range("C78:L78").Select
'            o_Excel.Selection.Copy
'            .Range("C79:L79").Select
'            .Paste
'            .Range("C77:L79").Select
'            With o_Excel.Selection.Borders(xlEdgeLeft)
'                .LineStyle = xlContinuous
'                .Weight = xlMedium
'                .ColorIndex = xlAutomatic
'            End With
'            With o_Excel.Selection.Borders(xlEdgeTop)
'                .LineStyle = xlContinuous
'                .Weight = xlMedium
'                .ColorIndex = xlAutomatic
'            End With
'            With o_Excel.Selection.Borders(xlEdgeBottom)
'                .LineStyle = xlContinuous
'                .Weight = xlMedium
'                .ColorIndex = xlAutomatic
'            End With
'            With o_Excel.Selection.Borders(xlEdgeRight)
'                .LineStyle = xlContinuous
'                .Weight = xlMedium
'                .ColorIndex = xlAutomatic
'            End With
'            .Cells(80, 3).Value = "TOTALES DEL RUBRO 13"
'            .Cells(80, 3).Font.Bold = True
'
'            .Range("C76:X80").Select
'            o_Excel.Selection.Copy
'            .Range("C81:X85").Select
'            .Paste
'            .Cells(81, 3).Value = "DONACIONES"
'            .Cells(82, 3).Value = "Entidad Beneficiaria y domicilio"
'            .Cells(85, 3).Value = "TOTALES DEL RUBRO 14"
'
'            .Range("C67:X69").Select
'            o_Excel.Selection.Copy
'            .Range("C86:X88").Select
'            .Paste
'            .Range("C69:X71").Select
'            o_Excel.Selection.Copy
'            .Range("C89:X91").Select
'            .Paste
'            .Cells(86, 3).Value = "OTRAS DEDUCCIONES"
'            .Cells(87, 3).Value = "Norma Legal y Concepto"
'            .Cells(87, 21).Value = "Monto Total"
'            .Range("U88:X90").Select
'            With o_Excel.Selection
'                .NumberFormat = "#,##0.00"
'                .HorizontalAlignment = xlRight
'            End With
'            .Cells(91, 3).Value = "TOTALES DEL RUBRO 15 (suma de los inc a), b) y c) "
'
'            .Range("Z67:AC91").Select
'            With o_Excel.Selection.Borders(xlEdgeLeft)
'                .LineStyle = xlContinuous
'                .Weight = xlMedium
'                .ColorIndex = xlAutomatic
'            End With
'            With o_Excel.Selection.Borders(xlEdgeTop)
'                .LineStyle = xlContinuous
'                .Weight = xlMedium
'                .ColorIndex = xlAutomatic
'            End With
'            With o_Excel.Selection.Borders(xlEdgeBottom)
'                .LineStyle = xlContinuous
'                .Weight = xlMedium
'                .ColorIndex = xlAutomatic
'            End With
'            With o_Excel.Selection.Borders(xlEdgeRight)
'                .LineStyle = xlContinuous
'                .Weight = xlMedium
'                .ColorIndex = xlAutomatic
'            End With
'            .Range("C67:AC71").Select
'            With o_Excel.Selection.Borders(xlEdgeLeft)
'                .LineStyle = xlContinuous
'                .Weight = xlMedium
'                .ColorIndex = xlAutomatic
'            End With
'            With o_Excel.Selection.Borders(xlEdgeTop)
'                .LineStyle = xlContinuous
'                .Weight = xlMedium
'                .ColorIndex = xlAutomatic
'            End With
'            With o_Excel.Selection.Borders(xlEdgeBottom)
'                .LineStyle = xlContinuous
'                .Weight = xlMedium
'                .ColorIndex = xlAutomatic
'            End With
'            With o_Excel.Selection.Borders(xlEdgeRight)
'                .LineStyle = xlContinuous
'                .Weight = xlMedium
'                .ColorIndex = xlAutomatic
'            End With
'            .Range("C76:AC80").Select
'            With o_Excel.Selection.Borders(xlEdgeLeft)
'                .LineStyle = xlContinuous
'                .Weight = xlMedium
'                .ColorIndex = xlAutomatic
'            End With
'            With o_Excel.Selection.Borders(xlEdgeTop)
'                .LineStyle = xlContinuous
'                .Weight = xlMedium
'                .ColorIndex = xlAutomatic
'            End With
'            With o_Excel.Selection.Borders(xlEdgeBottom)
'                .LineStyle = xlContinuous
'                .Weight = xlMedium
'                .ColorIndex = xlAutomatic
'            End With
'            With o_Excel.Selection.Borders(xlEdgeRight)
'                .LineStyle = xlContinuous
'                .Weight = xlMedium
'                .ColorIndex = xlAutomatic
'            End With
'            .Range("C86:AC91").Select
'            With o_Excel.Selection.Borders(xlEdgeLeft)
'                .LineStyle = xlContinuous
'                .Weight = xlMedium
'                .ColorIndex = xlAutomatic
'            End With
'            With o_Excel.Selection.Borders(xlEdgeTop)
'                .LineStyle = xlContinuous
'                .Weight = xlMedium
'                .ColorIndex = xlAutomatic
'            End With
'            With o_Excel.Selection.Borders(xlEdgeBottom)
'                .LineStyle = xlContinuous
'                .Weight = xlMedium
'                .ColorIndex = xlAutomatic
'            End With
'            With o_Excel.Selection.Borders(xlEdgeRight)
'                .LineStyle = xlContinuous
'                .Weight = xlMedium
'                .ColorIndex = xlAutomatic
'            End With
'
'            .Range("A93:AC100").Select
'            With o_Excel.Selection.Borders(xlEdgeLeft)
'                .LineStyle = xlContinuous
'                .Weight = xlMedium
'                .ColorIndex = xlAutomatic
'            End With
'            With o_Excel.Selection.Borders(xlEdgeTop)
'                .LineStyle = xlContinuous
'                .Weight = xlMedium
'                .ColorIndex = xlAutomatic
'            End With
'            With o_Excel.Selection.Borders(xlEdgeBottom)
'                .LineStyle = xlContinuous
'                .Weight = xlMedium
'                .ColorIndex = xlAutomatic
'            End With
'            With o_Excel.Selection.Borders(xlEdgeRight)
'                .LineStyle = xlContinuous
'                .Weight = xlMedium
'                .ColorIndex = xlAutomatic
'            End With
'            .Cells(93, 1).Value = "OBSERVACIONES"
'
'            .Cells(102, 1).Value = "El que suscribe, Don"
'            .Range("E102:P102").Select
'            With o_Excel.Selection
'                .MergeCells = True
'                .WrapText = True
'                .Font.Size = 12
'                .Interior.ColorIndex = 15
'                .Interior.Pattern = xlSolid
'                .Font.Bold = True
'                .HorizontalAlignment = xlVAlignCenter
'                .VerticalAlignment = xlVAlignCenter
'            End With
'            .Cells(102, 17).Value = "en su carácter de (2)"
'            .Range("A103:J103").Select
'            With o_Excel.Selection
'                .MergeCells = True
'                .WrapText = True
'                .Font.Size = 12
'                .Interior.ColorIndex = 15
'                .Interior.Pattern = xlSolid
'                .Font.Bold = True
'                .HorizontalAlignment = xlVAlignCenter
'                .VerticalAlignment = xlVAlignCenter
'            End With
'            .Cells(103, 11).Value = "de la entidad que actúa como agente de retención,"
'            .Cells(104, 1).Value = "declara bajo juramento que para el cálculo de las" _
'            & " retenciones relativas al período fiscal"
'            .Range("Q104:T104").Select
'            With o_Excel.Selection
'                .MergeCells = True
'                .WrapText = True
'                .Font.Size = 12
'                .Interior.ColorIndex = 15
'                .Interior.Pattern = xlSolid
'                .Font.Bold = True
'                .HorizontalAlignment = xlVAlignCenter
'                .VerticalAlignment = xlVAlignCenter
'            End With
'            .Cells(105, 1).Value = "han sido consideradas las normas legales, reglamentarias" _
'            & " y complementarias vigentes."
'            .Cells(102, 21).Value = "Lugar y fecha:"
'            .Range("U102:AC102").Select
'            With o_Excel.Selection
'                .MergeCells = True
'                .WrapText = True
'                .HorizontalAlignment = xlLeft
'                .VerticalAlignment = xlVAlignCenter
'            End With
'            .Cells(103, 21).Value = ""
'            .Range("U103:AC103").Select
'            With o_Excel.Selection
'                .MergeCells = True
'                .Font.Size = 12
'                .Font.Bold = True
'                .Interior.ColorIndex = 15
'                .Interior.Pattern = xlSolid
'                .HorizontalAlignment = xlLeft
'                .VerticalAlignment = xlVAlignCenter
'            End With
'            .Cells(104, 21).Value = "Firma y sello del agente de renteción:"
'            .Range("U104:AC104").Select
'            With o_Excel.Selection
'                .MergeCells = True
'                .WrapText = True
'                .HorizontalAlignment = xlLeft
'                .VerticalAlignment = xlVAlignCenter
'            End With
'            .Range("U105:AC108").Select
'            With o_Excel.Selection
'                .Interior.ColorIndex = 15
'                .Interior.Pattern = xlSolid
'            End With
'            .Range("U102:AC108").Select
'            With o_Excel.Selection.Borders(xlEdgeLeft)
'                .LineStyle = xlContinuous
'                .Weight = xlMedium
'                .ColorIndex = xlAutomatic
'            End With
'            With o_Excel.Selection.Borders(xlEdgeTop)
'                .LineStyle = xlContinuous
'                .Weight = xlMedium
'                .ColorIndex = xlAutomatic
'            End With
'            With o_Excel.Selection.Borders(xlEdgeBottom)
'                .LineStyle = xlContinuous
'                .Weight = xlMedium
'                .ColorIndex = xlAutomatic
'            End With
'            With o_Excel.Selection.Borders(xlEdgeRight)
'                .LineStyle = xlContinuous
'                .Weight = xlMedium
'                .ColorIndex = xlAutomatic
'            End With
'            .Range("A102:AC108").Select
'            With o_Excel.Selection.Borders(xlEdgeLeft)
'                .LineStyle = xlContinuous
'                .Weight = xlMedium
'                .ColorIndex = xlAutomatic
'            End With
'            With o_Excel.Selection.Borders(xlEdgeTop)
'                .LineStyle = xlContinuous
'                .Weight = xlMedium
'                .ColorIndex = xlAutomatic
'            End With
'            With o_Excel.Selection.Borders(xlEdgeBottom)
'                .LineStyle = xlContinuous
'                .Weight = xlMedium
'                .ColorIndex = xlAutomatic
'            End With
'            With o_Excel.Selection.Borders(xlEdgeRight)
'                .LineStyle = xlContinuous
'                .Weight = xlMedium
'                .ColorIndex = xlAutomatic
'            End With
'
'
'            .Range("U104:AC107").Select
'            o_Excel.Selection.Copy
'            .Range("U110:AC113").Select
'            .Paste
'            .Range("U106:AC108").Select
'            o_Excel.Selection.Copy
'            .Range("U114:AC116").Select
'            .Paste
'            .Cells(110, 1).Value = "A los efectos de cumplimentar lo dispuesto por el" _
'            & " artículo 6 de la Resolución General Nro."
'            .Range("R110:T110").Select
'            With o_Excel.Selection
'                .MergeCells = True
'                .WrapText = True
'                .Font.Size = 12
'                .Interior.ColorIndex = 15
'                .Interior.Pattern = xlSolid
'                .Font.Bold = True
'                .HorizontalAlignment = xlVAlignCenter
'                .VerticalAlignment = xlVAlignCenter
'            End With
'            .Cells(111, 1).Value = "día"
'            .Range("B111:C111").Select
'            With o_Excel.Selection
'                .MergeCells = True
'                .WrapText = True
'                .Font.Size = 12
'                .Interior.ColorIndex = 15
'                .Interior.Pattern = xlSolid
'                .Font.Bold = True
'                .HorizontalAlignment = xlVAlignCenter
'                .VerticalAlignment = xlVAlignCenter
'            End With
'            .Cells(111, 4).Value = "del mes"
'            .Range("F111:K111").Select
'            With o_Excel.Selection
'                .MergeCells = True
'                .WrapText = True
'                .Font.Size = 12
'                .Interior.ColorIndex = 15
'                .Interior.Pattern = xlSolid
'                .Font.Bold = True
'                .HorizontalAlignment = xlVAlignCenter
'                .VerticalAlignment = xlVAlignCenter
'            End With
'            .Cells(111, 12).Value = "de"
'            .Range("M111:P111").Select
'            With o_Excel.Selection
'                .MergeCells = True
'                .WrapText = True
'                .Font.Size = 12
'                .Interior.ColorIndex = 15
'                .Interior.Pattern = xlSolid
'                .Font.Bold = True
'                .HorizontalAlignment = xlVAlignCenter
'                .VerticalAlignment = xlVAlignCenter
'            End With
'            .Cells(111, 17).Value = "reintegraré al agente"
'            .Cells(112, 1).Value = "de retención el original y una copia (3) debidamente suscriptos."
'            .Cells(110, 21).Value = "Firma del beneficiario:"
'            .Range("A110:AC116").Select
'            With o_Excel.Selection.Borders(xlEdgeLeft)
'                .LineStyle = xlContinuous
'                .Weight = xlMedium
'                .ColorIndex = xlAutomatic
'            End With
'            With o_Excel.Selection.Borders(xlEdgeTop)
'                .LineStyle = xlContinuous
'                .Weight = xlMedium
'                .ColorIndex = xlAutomatic
'            End With
'            With o_Excel.Selection.Borders(xlEdgeBottom)
'                .LineStyle = xlContinuous
'                .Weight = xlMedium
'                .ColorIndex = xlAutomatic
'            End With
'            With o_Excel.Selection.Borders(xlEdgeRight)
'                .LineStyle = xlContinuous
'                .Weight = xlMedium
'                .ColorIndex = xlAutomatic
'            End With
'
'            .Cells(118, 1).Value = "Declaro que los datos consignados en este formulario" _
'            & " son correctos y completos y que he confeccionado la presente sin omitir" _
'            & " ni falsear dato alguno que deba contener, siendo fiel expresión de la verdad."
'            .Range("A118:T119").Select
'            With o_Excel.Selection
'                .MergeCells = True
'                .WrapText = True
'                .HorizontalAlignment = xlLeft
'                .VerticalAlignment = xlVAlignCenter
'            End With
'            .Range("A118:T124").Select
'            With o_Excel.Selection.Borders(xlEdgeLeft)
'                .LineStyle = xlContinuous
'                .Weight = xlMedium
'                .ColorIndex = xlAutomatic
'            End With
'            With o_Excel.Selection.Borders(xlEdgeTop)
'                .LineStyle = xlContinuous
'                .Weight = xlMedium
'                .ColorIndex = xlAutomatic
'            End With
'            With o_Excel.Selection.Borders(xlEdgeBottom)
'                .LineStyle = xlContinuous
'                .Weight = xlMedium
'                .ColorIndex = xlAutomatic
'            End With
'            With o_Excel.Selection.Borders(xlEdgeRight)
'                .LineStyle = xlContinuous
'                .Weight = xlMedium
'                .ColorIndex = xlAutomatic
'            End With
'            .Range("U102:AC108").Select
'            o_Excel.Selection.Copy
'            .Range("U118:AC124").Select
'            .Paste
'            .Cells(120, 21).Value = "Firma del beneficiario:"
'
'            .Cells(126, 1).Value = "(1) Testar lo que no corresponda."
'            .Cells(127, 1).Value = "(2) Presidente, gerente u otro responsable."
'            .Cells(128, 1).Value = "(3) Testar cuando no corresponda."
'            .Cells(129, 1).Value = "(4) Marcar con x el cuadro que corresponda."
'
'            'Llenamos los Datos Personales del Formulario
'            .Cells(1, 16).Value = LiquidacionFinalGanancias.txtPeriodo.Text
'            .Cells(1, 22).Value = "X"
'            Set rstBuscarSlave = New ADODB.Recordset
'            SQL = "Select * From AGENTES Where PUESTOLABORAL = '" & LiquidacionFinalGanancias.txtPuestoLaboral.Text & "'"
'            rstBuscarSlave.Open SQL, dbSlave, adOpenForwardOnly, adLockOptimistic
'            .Cells(4, 22).Value = Format(rstBuscarSlave!CUIL, "00-00000000-0")
'            .Cells(6, 14).Value = rstBuscarSlave!NombreCompleto
'            .Cells(10, 14).Value = "CORRIENTES (CAPITAL)"
'            .Cells(10, 21).Value = "CORRIENTES"
'            .Cells(10, 28).Value = "3400"
'            rstBuscarSlave.Close
'            Set rstBuscarSlave = Nothing
'            .Cells(17, 1).Value = "INSTITUTO DE VIVIENDA DE CORRIENTES"
'            .Cells(17, 14).Value = "30-63235151-4"
'            .Cells(17, 26).Interior.ColorIndex = 1
'            If LiquidacionFinalGanancias.optLiquidacionAnual.Value = True Then
'                .Cells(25, 6).Interior.ColorIndex = 1
'                .Cells(25, 14).Value = "01/01/" & LiquidacionFinalGanancias.txtPeriodo.Text
'                .Cells(25, 18).Value = "31/12/" & LiquidacionFinalGanancias.txtPeriodo.Text
'            Else
'                .Cells(25, 4).Interior.ColorIndex = 1
'                Set rstRegistroSlave = New ADODB.Recordset
'                SQL = "Select * From CODIGOLIQUIDACIONES Inner Join LIQUIDACIONSUELDOS On CODIGOLIQUIDACIONES.Codigo = LIQUIDACIONSUELDOS.CodigoLiquidacion " _
'                & "Where PUESTOLABORAL = '" & LiquidacionFinalGanancias.txtPuestoLaboral.Text & "' And Right(PERIODO,4) = '" & LiquidacionFinalGanancias.txtPeriodo.Text & "' " _
'                & "Order by CODIGOLIQUIDACION Asc"
'                rstRegistroSlave.Open SQL, dbSlave, adOpenForwardOnly, adLockOptimistic
'                rstRegistroSlave.MoveFirst
'                .Cells(25, 14).Value = "01/" & rstRegistroSlave!PERIODO
'                rstRegistroSlave.MoveLast
'                datFechaFin = DateSerial(Right(rstRegistroSlave!PERIODO, 4), Left(rstRegistroSlave!PERIODO, 2), 1)
'                datFechaFin = DateAdd("m", 1, datFecha)
'                datFechaFin = DateAdd("d", -1, datFecha)
'                .Cells(25, 18).Value = Day(datFecha) & "/" & rstRegistroSlave!PERIODO
'                rstRegistroSlave.Close
'                Set rstRegistroSlave = Nothing
'            End If
'            'POR EL MOMENTO, SE CARGA EN BASE A LO YA IMPUTADO
'            Set rstRegistroSlave = New ADODB.Recordset
'            'Completamos importe MontoBruto
'            SQL = "Select Sum(LIQUIDACIONGANANCIAS4TACATEGORIA.HaberOptimo) As SumaImporte " _
'            & "From CODIGOLIQUIDACIONES INNER JOIN LIQUIDACIONGANANCIAS4TACATEGORIA ON CODIGOLIQUIDACIONES.Codigo = LIQUIDACIONGANANCIAS4TACATEGORIA.CodigoLiquidacion " _
'            & "Where PUESTOLABORAL = '" & LiquidacionFinalGanancias.txtPuestoLaboral.Text & "' And Right(PERIODO,4) = '" & LiquidacionFinalGanancias.txtPeriodo.Text & "'"
'            rstRegistroSlave.Open SQL, dbSlave, adOpenForwardOnly, adLockOptimistic
'            If rstRegistroSlave!SumaImporte > 0 Then
'                .Cells(28, 26).Value = rstRegistroSlave!SumaImporte
'            Else
'                .Cells(28, 26).Value = 0
'            End If
'            rstRegistroSlave.Close
'            'Completamos importe Pluriempleo
'            SQL = "Select Sum(LIQUIDACIONGANANCIAS4TACATEGORIA.Pluriempleo) As SumaImporte " _
'            & "From CODIGOLIQUIDACIONES INNER JOIN LIQUIDACIONGANANCIAS4TACATEGORIA ON CODIGOLIQUIDACIONES.Codigo = LIQUIDACIONGANANCIAS4TACATEGORIA.CodigoLiquidacion " _
'            & "Where PUESTOLABORAL = '" & LiquidacionFinalGanancias.txtPuestoLaboral.Text & "' And Right(PERIODO,4) = '" & LiquidacionFinalGanancias.txtPeriodo.Text & "'"
'            rstRegistroSlave.Open SQL, dbSlave, adOpenForwardOnly, adLockOptimistic
'            If rstRegistroSlave!SumaImporte > 0 Then
'                .Cells(31, 26).Value = rstRegistroSlave!SumaImporte
'            Else
'                .Cells(31, 26).Value = 0
'            End If
'            rstRegistroSlave.Close
'            'Calculamos Total del Rubro 1
'            .Cells(35, 26).Value = .Cells(28, 26).Value + .Cells(31, 26).Value
'            'Completamos importe Aporte Jubilatorio
'            SQL = "Select Sum(LIQUIDACIONGANANCIAS4TACATEGORIA.Jubilacion) As SumaImporte " _
'            & "From CODIGOLIQUIDACIONES INNER JOIN LIQUIDACIONGANANCIAS4TACATEGORIA ON CODIGOLIQUIDACIONES.Codigo = LIQUIDACIONGANANCIAS4TACATEGORIA.CodigoLiquidacion " _
'            & "Where PUESTOLABORAL = '" & LiquidacionFinalGanancias.txtPuestoLaboral.Text & "' And Right(PERIODO,4) = '" & LiquidacionFinalGanancias.txtPeriodo.Text & "'"
'            rstRegistroSlave.Open SQL, dbSlave, adOpenForwardOnly, adLockOptimistic
'            If rstRegistroSlave!SumaImporte > 0 Then
'                .Cells(37, 26).Value = rstRegistroSlave!SumaImporte
'            Else
'                .Cells(37, 26).Value = 0
'            End If
'            rstRegistroSlave.Close
'            'Completamos importe Obra Social, Adherente Obra Social y Cuota Medico Asistencial
'            SQL = "Select Sum(LIQUIDACIONGANANCIAS4TACATEGORIA.ObraSocial + LIQUIDACIONGANANCIAS4TACATEGORIA.AdherenteObraSocial + LIQUIDACIONGANANCIAS4TACATEGORIA.CuotaMedicoAsistencial) As SumaImporte " _
'            & "From CODIGOLIQUIDACIONES INNER JOIN LIQUIDACIONGANANCIAS4TACATEGORIA ON CODIGOLIQUIDACIONES.Codigo = LIQUIDACIONGANANCIAS4TACATEGORIA.CodigoLiquidacion " _
'            & "Where PUESTOLABORAL = '" & LiquidacionFinalGanancias.txtPuestoLaboral.Text & "' And Right(PERIODO,4) = '" & LiquidacionFinalGanancias.txtPeriodo.Text & "'"
'            rstRegistroSlave.Open SQL, dbSlave, adOpenForwardOnly, adLockOptimistic
'            If rstRegistroSlave!SumaImporte > 0 Then
'                .Cells(38, 26).Value = rstRegistroSlave!SumaImporte
'            Else
'                .Cells(38, 26).Value = 0
'            End If
'            rstRegistroSlave.Close
'            'Completamos importe Cuota Medico Asistencial (REVERSO)
'            SQL = "Select Sum(LIQUIDACIONGANANCIAS4TACATEGORIA.CuotaMedicoAsistencial) As SumaImporte " _
'            & "From CODIGOLIQUIDACIONES INNER JOIN LIQUIDACIONGANANCIAS4TACATEGORIA ON CODIGOLIQUIDACIONES.Codigo = LIQUIDACIONGANANCIAS4TACATEGORIA.CodigoLiquidacion " _
'            & "Where PUESTOLABORAL = '" & LiquidacionFinalGanancias.txtPuestoLaboral.Text & "' And Right(PERIODO,4) = '" & LiquidacionFinalGanancias.txtPeriodo.Text & "'"
'            rstRegistroSlave.Open SQL, dbSlave, adOpenForwardOnly, adLockOptimistic
'            If rstRegistroSlave!SumaImporte > 0 Then
'                .Cells(69, 26).Value = rstRegistroSlave!SumaImporte
'                .Cells(71, 26).Value = rstRegistroSlave!SumaImporte
'            Else
'                .Cells(69, 26).Value = 0
'                .Cells(71, 26).Value = 0
'            End If
'            rstRegistroSlave.Close
'            'Completamos importe Seguro de Vida Optativo
'            SQL = "Select Sum(LIQUIDACIONGANANCIAS4TACATEGORIA.SeguroDeVidaOptativo) As SumaImporte " _
'            & "From CODIGOLIQUIDACIONES INNER JOIN LIQUIDACIONGANANCIAS4TACATEGORIA ON CODIGOLIQUIDACIONES.Codigo = LIQUIDACIONGANANCIAS4TACATEGORIA.CodigoLiquidacion " _
'            & "Where PUESTOLABORAL = '" & LiquidacionFinalGanancias.txtPuestoLaboral.Text & "' And Right(PERIODO,4) = '" & LiquidacionFinalGanancias.txtPeriodo.Text & "'"
'            rstRegistroSlave.Open SQL, dbSlave, adOpenForwardOnly, adLockOptimistic
'            If rstRegistroSlave!SumaImporte > 0 Then
'                .Cells(39, 26).Value = rstRegistroSlave!SumaImporte
'            Else
'                .Cells(39, 26).Value = 0
'            End If
'            rstRegistroSlave.Close
'            'Gasto de Sepelio (NO TENIDO EN CUENTA)
'            .Cells(40, 26).Value = 0
'            .Cells(78, 26).Value = 0
'            .Cells(80, 26).Value = 0
'            'Gastos estimativos de corredores y viajantes de comercio (NO TENIDO EN CUENTA)
'            .Cells(41, 26).Value = 0
'            'Completamos importe Otras Deducciones
'            SQL = "Select Sum(LIQUIDACIONGANANCIAS4TACATEGORIA.CuotaSindical + LIQUIDACIONGANANCIAS4TACATEGORIA.ServicioDomestico + LIQUIDACIONGANANCIAS4TACATEGORIA.SeguroDeVidaObligatorio + LIQUIDACIONGANANCIAS4TACATEGORIA.HonorariosMedicos) As SumaImporte " _
'            & "From CODIGOLIQUIDACIONES INNER JOIN LIQUIDACIONGANANCIAS4TACATEGORIA ON CODIGOLIQUIDACIONES.Codigo = LIQUIDACIONGANANCIAS4TACATEGORIA.CodigoLiquidacion " _
'            & "Where PUESTOLABORAL = '" & LiquidacionFinalGanancias.txtPuestoLaboral.Text & "' And Right(PERIODO,4) = '" & LiquidacionFinalGanancias.txtPeriodo.Text & "'"
'            rstRegistroSlave.Open SQL, dbSlave, adOpenForwardOnly, adLockOptimistic
'            If rstRegistroSlave!SumaImporte > 0 Then
'                .Cells(42, 26).Value = rstRegistroSlave!SumaImporte
'            Else
'                .Cells(42, 26).Value = 0
'            End If
'            rstRegistroSlave.Close
'            'Completamos importe Seguro de Vida Obligatorio
'            SQL = "Select Sum(LIQUIDACIONGANANCIAS4TACATEGORIA.SeguroDeVidaObligatorio) As SumaImporte " _
'            & "From CODIGOLIQUIDACIONES INNER JOIN LIQUIDACIONGANANCIAS4TACATEGORIA ON CODIGOLIQUIDACIONES.Codigo = LIQUIDACIONGANANCIAS4TACATEGORIA.CodigoLiquidacion " _
'            & "Where PUESTOLABORAL = '" & LiquidacionFinalGanancias.txtPuestoLaboral.Text & "' And Right(PERIODO,4) = '" & LiquidacionFinalGanancias.txtPeriodo.Text & "'"
'            rstRegistroSlave.Open SQL, dbSlave, adOpenForwardOnly, adLockOptimistic
'            If rstRegistroSlave!SumaImporte > 0 Then
'                .Cells(88, 3).Value = "Decreto Ley Nª 30/2000 - Seg.Vida Oblig."
'                .Cells(88, 21).Value = rstRegistroSlave!SumaImporte
'                .Cells(88, 26).Value = rstRegistroSlave!SumaImporte
'                .Cells(91, 26).Value = rstRegistroSlave!SumaImporte
'            End If
'            rstRegistroSlave.Close
'            'Cargamos Importe Cuota Sindical
'            SQL = "Select Sum(LIQUIDACIONGANANCIAS4TACATEGORIA.CuotaSindical) As SumaImporte " _
'            & "From CODIGOLIQUIDACIONES INNER JOIN LIQUIDACIONGANANCIAS4TACATEGORIA ON CODIGOLIQUIDACIONES.Codigo = LIQUIDACIONGANANCIAS4TACATEGORIA.CodigoLiquidacion " _
'            & "Where PUESTOLABORAL = '" & LiquidacionFinalGanancias.txtPuestoLaboral.Text & "' And Right(PERIODO,4) = '" & LiquidacionFinalGanancias.txtPeriodo.Text & "'"
'            rstRegistroSlave.Open SQL, dbSlave, adOpenForwardOnly, adLockOptimistic
'            If rstRegistroSlave!SumaImporte > 0 Then
'                If Trim(.Cells(88, 26).Value) = "" Then
'                    .Cells(88, 3).Value = "Ley 23551 Art. 37 - Cuota Sindical"
'                    .Cells(88, 21).Value = rstRegistroSlave!SumaImporte
'                    .Cells(88, 26).Value = rstRegistroSlave!SumaImporte
'                    .Cells(91, 26).Value = rstRegistroSlave!SumaImporte
'                Else
'                    .Cells(89, 3).Value = "Ley 23551 Art. 37 - Cuota Sindical"
'                    .Cells(89, 21).Value = rstRegistroSlave!SumaImporte
'                    .Cells(89, 26).Value = rstRegistroSlave!SumaImporte
'                    .Cells(91, 26).Value = rstRegistroSlave!SumaImporte + .Cells(91, 26).Value
'                End If
'            End If
'            rstRegistroSlave.Close
'            'Cargamos Importe ServicioDomestico
'            SQL = "Select Sum(LIQUIDACIONGANANCIAS4TACATEGORIA.ServicioDomestico) As SumaImporte " _
'            & "From CODIGOLIQUIDACIONES INNER JOIN LIQUIDACIONGANANCIAS4TACATEGORIA ON CODIGOLIQUIDACIONES.Codigo = LIQUIDACIONGANANCIAS4TACATEGORIA.CodigoLiquidacion " _
'            & "Where PUESTOLABORAL = '" & LiquidacionFinalGanancias.txtPuestoLaboral.Text & "' And Right(PERIODO,4) = '" & LiquidacionFinalGanancias.txtPeriodo.Text & "'"
'            rstRegistroSlave.Open SQL, dbSlave, adOpenForwardOnly, adLockOptimistic
'            If rstRegistroSlave!SumaImporte > 0 Then
'                If Trim(.Cells(89, 26).Value) = "" Then
'                    If Trim(.Cells(88, 26).Value) = "" Then
'                        .Cells(88, 3).Value = "Ley 26063 Art. 16 - Servicio Doméstico"
'                        .Cells(88, 21).Value = rstRegistroSlave!SumaImporte
'                        .Cells(88, 26).Value = rstRegistroSlave!SumaImporte
'                        .Cells(91, 26).Value = rstRegistroSlave!SumaImporte
'                    Else
'                        .Cells(89, 3).Value = "Ley 26063 Art. 16 - Servicio Doméstico"
'                        .Cells(89, 21).Value = rstRegistroSlave!SumaImporte
'                        .Cells(89, 26).Value = rstRegistroSlave!SumaImporte
'                        .Cells(91, 26).Value = rstRegistroSlave!SumaImporte + .Cells(91, 26).Value
'                    End If
'                Else
'                    .Cells(90, 3).Value = "Ley 26063 Art. 16 - Servicio Doméstico"
'                    .Cells(90, 21).Value = rstRegistroSlave!SumaImporte
'                    .Cells(90, 26).Value = rstRegistroSlave!SumaImporte
'                    .Cells(91, 26).Value = rstRegistroSlave!SumaImporte + .Cells(91, 26).Value
'                End If
'            End If
'            rstRegistroSlave.Close
'            'Cargamos Importe Honorarios Médicos
'            SQL = "Select Sum(LIQUIDACIONGANANCIAS4TACATEGORIA.HonorariosMedicos) As SumaImporte " _
'            & "From CODIGOLIQUIDACIONES INNER JOIN LIQUIDACIONGANANCIAS4TACATEGORIA ON CODIGOLIQUIDACIONES.Codigo = LIQUIDACIONGANANCIAS4TACATEGORIA.CodigoLiquidacion " _
'            & "Where PUESTOLABORAL = '" & LiquidacionFinalGanancias.txtPuestoLaboral.Text & "' And Right(PERIODO,4) = '" & LiquidacionFinalGanancias.txtPeriodo.Text & "'"
'            rstRegistroSlave.Open SQL, dbSlave, adOpenForwardOnly, adLockOptimistic
'            If rstRegistroSlave!SumaImporte > 0 Then
'                If Trim(.Cells(89, 26).Value) = "" Then
'                    If Trim(.Cells(88, 26).Value) = "" Then
'                        .Cells(88, 3).Value = "Honorarios Médicos"
'                        .Cells(88, 21).Value = rstRegistroSlave!SumaImporte
'                        .Cells(88, 26).Value = rstRegistroSlave!SumaImporte
'                    Else
'                        .Cells(89, 3).Value = "Honorarios Médicos"
'                        .Cells(89, 21).Value = rstRegistroSlave!SumaImporte
'                        .Cells(89, 26).Value = rstRegistroSlave!SumaImporte
'                        .Cells(91, 26).Value = rstRegistroSlave!SumaImporte + .Cells(91, 26).Value
'                    End If
'                Else
'                    .Cells(90, 3).Value = "Honorarios Médicos"
'                    .Cells(90, 21).Value = rstRegistroSlave!SumaImporte
'                    .Cells(90, 26).Value = rstRegistroSlave!SumaImporte
'                    .Cells(91, 26).Value = rstRegistroSlave!SumaImporte + .Cells(91, 26).Value
'                End If
'            End If
'            rstRegistroSlave.Close
'            'Calculamos Total del Rubro 2
'            .Cells(43, 26).Value = .Cells(37, 26).Value + .Cells(38, 26).Value + .Cells(39, 26).Value _
'            + .Cells(40, 26).Value + .Cells(41, 26).Value + .Cells(42, 26).Value
'            'Calculamos Resultado Neto
'            .Cells(44, 26).Value = .Cells(35, 26).Value - .Cells(43, 26).Value
'            'Completamos importe Donaciones
'            SQL = "Select Sum(LIQUIDACIONGANANCIAS4TACATEGORIA.Donaciones) As SumaImporte " _
'            & "From CODIGOLIQUIDACIONES INNER JOIN LIQUIDACIONGANANCIAS4TACATEGORIA ON CODIGOLIQUIDACIONES.Codigo = LIQUIDACIONGANANCIAS4TACATEGORIA.CodigoLiquidacion " _
'            & "Where PUESTOLABORAL = '" & LiquidacionFinalGanancias.txtPuestoLaboral.Text & "' And Right(PERIODO,4) = '" & LiquidacionFinalGanancias.txtPeriodo.Text & "'"
'            rstRegistroSlave.Open SQL, dbSlave, adOpenForwardOnly, adLockOptimistic
'            If rstRegistroSlave!SumaImporte > 0 Then
'                .Cells(45, 26).Value = rstRegistroSlave!SumaImporte
'                .Cells(83, 26).Value = rstRegistroSlave!SumaImporte
'                .Cells(85, 26).Value = rstRegistroSlave!SumaImporte
'            Else
'                .Cells(45, 26).Value = 0
'                .Cells(83, 26).Value = 0
'                .Cells(85, 26).Value = 0
'            End If
'            rstRegistroSlave.Close
'            'Calculamos Diferencia Rubro 3 - Rubro 4
'            .Cells(46, 26).Value = .Cells(44, 26).Value - .Cells(45, 26).Value
'            'Completamos importe Deducción Especial
'            SQL = "Select Sum(LIQUIDACIONGANANCIAS4TACATEGORIA.DeduccionEspecial) As SumaImporte " _
'            & "From CODIGOLIQUIDACIONES INNER JOIN LIQUIDACIONGANANCIAS4TACATEGORIA ON CODIGOLIQUIDACIONES.Codigo = LIQUIDACIONGANANCIAS4TACATEGORIA.CodigoLiquidacion " _
'            & "Where PUESTOLABORAL = '" & LiquidacionFinalGanancias.txtPuestoLaboral.Text & "' And Right(PERIODO,4) = '" & LiquidacionFinalGanancias.txtPeriodo.Text & "'"
'            rstRegistroSlave.Open SQL, dbSlave, adOpenForwardOnly, adLockOptimistic
'            If rstRegistroSlave!SumaImporte > 0 Then
'                .Cells(48, 26).Value = rstRegistroSlave!SumaImporte
'            Else
'                .Cells(48, 26).Value = 0
'            End If
'            rstRegistroSlave.Close
'            'Completamos importe Minimo no Imponible
'            SQL = "Select Sum(LIQUIDACIONGANANCIAS4TACATEGORIA.MinimoNoImponible) As SumaImporte " _
'            & "From CODIGOLIQUIDACIONES INNER JOIN LIQUIDACIONGANANCIAS4TACATEGORIA ON CODIGOLIQUIDACIONES.Codigo = LIQUIDACIONGANANCIAS4TACATEGORIA.CodigoLiquidacion " _
'            & "Where PUESTOLABORAL = '" & LiquidacionFinalGanancias.txtPuestoLaboral.Text & "' And Right(PERIODO,4) = '" & LiquidacionFinalGanancias.txtPeriodo.Text & "'"
'            rstRegistroSlave.Open SQL, dbSlave, adOpenForwardOnly, adLockOptimistic
'            If rstRegistroSlave!SumaImporte > 0 Then
'                .Cells(49, 26).Value = rstRegistroSlave!SumaImporte
'            Else
'                .Cells(49, 26).Value = 0
'            End If
'            rstRegistroSlave.Close
'            'Completamos importe Conyuge
'            SQL = "Select Sum(LIQUIDACIONGANANCIAS4TACATEGORIA.Conyuge) As SumaImporte " _
'            & "From CODIGOLIQUIDACIONES INNER JOIN LIQUIDACIONGANANCIAS4TACATEGORIA ON CODIGOLIQUIDACIONES.Codigo = LIQUIDACIONGANANCIAS4TACATEGORIA.CodigoLiquidacion " _
'            & "Where PUESTOLABORAL = '" & LiquidacionFinalGanancias.txtPuestoLaboral.Text & "' And Right(PERIODO,4) = '" & LiquidacionFinalGanancias.txtPeriodo.Text & "'"
'            rstRegistroSlave.Open SQL, dbSlave, adOpenForwardOnly, adLockOptimistic
'            If rstRegistroSlave!SumaImporte > 0 Then
'                .Cells(51, 26).Value = rstRegistroSlave!SumaImporte
'            Else
'                .Cells(51, 26).Value = 0
'            End If
'            rstRegistroSlave.Close
'            'Completamos importe Hijos
'            SQL = "Select Sum(LIQUIDACIONGANANCIAS4TACATEGORIA.Hijo) As SumaImporte " _
'            & "From CODIGOLIQUIDACIONES INNER JOIN LIQUIDACIONGANANCIAS4TACATEGORIA ON CODIGOLIQUIDACIONES.Codigo = LIQUIDACIONGANANCIAS4TACATEGORIA.CodigoLiquidacion " _
'            & "Where PUESTOLABORAL = '" & LiquidacionFinalGanancias.txtPuestoLaboral.Text & "' And Right(PERIODO,4) = '" & LiquidacionFinalGanancias.txtPeriodo.Text & "'"
'            rstRegistroSlave.Open SQL, dbSlave, adOpenForwardOnly, adLockOptimistic
'            If rstRegistroSlave!SumaImporte > 0 Then
'                .Cells(52, 26).Value = rstRegistroSlave!SumaImporte
'            Else
'                .Cells(52, 26).Value = 0
'            End If
'            rstRegistroSlave.Close
'            'Completamos importe Otras Cargas de Familia
'            SQL = "Select Sum(LIQUIDACIONGANANCIAS4TACATEGORIA.OtrasCargasDeFamilia) As SumaImporte " _
'            & "From CODIGOLIQUIDACIONES INNER JOIN LIQUIDACIONGANANCIAS4TACATEGORIA ON CODIGOLIQUIDACIONES.Codigo = LIQUIDACIONGANANCIAS4TACATEGORIA.CodigoLiquidacion " _
'            & "Where PUESTOLABORAL = '" & LiquidacionFinalGanancias.txtPuestoLaboral.Text & "' And Right(PERIODO,4) = '" & LiquidacionFinalGanancias.txtPeriodo.Text & "'"
'            rstRegistroSlave.Open SQL, dbSlave, adOpenForwardOnly, adLockOptimistic
'            If rstRegistroSlave!SumaImporte > 0 Then
'                .Cells(53, 26).Value = rstRegistroSlave!SumaImporte
'            Else
'                .Cells(53, 26).Value = 0
'            End If
'            rstRegistroSlave.Close
'            'Completamos importe Total Deducciones Personales
'            .Cells(54, 26).Value = .Cells(48, 26).Value + .Cells(49, 26).Value + .Cells(51, 26).Value + .Cells(52, 26).Value + .Cells(53, 26).Value
'            'Calculamos Ganancia Neta Sujeta a Impuesto
'            .Cells(55, 26).Value = .Cells(46, 26).Value - .Cells(54, 26).Value
'            'Calculamos Impuesto Determinado
'            If .Cells(55, 26).Value < 0 Then
'                .Cells(56, 26).Value = 0
'            Else
'                SQL = "Select * From ESCALAAPLICABLEGANANCIAS Order by IMPORTEMAXIMO Asc"
'                rstRegistroSlave.Open SQL, dbSlave, adOpenDynamic, adLockOptimistic
'                If rstRegistroSlave.BOF = False Then
'                    rstRegistroSlave.MoveFirst
'                    Do While rstRegistroSlave.EOF = False
'                        If rstRegistroSlave!ImporteMaximo > .Cells(55, 26).Value Then
'                            dblPorcentajeAplicable = rstRegistroSlave!ImporteVariable
'                            Exit Do
'                        End If
'                        rstRegistroSlave.MoveNext
'                    Loop
'                    If rstRegistroSlave!ImporteFijo = 0 Then
'                        .Cells(56, 26).Value = .Cells(55, 26).Value * rstRegistroSlave!ImporteVariable
'                    Else
'                        .Cells(56, 26).Value = rstRegistroSlave!ImporteFijo
'                        rstRegistroSlave.MovePrevious
'                        .Cells(56, 26).Value = .Cells(56, 26).Value + ((.Cells(55, 26).Value - rstRegistroSlave!ImporteMaximo) * dblPorcentajeAplicable)
'                    End If
'                End If
'                rstRegistroSlave.Close
'            End If
'            'Completamos importe Retenciones Efectuadas
'            SQL = "Select Sum(LIQUIDACIONGANANCIAS4TACATEGORIA.Retencion) As SumaImporte " _
'            & "From CODIGOLIQUIDACIONES INNER JOIN LIQUIDACIONGANANCIAS4TACATEGORIA ON CODIGOLIQUIDACIONES.Codigo = LIQUIDACIONGANANCIAS4TACATEGORIA.CodigoLiquidacion " _
'            & "Where PUESTOLABORAL = '" & LiquidacionFinalGanancias.txtPuestoLaboral.Text & "' And Right(PERIODO,4) = '" & LiquidacionFinalGanancias.txtPeriodo.Text & "'"
'            rstRegistroSlave.Open SQL, dbSlave, adOpenForwardOnly, adLockOptimistic
'            If rstRegistroSlave!SumaImporte > 0 Then
'                .Cells(58, 26).Value = rstRegistroSlave!SumaImporte
'            Else
'                .Cells(58, 26).Value = 0
'            End If
'            rstRegistroSlave.Close
'            Set rstRegistroSlave = Nothing
'            'Régimenes de Promoción (NO TENIDO EN CUENTA)
'            .Cells(59, 26).Value = 0
'            'Calculamos Ganancia Neta Sujeta a Impuesto
'            .Cells(60, 26).Value = .Cells(58, 26).Value - .Cells(59, 26).Value
'            'Determinamos Saldo a Favor
'            Select Case .Cells(56, 26).Value - .Cells(60, 26).Value
'            Case Is = 0
'                .Cells(62, 26).Value = 0
'                .Cells(63, 26).Value = 0
'            Case Is < 0
'                .Cells(62, 26).Value = 0
'                .Cells(63, 26).Value = .Cells(60, 26).Value - .Cells(56, 26).Value
'                .Cells(64, 7).Value = .Cells(60, 26).Value - .Cells(56, 26).Value
'            Case Is > 0
'                .Cells(62, 26).Value = .Cells(56, 26).Value - .Cells(60, 26).Value
'                .Cells(63, 26).Value = 0
'                .Cells(64, 7).Value = .Cells(56, 26).Value - .Cells(60, 26).Value
'            End Select
'        End With
'        o_Libro.Close True, strOutputPath
'        ' -- Cerrar Excel
'        o_Excel.Quit
'        ' -- Terminar instancias
'        Call ReleaseObjects(o_Excel, o_Libro, o_Hoja)
'        GenerarF649 = True
'    End If
'
'Exit Function
'
'' -- Controlador de Errores
'Error_Handler:
'    ' -- Cierra la hoja y el la aplicación Excel
'    If Not o_Libro Is Nothing Then: o_Libro.Close False
'    If Not o_Excel Is Nothing Then: o_Excel.Quit
'    Call ReleaseObjects(o_Excel, o_Libro, o_Hoja)
'    If Err.Number <> 1004 Then MsgBox Err.Description, vbCritical
'End Function

Public Function GenerarF649(strOutputPath As String) As Boolean
    On Error GoTo Error_Handler
  
    Dim o_Excel                 As Object
    Dim o_Libro                 As Object
    Dim o_Hoja                  As Object
    Dim ColumnaExcel            As Integer
    Dim FilaExcel               As Integer
    Dim SQL                     As String
    Dim datFechaInicio          As Date
    Dim datFechaFin             As Date
    Dim dblImporteMensual       As Double
    Dim intMesesLiquidados      As Integer
    Dim intFamiliares           As Integer
    Dim i                       As Integer
    Dim dblImporteControl       As Double
    Dim dblGananciaNeta         As Double
    Dim dblPorcentajeAplicable  As Double
    
    Dim strCLF649                   As String
    Dim strPLF649                   As String
    Dim strYearF649                 As String
    Dim dblImporteCalculado         As Double
    Dim dblRemuneracionComputable   As Double
    Dim dblDeduccionesGenerales     As Double
    Dim dblDeduccionesPersonales    As Double
    If ValidarGenerarF649 = True Then
        ' -- Crea el objeto Excel, el objeto workBook y el objeto sheet
        Set o_Excel = CreateObject("Excel.Application")
        Set o_Libro = o_Excel.Workbooks.Add
        Set o_Hoja = o_Libro.Worksheets(1)
           
        
        'Configurando Para Impresión
'        With o_Libro.ActiveSheet.PageSetup
'            .LeftHeader = ""
'            .CenterHeader = ""
'            .RightHeader = ""
'            .LeftFooter = ""
'            .CenterFooter = ""
'            .RightFooter = ""
'            .LeftMargin = Application.InchesToPoints(0.196850393700787)
'            .RightMargin = Application.InchesToPoints(0.196850393700787)
'            .TopMargin = Application.InchesToPoints(0.78740157480315)
'            .BottomMargin = Application.InchesToPoints(0.78740157480315)
'            .HeaderMargin = Application.InchesToPoints(0)
'            .FooterMargin = Application.InchesToPoints(0)
'            .PrintHeadings = False
'            .PrintGridlines = False
'            .PrintComments = xlPrintNoComments
'            .PrintQuality = 600
'            .CenterHorizontally = True
'            .CenterVertically = False
'            .Orientation = xlPortrait
'            .Draft = False
'            .PaperSize = xlPaperA4
'            .FirstPageNumber = xlAutomatic
'            .Order = xlDownThenOver
'            .BlackAndWhite = False
'            .Zoom = 75
'            .PrintErrors = xlPrintErrorsDisplayed
'        End With
    
        
        With o_Hoja
            'Configurando Ancho de Columnas
            .Columns("A:AD").Select
            o_Excel.Selection.ColumnWidth = 3.57
            
            'Configurando Ancho de Filas
            .Rows("1:62").Select
            o_Excel.Selection.RowHeight = 14.25
            
            'Configurando Fuente
            .Range("A1:AC138").Select
            'o_Excel.Selection.Font.Type = "Time New Roman"
            o_Excel.Selection.Font.Size = 9
    
            'Configurar Formulario
            .Cells(1, 1).Value = "ANEXO VII RESOLUCIÓN GENERAL N° 2.437, SUS MODIFICATORIAS Y COMPLEMENTARIAS (Artículo 14)"
            .Cells(2, 1).Value = "LIQUIDACIÓN DE IMPUESTO A LAS GANANCIAS - 4ta. CATEGORÍA RELACIÓN DE DEPENDENCIA"
            .Cells(3, 1).Value = ""
            For i = 1 To 3
                .Range("A" & i & ":AD" & i).Select
                With o_Excel.Selection
                    .MergeCells = True
                    .WrapText = True
                    .Font.Bold = True
                    .HorizontalAlignment = xlVAlignCenter
                    .VerticalAlignment = xlVAlignCenter
                End With
            Next i
            
            .Cells(7, 1).Value = "DATOS DEL EMPLEADO"
            .Range("A7:AD7").Select
            With o_Excel.Selection
                .MergeCells = True
                .WrapText = True
                .Font.Bold = True
                .HorizontalAlignment = xlVAlignCenter
                .VerticalAlignment = xlVAlignCenter
            End With
            With o_Excel.Selection.Borders(xlEdgeLeft)
                .LineStyle = xlContinuous
                .Weight = xlMedium
                .ColorIndex = xlAutomatic
            End With
            With o_Excel.Selection.Borders(xlEdgeTop)
                .LineStyle = xlContinuous
                .Weight = xlMedium
                .ColorIndex = xlAutomatic
            End With
            With o_Excel.Selection.Borders(xlEdgeBottom)
                .LineStyle = xlContinuous
                .Weight = xlMedium
                .ColorIndex = xlAutomatic
            End With
            With o_Excel.Selection.Borders(xlEdgeRight)
                .LineStyle = xlContinuous
                .Weight = xlMedium
                .ColorIndex = xlAutomatic
            End With
    
            .Cells(8, 1).Value = "Apellido y Nombres:"
            .Range("A8:O8").Select
            With o_Excel.Selection
                .MergeCells = True
                .WrapText = True
                .HorizontalAlignment = xlLeft
                .VerticalAlignment = xlVAlignCenter
            End With
            .Cells(9, 1).Value = ""
            .Range("A9:O9").Select
            With o_Excel.Selection
                .MergeCells = True
                .WrapText = True
                .Font.Size = 12
                .Font.Bold = True
                .Interior.ColorIndex = 15
                .Interior.Pattern = xlSolid
                .HorizontalAlignment = xlGeneral
                .VerticalAlignment = xlTop
            End With
            .Range("A8:O9").Select
            With o_Excel.Selection.Borders(xlEdgeLeft)
                .LineStyle = xlContinuous
                .Weight = xlMedium
                .ColorIndex = xlAutomatic
            End With
            With o_Excel.Selection.Borders(xlEdgeTop)
                .LineStyle = xlContinuous
                .Weight = xlMedium
                .ColorIndex = xlAutomatic
            End With
            With o_Excel.Selection.Borders(xlEdgeBottom)
                .LineStyle = xlContinuous
                .Weight = xlMedium
                .ColorIndex = xlAutomatic
            End With
            With o_Excel.Selection.Borders(xlEdgeRight)
                .LineStyle = xlContinuous
                .Weight = xlMedium
                .ColorIndex = xlAutomatic
            End With
            
            .Range("A8:O9").Select
            o_Excel.Selection.Copy
            .Range("P8:AD9").Select
            .Paste
            .Cells(8, 16).Value = "Clave Única de Identificación Tributaria / Laboral:"
            
            .Range("A7:AD9").Select
            o_Excel.Selection.Copy
            .Range("A11:AD13").Select
            .Paste
            .Cells(11, 1).Value = "DATOS DEL AGENTE DE RETENCIÓN"
            .Cells(12, 1).Value = "Apellido y Nombres o Razón Social:"
            .Cells(12, 16).Value = "Clave Única de Identificación Tributaria:"
            
            .Cells(16, 1).Value = "Rub."
            .Cells(16, 1).Select
            With o_Excel.Selection
                .MergeCells = True
                .WrapText = True
                .Font.Bold = True
                .HorizontalAlignment = xlVAlignCenter
                .VerticalAlignment = xlVAlignCenter
            End With
            With o_Excel.Selection.Borders(xlEdgeLeft)
                .LineStyle = xlContinuous
                .Weight = xlMedium
                .ColorIndex = xlAutomatic
            End With
            With o_Excel.Selection.Borders(xlEdgeTop)
                .LineStyle = xlContinuous
                .Weight = xlMedium
                .ColorIndex = xlAutomatic
            End With
            With o_Excel.Selection.Borders(xlEdgeBottom)
                .LineStyle = xlContinuous
                .Weight = xlMedium
                .ColorIndex = xlAutomatic
            End With
            With o_Excel.Selection.Borders(xlEdgeRight)
                .LineStyle = xlContinuous
                .Weight = xlMedium
                .ColorIndex = xlAutomatic
            End With

            .Cells(16, 1).Select
            o_Excel.Selection.Copy
            .Cells(16, 2).Select
            .Paste
            .Cells(16, 2).Value = "Ins."

            .Cells(16, 3).Value = "DETERMINACION DE LA GANANCIA NETA Y LIQUIDACION DEL IMPUESTO"
            .Range("C16:Y16").Select
            With o_Excel.Selection
                .MergeCells = True
                .WrapText = True
                .Font.Bold = True
                .HorizontalAlignment = xlLeft
                .VerticalAlignment = xlVAlignCenter
            End With
            
            .Cells(16, 26).Value = "IMPORTES"
            .Range("Z16:AD16").Select
            With o_Excel.Selection
                .MergeCells = True
                .WrapText = True
                .Font.Bold = True
                .HorizontalAlignment = xlVAlignCenter
                .VerticalAlignment = xlVAlignCenter
            End With
            
            .Range("A16:AD16").Select
            With o_Excel.Selection.Borders(xlEdgeLeft)
                .LineStyle = xlContinuous
                .Weight = xlMedium
                .ColorIndex = xlAutomatic
            End With
            With o_Excel.Selection.Borders(xlEdgeTop)
                .LineStyle = xlContinuous
                .Weight = xlMedium
                .ColorIndex = xlAutomatic
            End With
            With o_Excel.Selection.Borders(xlEdgeBottom)
                .LineStyle = xlContinuous
                .Weight = xlMedium
                .ColorIndex = xlAutomatic
            End With
            With o_Excel.Selection.Borders(xlEdgeRight)
                .LineStyle = xlContinuous
                .Weight = xlMedium
                .ColorIndex = xlAutomatic
            End With

            .Range("A17:A57").Select
            With o_Excel.Selection
                .NumberFormat = "@"
                .HorizontalAlignment = xlVAlignCenter
                .VerticalAlignment = xlVAlignCenter
            End With
            With o_Excel.Selection.Borders(xlEdgeLeft)
                .LineStyle = xlContinuous
                .Weight = xlMedium
                .ColorIndex = xlAutomatic
            End With
            With o_Excel.Selection.Borders(xlEdgeTop)
                .LineStyle = xlContinuous
                .Weight = xlMedium
                .ColorIndex = xlAutomatic
            End With
            With o_Excel.Selection.Borders(xlEdgeBottom)
                .LineStyle = xlContinuous
                .Weight = xlMedium
                .ColorIndex = xlAutomatic
            End With
            With o_Excel.Selection.Borders(xlEdgeRight)
                .LineStyle = xlContinuous
                .Weight = xlMedium
                .ColorIndex = xlAutomatic
            End With

            .Range("A17:A57").Select
            o_Excel.Selection.Copy
            .Range("B17:B57").Select
            .Paste

            .Cells(17, 1).Value = "1"
            .Cells(26, 1).Value = "2"
            .Cells(45, 1).Value = "3"
            
            i = 53
            .Cells(i, 1).Value = "4"
            i = i + 1
            .Cells(i, 1).Value = "5"
            i = i + 1
            .Cells(i, 1).Value = "6"
            i = i + 1
            .Cells(i, 1).Value = "7"
            i = i + 1
            .Cells(i, 1).Value = "8"
            
            i = 18
            .Cells(i, 2).Value = "a"
            i = i + 1
            .Cells(i, 2).Value = "b"
            i = i + 1
            .Cells(i, 2).Value = "c"
            i = i + 1
            .Cells(i, 2).Value = "d"
            i = i + 1
            .Cells(i, 2).Value = "e"
            i = i + 1
            .Cells(i, 2).Value = "f"
            i = i + 1
            .Cells(i, 2).Value = "g"
            
            i = 27
            .Cells(i, 2).Value = "a"
            i = i + 1
            .Cells(i, 2).Value = "b"
            i = i + 1
            .Cells(i, 2).Value = "c"
            i = i + 1
            .Cells(i, 2).Value = "d"
            i = i + 1
            .Cells(i, 2).Value = "e"
            i = i + 1
            .Cells(i, 2).Value = "f"
            i = i + 1
            .Cells(i, 2).Value = "g"
            i = i + 1
            .Cells(i, 2).Value = "h"
            i = i + 1
            .Cells(i, 2).Value = "i"
            i = i + 1
            .Cells(i, 2).Value = "j"
            i = i + 1
            .Cells(i, 2).Value = "k"
            i = i + 1
            .Cells(i, 2).Value = "l"
            i = i + 1
            .Cells(i, 2).Value = "m"
            i = i + 1
            .Cells(i, 2).Value = "n"
            i = i + 1
            .Cells(i, 2).Value = "o"
            i = i + 1
            .Cells(i, 2).Value = "p"
            i = i + 1
            .Cells(i, 2).Value = "q"

            i = 46
            .Cells(i, 2).Value = "a"
            i = i + 1
            .Cells(i, 2).Value = "b"
            i = i + 2
            .Cells(i, 2).Value = "c"
            i = i + 1
            .Cells(i, 2).Value = "d"
            i = i + 1
            .Cells(i, 2).Value = "e"

            i = 17
            .Cells(i, 3).Value = "REMUNERACIONES"
            .Cells(i, 3).Font.Bold = True
            i = i + 1
            .Cells(i, 3).Value = "Remuneración Bruta"
            i = i + 1
            .Cells(i, 3).Value = "Retribuciones no Habituales"
            i = i + 1
            .Cells(i, 3).Value = "SAC Primera Cuota"
            i = i + 1
            .Cells(i, 3).Value = "SAC Segunda Cuota"
            i = i + 1
            .Cells(i, 3).Value = "Remuneración No Alcanzada"
            i = i + 1
            .Cells(i, 3).Value = "Remuneración Exenta"
            i = i + 1
            .Cells(i, 3).Value = "Remuneración Otros Empleos"
            i = i + 1
            .Cells(i, 3).Value = "REMUNERACIÓN COMPUTABLE (suma de los incisos a, b, c, d y g)"
            i = i + 1
            .Cells(i, 3).Value = "DEDUCCIONES GENERALES"
            .Cells(i, 3).Font.Bold = True
            i = i + 1
            .Cells(i, 3).Value = "Aportes a fondos de jubilaciones, retiros, pensiones o subsidios que se destinen a cajas nacionales, prov. o municipales."
            i = i + 1
            .Cells(i, 3).Value = "Aportes Obra Social"
            i = i + 1
            .Cells(i, 3).Value = "Cuota sindical"
            i = i + 1
            .Cells(i, 3).Value = "Aportes Jubilatorios Otros Empleos"
            i = i + 1
            .Cells(i, 3).Value = "Aportes Obra Social otros empleos"
            i = i + 1
            .Cells(i, 3).Value = "Cuota sindical otros empleos"
            i = i + 1
            .Cells(i, 3).Value = "Cuotas médico asistenciales"
            i = i + 1
            .Cells(i, 3).Value = "Primas de Seguro para el caso de muerte"
            i = i + 1
            .Cells(i, 3).Value = "Gastos de Sepelio"
            i = i + 1
            .Cells(i, 3).Value = "Gastos estimativos para corredores y viajantes de comercio "
            i = i + 1
            .Cells(i, 3).Value = "Donaciones a fiscos nacional, provinciales y municipales, y a instituciones comprendidas en el art. 20, inc. e) y f) de la ley"
            i = i + 1
            .Cells(i, 3).Value = "Descuentos obligatorios establecidos por ley nacional, provincial o municipal"
            i = i + 1
            .Cells(i, 3).Value = "Honorarios por servicios de asistencia sanitaria, médica y paramédica"
            i = i + 1
            .Cells(i, 3).Value = "Alquileres de inmuebles destinados a su casa habitación"
            i = i + 1
            .Cells(i, 3).Value = "Aportes al capital social o al fondo de riesgo de socios protectores de Sociedades de Garantía Recíproca"
            i = i + 1
            .Cells(i, 3).Value = "Empleados del servicio domestico (Ley 26.063, art. 16)"
            i = i + 1
            .Cells(i, 3).Value = "Otras Deducciones"
            i = i + 1
            .Cells(i, 3).Value = "TOTAL DEDUCCIONES GENERALES (suma de los incisos a al q)"
            i = i + 1
            .Cells(i, 3).Value = "DEDUCCIONES ART. 23"
            .Cells(i, 3).Font.Bold = True
            i = i + 1
            .Cells(i, 3).Value = "Ganancia no Imponible"
            i = i + 1
            .Cells(i, 3).Value = "Deducción Especial"
            i = i + 1
            .Cells(i, 3).Value = "Cargas de Familia"
            i = i + 1
            .Cells(i, 3).Value = "  - Cónyuge"
            i = i + 1
            .Cells(i, 3).Value = "  - Hijos"
            i = i + 1
            .Cells(i, 3).Value = "  - Otras Cargas"
            i = i + 1
            .Cells(i, 3).Value = "TOTAL DEDUCCIONES PERSONALES (suma de los incisos a al e)"
            i = i + 1
            .Cells(i, 3).Value = "REMUNERACIÓN SUJETA A IMPUESTO (suma algebraica rubro 1 al 3)"
            .Cells(i, 3).Font.Bold = True
            i = i + 1
            .Cells(i, 3).Value = "IMPUESTO DETERMINADO"
            .Cells(i, 3).Font.Bold = True
            i = i + 1
            .Cells(i, 3).Value = "IMPUESTO RETENIDO"
            .Cells(i, 3).Font.Bold = True
            i = i + 1
            .Cells(i, 3).Value = "PAGOS A CUENTA"
            .Cells(i, 3).Font.Bold = True
            i = i + 1
            .Cells(i, 3).Value = "SALDO (suma algebraica rubro 5 al 7)"
            .Cells(i, 3).Font.Bold = True
                      
            i = 18
            .Cells(i, 26).Value = ""
            .Range("Z" & i & ":AD" & i).Select
            With o_Excel.Selection
                .MergeCells = True
                .WrapText = True
                .Font.Size = 12
                .NumberFormat = "#,##0.00"
                .Interior.ColorIndex = 15
                .Interior.Pattern = xlSolid
                .Font.Bold = True
                .HorizontalAlignment = xlRight
                .VerticalAlignment = xlVAlignCenter
            End With
            .Range("Z" & i & ":AD" & i).Select
            o_Excel.Selection.Copy
            For i = 19 To 25
                .Range("Z" & i & ":AD" & i).Select
                .Paste
            Next i
            For i = 27 To 44
                .Range("Z" & i & ":AD" & i).Select
                .Paste
            Next i
            For i = 46 To 57
                .Range("Z" & i & ":AD" & i).Select
                .Paste
            Next i
            
            .Range("Z16:AD57").Select
            With o_Excel.Selection.Borders(xlEdgeLeft)
                .LineStyle = xlContinuous
                .Weight = xlMedium
                .ColorIndex = xlAutomatic
            End With
            With o_Excel.Selection.Borders(xlEdgeTop)
                .LineStyle = xlContinuous
                .Weight = xlMedium
                .ColorIndex = xlAutomatic
            End With
            With o_Excel.Selection.Borders(xlEdgeBottom)
                .LineStyle = xlContinuous
                .Weight = xlMedium
                .ColorIndex = xlAutomatic
            End With
            With o_Excel.Selection.Borders(xlEdgeRight)
                .LineStyle = xlContinuous
                .Weight = xlMedium
                .ColorIndex = xlAutomatic
            End With
            
            .Range("C17:AD25").Select
            With o_Excel.Selection.Borders(xlEdgeLeft)
                .LineStyle = xlContinuous
                .Weight = xlMedium
                .ColorIndex = xlAutomatic
            End With
            With o_Excel.Selection.Borders(xlEdgeTop)
                .LineStyle = xlContinuous
                .Weight = xlMedium
                .ColorIndex = xlAutomatic
            End With
            With o_Excel.Selection.Borders(xlEdgeBottom)
                .LineStyle = xlContinuous
                .Weight = xlMedium
                .ColorIndex = xlAutomatic
            End With
            With o_Excel.Selection.Borders(xlEdgeRight)
                .LineStyle = xlContinuous
                .Weight = xlMedium
                .ColorIndex = xlAutomatic
            End With

            .Range("C26:AD44").Select
            With o_Excel.Selection.Borders(xlEdgeLeft)
                .LineStyle = xlContinuous
                .Weight = xlMedium
                .ColorIndex = xlAutomatic
            End With
            With o_Excel.Selection.Borders(xlEdgeTop)
                .LineStyle = xlContinuous
                .Weight = xlMedium
                .ColorIndex = xlAutomatic
            End With
            With o_Excel.Selection.Borders(xlEdgeBottom)
                .LineStyle = xlContinuous
                .Weight = xlMedium
                .ColorIndex = xlAutomatic
            End With
            With o_Excel.Selection.Borders(xlEdgeRight)
                .LineStyle = xlContinuous
                .Weight = xlMedium
                .ColorIndex = xlAutomatic
            End With

            .Range("C45:AD52").Select
            With o_Excel.Selection.Borders(xlEdgeLeft)
                .LineStyle = xlContinuous
                .Weight = xlMedium
                .ColorIndex = xlAutomatic
            End With
            With o_Excel.Selection.Borders(xlEdgeTop)
                .LineStyle = xlContinuous
                .Weight = xlMedium
                .ColorIndex = xlAutomatic
            End With
            With o_Excel.Selection.Borders(xlEdgeBottom)
                .LineStyle = xlContinuous
                .Weight = xlMedium
                .ColorIndex = xlAutomatic
            End With
            With o_Excel.Selection.Borders(xlEdgeRight)
                .LineStyle = xlContinuous
                .Weight = xlMedium
                .ColorIndex = xlAutomatic
            End With

            For i = 53 To 57
                .Range("C" & i & ":AD" & i).Select
                With o_Excel.Selection.Borders(xlEdgeLeft)
                    .LineStyle = xlContinuous
                    .Weight = xlMedium
                    .ColorIndex = xlAutomatic
                End With
                With o_Excel.Selection.Borders(xlEdgeTop)
                    .LineStyle = xlContinuous
                    .Weight = xlMedium
                    .ColorIndex = xlAutomatic
                End With
                With o_Excel.Selection.Borders(xlEdgeBottom)
                    .LineStyle = xlContinuous
                    .Weight = xlMedium
                    .ColorIndex = xlAutomatic
                End With
                With o_Excel.Selection.Borders(xlEdgeRight)
                    .LineStyle = xlContinuous
                    .Weight = xlMedium
                    .ColorIndex = xlAutomatic
                End With
            Next i


            .Range("A59:AD62").Select
            With o_Excel.Selection.Borders(xlEdgeLeft)
                .LineStyle = xlContinuous
                .Weight = xlMedium
                .ColorIndex = xlAutomatic
            End With
            With o_Excel.Selection.Borders(xlEdgeTop)
                .LineStyle = xlContinuous
                .Weight = xlMedium
                .ColorIndex = xlAutomatic
            End With
            With o_Excel.Selection.Borders(xlEdgeBottom)
                .LineStyle = xlContinuous
                .Weight = xlMedium
                .ColorIndex = xlAutomatic
            End With
            With o_Excel.Selection.Borders(xlEdgeRight)
                .LineStyle = xlContinuous
                .Weight = xlMedium
                .ColorIndex = xlAutomatic
            End With
            .Cells(59, 1).Value = "OBSERVACIONES"

            
            .Cells(64, 3).Value = "Firma e identificación del agente de renteción:"
            .Range("C64:M64").Select
            With o_Excel.Selection
                .MergeCells = True
                .WrapText = True
                .HorizontalAlignment = xlVAlignCenter
                .VerticalAlignment = xlVAlignCenter
            End With
            .Range("C65:M67").Select
            With o_Excel.Selection
                .Interior.ColorIndex = 15
                .Interior.Pattern = xlSolid
            End With
            With o_Excel.Selection.Borders(xlEdgeLeft)
                .LineStyle = xlContinuous
                .Weight = xlMedium
                .ColorIndex = xlAutomatic
            End With
            With o_Excel.Selection.Borders(xlEdgeTop)
                .LineStyle = xlContinuous
                .Weight = xlMedium
                .ColorIndex = xlAutomatic
            End With
            With o_Excel.Selection.Borders(xlEdgeBottom)
                .LineStyle = xlContinuous
                .Weight = xlMedium
                .ColorIndex = xlAutomatic
            End With
            With o_Excel.Selection.Borders(xlEdgeRight)
                .LineStyle = xlContinuous
                .Weight = xlMedium
                .ColorIndex = xlAutomatic
            End With
            .Range("C64:M68").Select
            With o_Excel.Selection.Borders(xlEdgeLeft)
                .LineStyle = xlContinuous
                .Weight = xlMedium
                .ColorIndex = xlAutomatic
            End With
            With o_Excel.Selection.Borders(xlEdgeTop)
                .LineStyle = xlContinuous
                .Weight = xlMedium
                .ColorIndex = xlAutomatic
            End With
            With o_Excel.Selection.Borders(xlEdgeBottom)
                .LineStyle = xlContinuous
                .Weight = xlMedium
                .ColorIndex = xlAutomatic
            End With
            With o_Excel.Selection.Borders(xlEdgeRight)
                .LineStyle = xlContinuous
                .Weight = xlMedium
                .ColorIndex = xlAutomatic
            End With

            .Cells(68, 3).Value = "Lugar y fecha:"
            .Range("C68:F68").Select
            With o_Excel.Selection
                .MergeCells = True
                .WrapText = True
                .HorizontalAlignment = xlVAlignCenter
                .VerticalAlignment = xlVAlignCenter
            End With
            .Cells(68, 7).Value = ""
            .Range("G68:M68").Select
            With o_Excel.Selection
                .MergeCells = True
                .Font.Size = 12
                .Font.Bold = True
                .Interior.ColorIndex = 15
                .Interior.Pattern = xlSolid
                .HorizontalAlignment = xlLeft
                .VerticalAlignment = xlVAlignCenter
            End With

            .Range("C64:M68").Select
            o_Excel.Selection.Copy
            .Range("R64:AB68").Select
            .Paste
            .Cells(64, 18).Value = "Firma e identificación del beneficiario:"

            'Llenamos los Datos Personales del Formulario
            strPLF649 = LiquidacionFinalGanancias.txtPuestoLaboral.Text
            strYearF649 = LiquidacionFinalGanancias.txtPeriodo.Text
            .Cells(3, 1).Value = "Ejercicio " & strYearF649
            Set rstBuscarSlave = New ADODB.Recordset
            SQL = "Select * From AGENTES Where PUESTOLABORAL = '" & strPLF649 & "'"
            rstBuscarSlave.Open SQL, dbSlave, adOpenForwardOnly, adLockOptimistic
            .Cells(9, 1).Value = rstBuscarSlave!NombreCompleto
            .Cells(9, 16).Value = Format(rstBuscarSlave!CUIL, "00-00000000-0")
            rstBuscarSlave.Close
            Set rstBuscarSlave = Nothing
            .Cells(13, 1).Value = "INSTITUTO DE VIVIENDA DE CORRIENTES"
            .Cells(13, 16).Value = "30-63235151-4"

            'Buscamos el último código de liquidacion del año
'            SQL = "Select CODIGO From " _
'            & "CODIGOLIQUIDACIONES Inner Join LIQUIDACIONSUELDOS On " _
'            & "CODIGOLIQUIDACIONES.CODIGO = LIQUIDACIONSUELDOS.CODIGOLIQUIDACION " _
'            & "Where Right(PERIODO,4) = '" & strYearF649 & "' " _
'            & "Order by CODIGO Desc"
            SQL = "Select CODIGO From CODIGOLIQUIDACIONES " _
            & "Where Right(PERIODO,4) = '" & strYearF649 & "' " _
            & "Order by CODIGO Desc"
            Set rstBuscarSlave = New ADODB.Recordset
            rstBuscarSlave.Open SQL, dbSlave, adOpenForwardOnly, adLockOptimistic
            strCLF649 = rstBuscarSlave!Codigo
            rstBuscarSlave.Close
            Set rstBuscarSlave = Nothing
            'Completamos las Retribuciones no Habituales
            dblImporteCalculado = CalcularRetribucionNoHabitualAcumulada(strPLF649, strCLF649, True)
            dblRemuneracionComputable = dblImporteCalculado
            .Cells(19, 26).Value = dblImporteCalculado
            'Completamos SAC Primera Cuota
            dblImporteCalculado = CalcularSAC(strPLF649, strCLF649, "PrimerSAC", True)
            dblRemuneracionComputable = dblRemuneracionComputable + dblImporteCalculado
            .Cells(20, 26).Value = dblImporteCalculado
            'Completamos SAC Segunda Cuota
            dblImporteCalculado = CalcularSAC(strPLF649, strCLF649, "SegundoSAC", True)
            dblRemuneracionComputable = dblRemuneracionComputable + dblImporteCalculado
            .Cells(21, 26).Value = dblImporteCalculado
            'Completamos la Remuneración Bruta
            dblImporteCalculado = CalcularHaberBrutoAcumulado(strPLF649, strCLF649, True, True) - dblRemuneracionComputable
            dblRemuneracionComputable = dblRemuneracionComputable + dblImporteCalculado
            .Cells(18, 26).Value = dblImporteCalculado
            'Completamos la Remuneración No Alcanzada
            .Cells(22, 26).Value = CalcularRemuneracionNoAlcanzadaAcumulada(strPLF649, strCLF649, True)
            'Completamos la Remuneración Exenta (NO PREVISTO)
            .Cells(23, 26).Value = 0
            'Completamos importe Pluriempleo
            dblImporteCalculado = CalcularPluriempleoAcumulado(strPLF649, strCLF649, True)
            dblRemuneracionComputable = dblRemuneracionComputable + dblImporteCalculado
            .Cells(24, 26).Value = dblImporteCalculado
            'Completamos Remuneración Computable
            .Cells(25, 26).Value = dblRemuneracionComputable

            'Deducciones Generales
            'Completamos importe Aporte Jubilatorio
            dblImporteCalculado = CalcularJubilacionAcumulada(strPLF649, strCLF649, True)
            dblDeduccionesGenerales = dblImporteCalculado
            .Cells(27, 26).Value = dblImporteCalculado
            'Completamos importe Aporte Obra Social
            dblImporteCalculado = CalcularObraSocialAcumulada(strPLF649, strCLF649, True)
            dblImporteCalculado = dblImporteCalculado + ImporteRegistradoAcumuladoDeduccionEspecifica("AdherenteObraSocial", _
            strPLF649, "12/" & strYearF649, strCLF649)
            dblDeduccionesGenerales = dblDeduccionesGenerales + dblImporteCalculado
            .Cells(28, 26).Value = dblImporteCalculado
            'Completamos importe Cuota Sindical
            dblImporteCalculado = CalcularDescuentoCuotaSindicalAcumulado(strPLF649, strCLF649, True, True)
            dblDeduccionesGenerales = dblDeduccionesGenerales + dblImporteCalculado
            .Cells(29, 26).Value = dblImporteCalculado
            'Completamos importe Aportes Jubilatorios Otros Empleos (No Previsto)
            .Cells(30, 26).Value = 0
            'Completamos importe Aportes Obra Social otros empleos (No Previsto)
            .Cells(31, 26).Value = 0
            'Completamos importe Cuota Sindical otros empleos (No Previsto)
            .Cells(32, 26).Value = 0
            'Completamos importe Cuotas médico asistenciales
            dblImporteCalculado = ImporteRegistradoAcumuladoDeduccionEspecifica("CuotaMedicoAsistencial", _
            strPLF649, "12/" & strYearF649, strCLF649)
            dblDeduccionesGenerales = dblDeduccionesGenerales + dblImporteCalculado
            .Cells(33, 26).Value = dblImporteCalculado
            'Completamos importe Primas de Seguro para el caso de muerte
            dblImporteCalculado = ImporteRegistradoAcumuladoDeduccionEspecifica("SeguroDeVidaOptativo", _
            strPLF649, "12/" & strYearF649, strCLF649)
            dblDeduccionesGenerales = dblDeduccionesGenerales + dblImporteCalculado
            .Cells(34, 26).Value = dblImporteCalculado
            'Completamos importe Gasto de Sepelio (No Previsto)
            .Cells(35, 26).Value = 0
            'Completamos importe Gastos estimativos para corredores y viajantes de comercio (No Previsto)
            .Cells(36, 26).Value = 0
            'Completamos importe Donaciones
            dblImporteCalculado = ImporteRegistradoAcumuladoDeduccionEspecifica("Donaciones", _
            strPLF649, "12/" & strYearF649, strCLF649)
            dblDeduccionesGenerales = dblDeduccionesGenerales + dblImporteCalculado
            .Cells(37, 26).Value = dblImporteCalculado
            'Completamos importe Descuentos obligatorios establecidos por ley nacional, provincial o municipal
            dblImporteCalculado = ImporteRegistradoAcumuladoDeduccionEspecifica("SeguroDeVidaObligatorio", _
            strPLF649, "12/" & strYearF649, strCLF649)
            dblDeduccionesGenerales = dblDeduccionesGenerales + dblImporteCalculado
            .Cells(38, 26).Value = dblImporteCalculado
            'Completamos importe Honorarios por servicios de asistencia sanitaria, médica y paramédica
            dblImporteCalculado = ImporteRegistradoAcumuladoDeduccionEspecifica("HonorariosMedicos", _
            strPLF649, "12/" & strYearF649, strCLF649)
            dblDeduccionesGenerales = dblDeduccionesGenerales + dblImporteCalculado
            .Cells(39, 26).Value = dblImporteCalculado
            'Completamos importe Intereses Créditos Hipotecarios (No Previsto)
            dblImporteCalculado = ImporteRegistradoAcumuladoDeduccionEspecifica("Alquileres", _
            strPLF649, "12/" & strYearF649, strCLF649)
            dblDeduccionesGenerales = dblDeduccionesGenerales + dblImporteCalculado
            .Cells(40, 26).Value = dblImporteCalculado
            'Completamos importe Aportes al capital social o al fondo de riesgo de socios protectores de Sociedades de Garantía Recíproca (No Previsto)
            .Cells(41, 26).Value = 0
            'Completamos importe Empleados del servicio domestico (Ley 26.063, art. 16)
            dblImporteCalculado = ImporteRegistradoAcumuladoDeduccionEspecifica("ServicioDomestico", _
            strPLF649, "12/" & strYearF649, strCLF649)
            dblDeduccionesGenerales = dblDeduccionesGenerales + dblImporteCalculado
            .Cells(42, 26).Value = dblImporteCalculado
            'Completamos importe Otras Deducciones (No Previsto)
            .Cells(43, 26).Value = 0
            'Completamos TOTAL DEDUCCIONES GENERALES
            .Cells(44, 26).Value = dblDeduccionesGenerales
            
            'Deducciones Personales
            'Completamos importe Ganancia No Imponible
            dblImporteCalculado = ImporteRegistradoAcumuladoDeduccionEspecifica("MinimoNoImponible", _
            strPLF649, "12/" & strYearF649, strCLF649)
            dblDeduccionesPersonales = dblImporteCalculado
            .Cells(46, 26).Value = dblImporteCalculado
            'Completamos importe Deducción Especial
            dblImporteCalculado = ImporteRegistradoAcumuladoDeduccionEspecifica("DeduccionEspecial", _
            strPLF649, "12/" & strYearF649, strCLF649)
            dblDeduccionesPersonales = dblDeduccionesPersonales + dblImporteCalculado
            .Cells(47, 26).Value = dblImporteCalculado
            'Completamos importe Conyuge
            dblImporteCalculado = ImporteRegistradoAcumuladoDeduccionEspecifica("Conyuge", _
            strPLF649, "12/" & strYearF649, strCLF649)
            dblDeduccionesPersonales = dblDeduccionesPersonales + dblImporteCalculado
            .Cells(49, 26).Value = dblImporteCalculado
            'Completamos importe Hijo
            dblImporteCalculado = ImporteRegistradoAcumuladoDeduccionEspecifica("Hijo", _
            strPLF649, "12/" & strYearF649, strCLF649)
            dblDeduccionesPersonales = dblDeduccionesPersonales + dblImporteCalculado
            .Cells(50, 26).Value = dblImporteCalculado
            'Completamos importe Otras Cargas de Familia
            dblImporteCalculado = ImporteRegistradoAcumuladoDeduccionEspecifica("OtrasCargasDeFamilia", _
            strPLF649, "12/" & strYearF649, strCLF649)
            dblDeduccionesPersonales = dblDeduccionesPersonales + dblImporteCalculado
            .Cells(51, 26).Value = dblImporteCalculado
            'Completamos TOTAL DEDUCCIONES PERSONALES
            .Cells(52, 26).Value = dblDeduccionesPersonales
            
            'Completamos REMUNERACIÓN SUJETA A IMPUESTO
            dblRemuneracionComputable = dblRemuneracionComputable - dblDeduccionesGenerales - dblDeduccionesPersonales
            .Cells(53, 26).Value = dblRemuneracionComputable
            
            'Completamos Impuesto Determinado
            dblRemuneracionComputable = CalcularImporteFijo(strPLF649, strCLF649, dblRemuneracionComputable) _
            + CalcularImporteVariable(strPLF649, strCLF649, , dblRemuneracionComputable)
            .Cells(54, 26).Value = dblRemuneracionComputable
            
            'Completamos Impuesto Retenido
            .Cells(55, 26).Value = ImporteRegistradoAcumuladoDeduccionEspecifica("Retencion", _
            strPLF649, "12/" & strYearF649, strCLF649, False)
            '.Cells(55, 26).Value = CalcularConceptoSISPERAcumulado(strPLF649, strCLF649, _
            "0276", True)
            
            'Completamos Pago a Cuenta
            .Cells(56, 26).Value = 0

            'Completamos Saldo
            '.Cells(57, 26).Value = dblRemuneracionComputable - CalcularConceptoSISPERAcumulado(strPLF649, strCLF649, _
            "0276", True)
            .Cells(57, 26).Value = dblRemuneracionComputable - ImporteRegistradoAcumuladoDeduccionEspecifica("Retencion", _
            strPLF649, "12/" & strYearF649, strCLF649, False)

        End With
       o_Libro.Close True, strOutputPath
        ' -- Cerrar Excel
        o_Excel.Quit
        ' -- Terminar instancias
        Call ReleaseObjects(o_Excel, o_Libro, o_Hoja)
        GenerarF649 = True
    End If
    
Exit Function
  
' -- Controlador de Errores
Error_Handler:
    ' -- Cierra la hoja y el la aplicación Excel
    If Not o_Libro Is Nothing Then: o_Libro.Close False
    If Not o_Excel Is Nothing Then: o_Excel.Quit
    Call ReleaseObjects(o_Excel, o_Libro, o_Hoja)
    If Err.Number <> 1004 Then MsgBox Err.Description, vbCritical
End Function

Public Function GenerarReporteMensualGanancias(strOutputPath As String, CodigoLiquidacion As String) As Boolean
    On Error GoTo Error_Handler
  
    Dim o_Excel                 As Object
    Dim o_Libro                 As Object
    Dim o_Hoja                  As Object
    Dim ColumnaExcel            As Integer
    Dim FilaExcel               As Integer
    Dim SQL                     As String
    Dim PeriodoLiquidacion      As String
    Dim DescripcionLiquidacion  As String
    Dim ApellidoAgente()        As String
    Dim SumaTotal1              As Double
    Dim SumaTotal2              As Double
    Dim dblAlicuotaCalculada    As Double
    Dim strApellidoYNombreCorto As String
       
    ' -- Crea el objeto Excel, el objeto workBook y el objeto sheet
    Set o_Excel = CreateObject("Excel.Application")
    Set o_Libro = o_Excel.Workbooks.Add
    Set o_Hoja = o_Libro.Worksheets(1)
       
    SQL = "Select * From CODIGOLIQUIDACIONES Where CODIGO = '" & CodigoLiquidacion & "'"
    Set rstRegistroSlave = New ADODB.Recordset
    rstRegistroSlave.Open SQL, dbSlave, adOpenForwardOnly, adLockOptimistic
    DescripcionLiquidacion = rstRegistroSlave!Descripcion
    PeriodoLiquidacion = rstRegistroSlave!Periodo
    rstRegistroSlave.Close
    Set rstRegistroSlave = Nothing
    
    ' -- Bucle para Exportar los datos
    Set rstListadoSlave = New ADODB.Recordset
    Set rstBuscarSlave = New ADODB.Recordset
    SQL = "Select LIQUIDACIONGANANCIAS4TACATEGORIA.* " _
    & "From LIQUIDACIONGANANCIAS4TACATEGORIA Inner Join AGENTES " _
    & "On LIQUIDACIONGANANCIAS4TACATEGORIA.PuestoLaboral = AGENTES.PuestoLaboral " _
    & "Where LIQUIDACIONGANANCIAS4TACATEGORIA.CODIGOLIQUIDACION = '" & CodigoLiquidacion & "' " _
    & "And LIQUIDACIONGANANCIAS4TACATEGORIA.RETENCION <> 0 " _
    & "Order by AGENTES.NombreCompleto Asc"
    rstListadoSlave.Open SQL, dbSlave, adOpenDynamic, adLockOptimistic
    FilaExcel = 8
    rstListadoSlave.MoveFirst
    With o_Hoja
        Do
            ColumnaExcel = 1
            .Cells(FilaExcel - 6, 1).Value = "RETENCIÓN DE IMPUESTO A LAS GANANCIAS 4TA CATEGORÍA"
            .Range("A" & FilaExcel - 6 & ":O" & FilaExcel - 6).MergeCells = True
            .Range("A" & FilaExcel - 6 & ":O" & FilaExcel - 6).Font.Size = 12
            .Range("A" & FilaExcel - 6 & ":O" & FilaExcel - 6).Font.Bold = True
            .Range("A" & FilaExcel - 6 & ":O" & FilaExcel - 6).HorizontalAlignment = xlVAlignCenter
            .Cells(FilaExcel - 4, 1).Value = DescripcionLiquidacion
            .Range("A" & FilaExcel - 4 & ":O" & FilaExcel - 4).MergeCells = True
            .Range("A" & FilaExcel - 4 & ":O" & FilaExcel - 4).Font.Size = 11
            .Range("A" & FilaExcel - 4 & ":O" & FilaExcel - 4).Font.Bold = True
            .Range("A" & FilaExcel - 4 & ":O" & FilaExcel - 4).HorizontalAlignment = xlVAlignCenter
            
            'Configurando Bordes de la Grilla
            .Range("B" & FilaExcel & ":O" & FilaExcel + 36).Select
            o_Excel.Selection.Borders(xlDiagonalDown).LineStyle = xlNone
            o_Excel.Selection.Borders(xlDiagonalUp).LineStyle = xlNone
            With o_Excel.Selection.Borders(xlEdgeLeft)
                .LineStyle = xlContinuous
                .Weight = xlMedium
                .ColorIndex = xlAutomatic
            End With
            With o_Excel.Selection.Borders(xlEdgeTop)
                .LineStyle = xlContinuous
                .Weight = xlMedium
                .ColorIndex = xlAutomatic
            End With
            With o_Excel.Selection.Borders(xlEdgeBottom)
                .LineStyle = xlContinuous
                .Weight = xlMedium
                .ColorIndex = xlAutomatic
            End With
            With o_Excel.Selection.Borders(xlEdgeRight)
                .LineStyle = xlContinuous
                .Weight = xlMedium
                .ColorIndex = xlAutomatic
            End With
            With o_Excel.Selection.Borders(xlInsideVertical)
                .LineStyle = xlContinuous
                .Weight = xlMedium
                .ColorIndex = xlAutomatic
            End With
            With o_Excel.Selection.Borders(xlInsideHorizontal)
                .LineStyle = xlContinuous
                .Weight = xlThin
                .ColorIndex = xlAutomatic
            End With
            .Range("A" & FilaExcel & ":A" & FilaExcel + 38).Select
            o_Excel.Selection.Borders(xlDiagonalDown).LineStyle = xlNone
            o_Excel.Selection.Borders(xlDiagonalUp).LineStyle = xlNone
            With o_Excel.Selection.Borders(xlEdgeLeft)
                .LineStyle = xlContinuous
                .Weight = xlMedium
                .ColorIndex = xlAutomatic
            End With
            With o_Excel.Selection.Borders(xlEdgeTop)
                .LineStyle = xlContinuous
                .Weight = xlMedium
                .ColorIndex = xlAutomatic
            End With
            With o_Excel.Selection.Borders(xlEdgeBottom)
                .LineStyle = xlContinuous
                .Weight = xlMedium
                .ColorIndex = xlAutomatic
            End With
            With o_Excel.Selection.Borders(xlEdgeRight)
                .LineStyle = xlContinuous
                .Weight = xlMedium
                .ColorIndex = xlAutomatic
            End With
            .Range("A" & FilaExcel + 37 & ":O" & FilaExcel + 38).Select
            o_Excel.Selection.Borders(xlDiagonalDown).LineStyle = xlNone
            o_Excel.Selection.Borders(xlDiagonalUp).LineStyle = xlNone
            With o_Excel.Selection.Borders(xlEdgeLeft)
                .LineStyle = xlContinuous
                .Weight = xlMedium
                .ColorIndex = xlAutomatic
            End With
            With o_Excel.Selection.Borders(xlEdgeTop)
                .LineStyle = xlContinuous
                .Weight = xlMedium
                .ColorIndex = xlAutomatic
            End With
            With o_Excel.Selection.Borders(xlEdgeBottom)
                .LineStyle = xlContinuous
                .Weight = xlMedium
                .ColorIndex = xlAutomatic
            End With
            With o_Excel.Selection.Borders(xlEdgeRight)
                .LineStyle = xlContinuous
                .Weight = xlMedium
                .ColorIndex = xlAutomatic
            End With
            With o_Excel.Selection.Borders(xlInsideVertical)
                .LineStyle = xlContinuous
                .Weight = xlMedium
                .ColorIndex = xlAutomatic
            End With
            .Range("B" & FilaExcel - 2 & ":O" & FilaExcel - 1).Select
            o_Excel.Selection.Borders(xlDiagonalDown).LineStyle = xlNone
            o_Excel.Selection.Borders(xlDiagonalUp).LineStyle = xlNone
            With o_Excel.Selection.Borders(xlEdgeLeft)
                .LineStyle = xlContinuous
                .Weight = xlMedium
                .ColorIndex = xlAutomatic
            End With
            With o_Excel.Selection.Borders(xlEdgeTop)
                .LineStyle = xlContinuous
                .Weight = xlMedium
                .ColorIndex = xlAutomatic
            End With
            With o_Excel.Selection.Borders(xlEdgeBottom)
                .LineStyle = xlContinuous
                .Weight = xlMedium
                .ColorIndex = xlAutomatic
            End With
            With o_Excel.Selection.Borders(xlEdgeRight)
                .LineStyle = xlContinuous
                .Weight = xlMedium
                .ColorIndex = xlAutomatic
            End With
            With o_Excel.Selection.Borders(xlInsideVertical)
                .LineStyle = xlContinuous
                .Weight = xlMedium
                .ColorIndex = xlAutomatic
            End With
            .Range("A" & FilaExcel & ":O" & FilaExcel).Select
            o_Excel.Selection.Borders(xlDiagonalDown).LineStyle = xlNone
            o_Excel.Selection.Borders(xlDiagonalUp).LineStyle = xlNone
            With o_Excel.Selection.Borders(xlEdgeLeft)
                .LineStyle = xlContinuous
                .Weight = xlMedium
                .ColorIndex = xlAutomatic
            End With
            With o_Excel.Selection.Borders(xlEdgeTop)
                .LineStyle = xlContinuous
                .Weight = xlMedium
                .ColorIndex = xlAutomatic
            End With
            With o_Excel.Selection.Borders(xlEdgeBottom)
                .LineStyle = xlContinuous
                .Weight = xlMedium
                .ColorIndex = xlAutomatic
            End With
            With o_Excel.Selection.Borders(xlEdgeRight)
                .LineStyle = xlContinuous
                .Weight = xlMedium
                .ColorIndex = xlAutomatic
            End With
            
            'Configurando Tamaño Fuente
            .Range("B" & FilaExcel - 2 & ":O" & FilaExcel + 38).Select
            o_Excel.Selection.Font.Size = 10
            
            'Configurando Formato de Números
            .Range("B" & FilaExcel + 2 & ":O" & FilaExcel + 38).Select
            o_Excel.Selection.Style = "Currency"
            
            .Cells(FilaExcel, 1).Value = "CONCEPTO"
            .Cells(FilaExcel, 1).Font.Bold = True
            .Cells(FilaExcel, 1).HorizontalAlignment = xlVAlignCenter
            .Cells(FilaExcel + 2, 1).Value = "Rentas Habituales"
            .Cells(FilaExcel + 3, 1).Value = "Pluriempleo (Neto)"
            .Cells(FilaExcel + 4, 1).Value = "Ajustes"
            .Cells(FilaExcel + 6, 1).Value = "GCIA. BRUTA"
            .Range("A" & FilaExcel + 6 & ":O" & FilaExcel + 6).Select
            o_Excel.Selection.Font.Bold = True
            .Cells(FilaExcel + 6, 1).HorizontalAlignment = xlVAlignCenter
            .Cells(FilaExcel + 8, 1).Value = "Jubilación Personal"
            .Cells(FilaExcel + 9, 1).Value = "O.Social Personal"
            .Cells(FilaExcel + 10, 1).Value = "Adherente O.Social"
            .Cells(FilaExcel + 11, 1).Value = "Seguro Obligatorio"
            .Cells(FilaExcel + 12, 1).Value = "Cuota Sindical"
            .Cells(FilaExcel + 14, 1).Value = "GCIA. NETA"
            .Range("A" & FilaExcel + 14 & ":O" & FilaExcel + 14).Select
            o_Excel.Selection.Font.Bold = True
            .Cells(FilaExcel + 14, 1).HorizontalAlignment = xlVAlignCenter
            .Cells(FilaExcel + 16, 1).Value = "Honorarios Médicos" 'Antes decía Gtos. Sepelio
            .Cells(FilaExcel + 17, 1).Value = "Alquiler" 'Antes decía Int. Hipotecarios
            .Cells(FilaExcel + 18, 1).Value = "Seguro de Vida"
            .Cells(FilaExcel + 19, 1).Value = "Serv. Doméstico"
            .Cells(FilaExcel + 20, 1).Value = "Cuota Médico Asist."
            .Cells(FilaExcel + 21, 1).Value = "Donaciones"
            .Cells(FilaExcel + 22, 1).Value = "Gcia. No Imponible"
            .Cells(FilaExcel + 23, 1).Value = "Conyuge"
            .Cells(FilaExcel + 24, 1).Value = "Hijos"
            .Cells(FilaExcel + 25, 1).Value = "Otras Cargas"
            .Cells(FilaExcel + 26, 1).Value = "Deducción Especial"
            .Cells(FilaExcel + 28, 1).Value = "GCIA. IMPONIBLE"
            .Range("A" & FilaExcel + 28 & ":O" & FilaExcel + 28).Select
            o_Excel.Selection.Font.Bold = True
            .Cells(FilaExcel + 28, 1).HorizontalAlignment = xlVAlignCenter
            .Cells(FilaExcel + 30, 1).Value = "Porcentaje"
            .Cells(FilaExcel + 31, 1).Value = "Importe Fijo"
            .Cells(FilaExcel + 32, 1).Value = "Importe Variable"
            .Cells(FilaExcel + 33, 1).Value = "IMP. DETERMINADO"
            .Range("A" & FilaExcel + 33 & ":O" & FilaExcel + 35).Select
            o_Excel.Selection.Font.Bold = True
            .Cells(FilaExcel + 33, 1).HorizontalAlignment = xlVAlignCenter
            .Cells(FilaExcel + 34, 1).Value = "RETENCIÓN ACUM."
            .Cells(FilaExcel + 34, 1).HorizontalAlignment = xlVAlignCenter
            .Cells(FilaExcel + 35, 1).Value = "AJUSTES ACUM."
            .Cells(FilaExcel + 35, 1).HorizontalAlignment = xlVAlignCenter
            .Cells(FilaExcel + 37, 1).Value = "RETENER"
            .Range("A" & FilaExcel + 37 & ":O" & FilaExcel + 38).Select
            o_Excel.Selection.Font.Bold = True
            .Cells(FilaExcel + 37, 1).HorizontalAlignment = xlVAlignCenter
            .Cells(FilaExcel + 38, 1).Value = "REINTEGRAR"
            .Cells(FilaExcel + 38, 1).HorizontalAlignment = xlVAlignCenter
            .Cells(FilaExcel + 43, 4).Value = "C.P.MIRTA S.SÁNCHEZ GÓMEZ"
            .Range("D" & FilaExcel + 43 & ":F" & FilaExcel + 43).MergeCells = True
            .Range("D" & FilaExcel + 43 & ":F" & FilaExcel + 43).Font.Size = 10
            .Range("D" & FilaExcel + 43 & ":F" & FilaExcel + 43).Font.Bold = True
            .Range("D" & FilaExcel + 43 & ":F" & FilaExcel + 43).HorizontalAlignment = xlVAlignCenter
            .Cells(FilaExcel + 44, 4).Value = "JEFE DPTO.CONTABLE"
            .Range("D" & FilaExcel + 44 & ":F" & FilaExcel + 44).MergeCells = True
            .Range("D" & FilaExcel + 44 & ":F" & FilaExcel + 44).Font.Size = 10
            .Range("D" & FilaExcel + 44 & ":F" & FilaExcel + 44).Font.Bold = True
            .Range("D" & FilaExcel + 44 & ":F" & FilaExcel + 44).HorizontalAlignment = xlVAlignCenter
            .Cells(FilaExcel + 45, 4).Value = "IN.VI.CO"
            .Range("D" & FilaExcel + 45 & ":F" & FilaExcel + 45).MergeCells = True
            .Range("D" & FilaExcel + 45 & ":F" & FilaExcel + 45).Font.Size = 10
            .Range("D" & FilaExcel + 45 & ":F" & FilaExcel + 45).Font.Bold = True
            .Range("D" & FilaExcel + 45 & ":F" & FilaExcel + 45).HorizontalAlignment = xlVAlignCenter
            .Cells(FilaExcel + 43, 10).Value = "Cra. MARÍA BELÉN BORAKIEVICH"
            .Range("J" & FilaExcel + 43 & ":L" & FilaExcel + 43).MergeCells = True
            .Range("J" & FilaExcel + 43 & ":L" & FilaExcel + 43).Font.Size = 10
            .Range("J" & FilaExcel + 43 & ":L" & FilaExcel + 43).Font.Bold = True
            .Range("J" & FilaExcel + 43 & ":L" & FilaExcel + 43).HorizontalAlignment = xlVAlignCenter
            .Cells(FilaExcel + 44, 10).Value = "A/C Gerencia de Administración y Finanzas"
            .Range("J" & FilaExcel + 44 & ":L" & FilaExcel + 44).MergeCells = True
            .Range("J" & FilaExcel + 44 & ":L" & FilaExcel + 44).Font.Size = 10
            .Range("J" & FilaExcel + 44 & ":L" & FilaExcel + 44).Font.Bold = True
            .Range("J" & FilaExcel + 44 & ":L" & FilaExcel + 44).HorizontalAlignment = xlVAlignCenter
            .Cells(FilaExcel + 45, 10).Value = "Instituto de Vivienda de Corrientes"
            .Range("J" & FilaExcel + 45 & ":L" & FilaExcel + 45).MergeCells = True
            .Range("J" & FilaExcel + 45 & ":L" & FilaExcel + 45).Font.Size = 10
            .Range("J" & FilaExcel + 45 & ":L" & FilaExcel + 45).Font.Bold = True
            .Range("J" & FilaExcel + 45 & ":L" & FilaExcel + 45).HorizontalAlignment = xlVAlignCenter
            .Cells(FilaExcel + 44, 15).Value = "Generado por"
            .Range("N" & FilaExcel + 44 & ":O" & FilaExcel + 44).MergeCells = True
            .Range("N" & FilaExcel + 44 & ":O" & FilaExcel + 44).Font.Size = 8
            .Range("N" & FilaExcel + 44 & ":O" & FilaExcel + 44).Font.Bold = True
            .Range("N" & FilaExcel + 44 & ":O" & FilaExcel + 44).HorizontalAlignment = xlVAlignCenter
            .Cells(FilaExcel + 45, 15).Value = "SLAVE v 2.0"
            .Range("N" & FilaExcel + 45 & ":O" & FilaExcel + 45).MergeCells = True
            .Range("N" & FilaExcel + 45 & ":O" & FilaExcel + 45).Font.Size = 8
            .Range("N" & FilaExcel + 45 & ":O" & FilaExcel + 45).Font.Bold = True
            .Range("N" & FilaExcel + 45 & ":O" & FilaExcel + 45).HorizontalAlignment = xlVAlignCenter

            While rstListadoSlave.EOF = False And ColumnaExcel <= 14
                SQL = "Select * From AGENTES Where PUESTOLABORAL = '" & rstListadoSlave!PuestoLaboral & "'"
                rstBuscarSlave.Open SQL, dbSlave, adOpenDynamic, adLockOptimistic
                ApellidoAgente = Split(rstBuscarSlave!NombreCompleto, " ")
                i = 1
                strApellidoYNombreCorto = ApellidoAgente(i)
                While Left(strApellidoYNombreCorto, 1) = ""
                    i = i + 1
                    strApellidoYNombreCorto = ApellidoAgente(i)
                Wend
                strApellidoYNombreCorto = ApellidoAgente(0) & " " & Left(ApellidoAgente(i), 1) & "."
                .Cells(FilaExcel - 2, ColumnaExcel + 1).Value = strApellidoYNombreCorto
                .Cells(FilaExcel - 2, ColumnaExcel + 1).Font.Bold = True
                .Cells(FilaExcel - 2, ColumnaExcel + 1).HorizontalAlignment = xlVAlignCenter
                .Cells(FilaExcel - 1, ColumnaExcel + 1).Value = Format(rstBuscarSlave!CUIL, "00-00000000-0")
                .Cells(FilaExcel - 1, ColumnaExcel + 1).HorizontalAlignment = xlVAlignCenter
                .Cells(FilaExcel, ColumnaExcel + 1).Value = "Acumulado"
                .Cells(FilaExcel, ColumnaExcel + 1).HorizontalAlignment = xlVAlignCenter
                rstBuscarSlave.Close
                SQL = "Select Sum(LIQUIDACIONGANANCIAS4TACATEGORIA.HABEROPTIMO) As SumaHaberOptimo, Sum(LIQUIDACIONGANANCIAS4TACATEGORIA.PLURIEMPLEO) As SumaPluriempleo, " _
                & "Sum(LIQUIDACIONGANANCIAS4TACATEGORIA.AJUSTE) As SumaAjuste, Sum(LIQUIDACIONGANANCIAS4TACATEGORIA.JUBILACION) As SumaJubilacion, " _
                & "Sum(LIQUIDACIONGANANCIAS4TACATEGORIA.OBRASOCIAL) As SumaObraSocial, Sum(LIQUIDACIONGANANCIAS4TACATEGORIA.ADHERENTEOBRASOCIAL) As SumaAdherente, " _
                & "Sum(LIQUIDACIONGANANCIAS4TACATEGORIA.DONACIONES) As SumaDonaciones, Sum(LIQUIDACIONGANANCIAS4TACATEGORIA.HONORARIOSMEDICOS) As SumaHonorariosMedicos, Sum(LIQUIDACIONGANANCIAS4TACATEGORIA.ALQUILERES) As SumaAlquileres, " _
                & "Sum(LIQUIDACIONGANANCIAS4TACATEGORIA.SEGURODEVIDAOBLIGATORIO) As SumaSeguroObligatorio, Sum(LIQUIDACIONGANANCIAS4TACATEGORIA.CUOTASINDICAL) As SumaCuotaSindical, " _
                & "Sum(LIQUIDACIONGANANCIAS4TACATEGORIA.SEGURODEVIDAOPTATIVO) As SumaSeguroOptativo, Sum(LIQUIDACIONGANANCIAS4TACATEGORIA.SERVICIODOMESTICO) As SumaServicioDomestico, " _
                & "Sum(LIQUIDACIONGANANCIAS4TACATEGORIA.CUOTAMEDICOASISTENCIAL) As SumaCuotaMedicoAsistencial, Sum(LIQUIDACIONGANANCIAS4TACATEGORIA.MINIMONOIMPONIBLE) As SumaMinimoNoImponible, " _
                & "Sum(LIQUIDACIONGANANCIAS4TACATEGORIA.CONYUGE) As SumaConyuge, Sum(LIQUIDACIONGANANCIAS4TACATEGORIA.HIJO) As SumaHijo, " _
                & "Sum(LIQUIDACIONGANANCIAS4TACATEGORIA.OTRASCARGASDEFAMILIA) As SumaOtrasCargas, Sum(LIQUIDACIONGANANCIAS4TACATEGORIA.DEDUCCIONESPECIAL) As SumaDeduccionEspecial " _
                & "From CODIGOLIQUIDACIONES INNER JOIN LIQUIDACIONGANANCIAS4TACATEGORIA ON CODIGOLIQUIDACIONES.Codigo = LIQUIDACIONGANANCIAS4TACATEGORIA.CodigoLiquidacion " _
                & "Where PUESTOLABORAL = '" & rstListadoSlave!PuestoLaboral & "' And CODIGOLIQUIDACION <= '" & CodigoLiquidacion & "' " _
                & "And Right(PERIODO,4) = '" & Right(PeriodoLiquidacion, 4) & "'"
                rstBuscarSlave.Open SQL, dbSlave, adOpenDynamic, adLockOptimistic
                .Cells(FilaExcel + 2, ColumnaExcel + 1).Value = rstBuscarSlave!SumaHaberOptimo
                SumaTotal1 = rstBuscarSlave!SumaHaberOptimo
                .Cells(FilaExcel + 3, ColumnaExcel + 1).Value = rstBuscarSlave!SumaPluriempleo
                SumaTotal1 = SumaTotal1 + rstBuscarSlave!SumaPluriempleo
                .Cells(FilaExcel + 4, ColumnaExcel + 1).Value = rstBuscarSlave!SumaAjuste
                SumaTotal1 = SumaTotal1 + rstBuscarSlave!SumaAjuste
                .Cells(FilaExcel + 6, ColumnaExcel + 1).Value = SumaTotal1
                .Cells(FilaExcel + 8, ColumnaExcel + 1).Value = rstBuscarSlave!SumaJubilacion
                SumaTotal2 = SumaTotal2 + rstBuscarSlave!SumaJubilacion
                .Cells(FilaExcel + 9, ColumnaExcel + 1).Value = rstBuscarSlave!SumaObraSocial
                SumaTotal2 = SumaTotal2 + rstBuscarSlave!SumaObraSocial
                .Cells(FilaExcel + 10, ColumnaExcel + 1).Value = rstBuscarSlave!SumaAdherente
                SumaTotal2 = SumaTotal2 + rstBuscarSlave!SumaAdherente
                .Cells(FilaExcel + 11, ColumnaExcel + 1).Value = rstBuscarSlave!SumaSeguroObligatorio
                SumaTotal2 = SumaTotal2 + rstBuscarSlave!SumaSeguroObligatorio
                .Cells(FilaExcel + 12, ColumnaExcel + 1).Value = rstBuscarSlave!SumaCuotaSindical
                SumaTotal2 = SumaTotal2 + rstBuscarSlave!SumaCuotaSindical
                SumaTotal1 = SumaTotal1 - SumaTotal2
                SumaTotal2 = 0
                .Cells(FilaExcel + 14, ColumnaExcel + 1).Value = SumaTotal1
                'Gastos de Sepelio no tenido en cuenta (antes estaba en vez de Honorarios Médicos)
                .Cells(FilaExcel + 16, ColumnaExcel + 1).Value = rstBuscarSlave!SumaHonorariosMedicos
                SumaTotal2 = SumaTotal2 + rstBuscarSlave!SumaHonorariosMedicos
                'Intereses Hipotecarios no tenido en cuenta (antes estaba en vez de Alquileres)
                .Cells(FilaExcel + 17, ColumnaExcel + 1).Value = rstBuscarSlave!SumaAlquileres
                SumaTotal2 = SumaTotal2 + rstBuscarSlave!SumaAlquileres
                .Cells(FilaExcel + 18, ColumnaExcel + 1).Value = rstBuscarSlave!SumaSeguroOptativo
                SumaTotal2 = SumaTotal2 + rstBuscarSlave!SumaSeguroOptativo
                .Cells(FilaExcel + 19, ColumnaExcel + 1).Value = rstBuscarSlave!SumaServicioDomestico
                SumaTotal2 = SumaTotal2 + rstBuscarSlave!SumaServicioDomestico
                .Cells(FilaExcel + 20, ColumnaExcel + 1).Value = rstBuscarSlave!SumaCuotaMedicoAsistencial
                SumaTotal2 = SumaTotal2 + rstBuscarSlave!SumaCuotaMedicoAsistencial
                .Cells(FilaExcel + 21, ColumnaExcel + 1).Value = rstBuscarSlave!SumaDonaciones
                SumaTotal2 = SumaTotal2 + rstBuscarSlave!SumaDonaciones
                .Cells(FilaExcel + 22, ColumnaExcel + 1).Value = rstBuscarSlave!SumaMinimoNoImponible
                SumaTotal2 = SumaTotal2 + rstBuscarSlave!SumaMinimoNoImponible
                .Cells(FilaExcel + 23, ColumnaExcel + 1).Value = rstBuscarSlave!SumaConyuge
                SumaTotal2 = SumaTotal2 + rstBuscarSlave!SumaConyuge
                .Cells(FilaExcel + 24, ColumnaExcel + 1).Value = rstBuscarSlave!SumaHijo
                SumaTotal2 = SumaTotal2 + rstBuscarSlave!SumaHijo
                .Cells(FilaExcel + 25, ColumnaExcel + 1).Value = rstBuscarSlave!SumaOtrasCargas
                SumaTotal2 = SumaTotal2 + rstBuscarSlave!SumaOtrasCargas
                .Cells(FilaExcel + 26, ColumnaExcel + 1).Value = rstBuscarSlave!SumaDeduccionEspecial
                SumaTotal2 = SumaTotal2 + rstBuscarSlave!SumaDeduccionEspecial
                SumaTotal1 = SumaTotal1 - SumaTotal2
                SumaTotal2 = 0
                .Cells(FilaExcel + 28, ColumnaExcel + 1).Value = SumaTotal1
                rstBuscarSlave.Close
                If SumaTotal1 > 0 Then
                    dblAlicuotaCalculada = CalcularAlicuotaAplicable(rstListadoSlave!PuestoLaboral, _
                     CodigoLiquidacion, SumaTotal1)
                     .Cells(FilaExcel + 30, ColumnaExcel + 1).Value = dblAlicuotaCalculada * 100 & " %"
                     .Cells(FilaExcel + 31, ColumnaExcel + 1).Value = CalcularImporteFijo(rstListadoSlave!PuestoLaboral, _
                     CodigoLiquidacion, SumaTotal1)
                     .Cells(FilaExcel + 32, ColumnaExcel + 1).Value = CalcularImporteVariable(rstListadoSlave!PuestoLaboral, _
                     CodigoLiquidacion, dblAlicuotaCalculada, SumaTotal1)
'                    SQL = "Select * From ESCALAAPLICABLEGANANCIAS Order by IMPORTEMAXIMO Asc"
'                    rstBuscarSlave.Open SQL, dbSlave, adOpenDynamic, adLockOptimistic
'                    If rstBuscarSlave.BOF = False Then
'                        rstBuscarSlave.MoveFirst
'                        Do While rstBuscarSlave.EOF = False
'                            If (rstBuscarSlave!ImporteMaximo / 12 * Left(PeriodoLiquidacion, 2)) > SumaTotal1 Then
'                                .Cells(FilaExcel + 30, ColumnaExcel + 1).Value = rstBuscarSlave!ImporteVariable * 100 & " %"
'                                Exit Do
'                            End If
'                            rstBuscarSlave.MoveNext
'                        Loop
'                        If rstBuscarSlave!ImporteFijo = 0 Then
'                            .Cells(FilaExcel + 31, ColumnaExcel + 1).Value = 0
'                            .Cells(FilaExcel + 32, ColumnaExcel + 1).Value = SumaTotal1 * rstBuscarSlave!ImporteVariable
'                        Else
'                            .Cells(FilaExcel + 31, ColumnaExcel + 1).Value = (rstBuscarSlave!ImporteFijo / 12 * Left(PeriodoLiquidacion, 2))
'                            SumaTotal2 = rstBuscarSlave!ImporteVariable
'                            rstBuscarSlave.MovePrevious
'                            .Cells(FilaExcel + 32, ColumnaExcel + 1).Value = (SumaTotal1 - (rstBuscarSlave!ImporteMaximo / 12 * Left(PeriodoLiquidacion, 2))) * SumaTotal2
'                        End If
'                    End If
                    .Cells(FilaExcel + 33, ColumnaExcel + 1).Value = CDbl(.Cells(FilaExcel + 31, ColumnaExcel + 1).Value) + CDbl(.Cells(FilaExcel + 32, ColumnaExcel + 1).Value)
                Else
                    .Cells(FilaExcel + 30, ColumnaExcel + 1).Value = 0
                    .Cells(FilaExcel + 31, ColumnaExcel + 1).Value = 0
                    .Cells(FilaExcel + 32, ColumnaExcel + 1).Value = 0
                    .Cells(FilaExcel + 33, ColumnaExcel + 1).Value = 0
                End If
                SQL = "Select Sum(LIQUIDACIONSUELDOS.Importe) AS SumaDeImporte " _
                & "From CODIGOLIQUIDACIONES INNER JOIN LIQUIDACIONSUELDOS ON CODIGOLIQUIDACIONES.Codigo = LIQUIDACIONSUELDOS.CodigoLiquidacion " _
                & "Where PUESTOLABORAL = '" & rstListadoSlave!PuestoLaboral & "' And CODIGOCONCEPTO= '0276' " _
                & "And Right(PERIODO,4) = '" & Right(PeriodoLiquidacion, 4) & "' And CODIGO < '" & CodigoLiquidacion & "'"
                Set rstBuscarSlave = New ADODB.Recordset
                rstBuscarSlave.Open SQL, dbSlave, adOpenForwardOnly, adLockOptimistic
                If rstBuscarSlave.BOF = False And IsNull(rstBuscarSlave!SumaDeImporte) = False Then
                    .Cells(FilaExcel + 34, ColumnaExcel + 1).Value = CDbl(-rstBuscarSlave!SumaDeImporte)
                Else
                    .Cells(FilaExcel + 34, ColumnaExcel + 1).Value = 0
                End If
                rstBuscarSlave.Close
                .Cells(FilaExcel + 35, ColumnaExcel + 1).Value = rstListadoSlave!AjusteRetencion
                If rstListadoSlave!Retencion > 0 Then
                    .Cells(FilaExcel + 37, ColumnaExcel + 1).Value = rstListadoSlave!Retencion
                Else
                    .Cells(FilaExcel + 38, ColumnaExcel + 1).Value = rstListadoSlave!Retencion * (-1)
                End If
                SumaTotal1 = 0
                SumaTotal2 = 0
                rstListadoSlave.MoveNext
                ColumnaExcel = ColumnaExcel + 1
            Wend
            
            FilaExcel = FilaExcel + 53
        Loop Until rstListadoSlave.EOF = True
        .Columns(1).ColumnWidth = 19
        For ColumnaExcel = 2 To 15
            .Columns(ColumnaExcel).ColumnWidth = 15
        Next
    End With
    rstListadoSlave.Close
    Set rstListadoSlave = Nothing
    Set rstBuscarSlave = Nothing
    
''Configurando Para Impresión
'    With o_Libro.ActiveSheet.PageSetup
'        .LeftHeader = ""
'        .CenterHeader = ""
'        .RightHeader = ""
'        .LeftFooter = ""
'        .CenterFooter = ""
'        .RightFooter = ""
'        .LeftMargin = Application.InchesToPoints(0.393700787401575)
'        .RightMargin = Application.InchesToPoints(0.393700787401575)
'        .TopMargin = Application.InchesToPoints(0.984251968503937)
'        .BottomMargin = Application.InchesToPoints(0.984251968503937)
'        .HeaderMargin = Application.InchesToPoints(0)
'        .FooterMargin = Application.InchesToPoints(0)
'        .PrintHeadings = False
'        .PrintGridlines = False
'        .PrintComments = xlPrintNoComments
'        .PrintQuality = 600
'        .CenterHorizontally = True
'        .CenterVertically = False
'        .Orientation = xlLandscape
'        .Draft = False
'        .PaperSize = xlPaperLegal
'        .FirstPageNumber = xlAutomatic
'        .Order = xlDownThenOver
'        .BlackAndWhite = False
'        .Zoom = 70
'        .PrintErrors = xlPrintErrorsDisplayed
'    End With


    o_Libro.Close True, strOutputPath
    ' -- Cerrar Excel
    o_Excel.Quit
    ' -- Terminar instancias
    Call ReleaseObjects(o_Excel, o_Libro, o_Hoja)
    GenerarReporteMensualGanancias = True
Exit Function
  
' -- Controlador de Errores
Error_Handler:
    ' -- Cierra la hoja y el la aplicación Excel
    If Not o_Libro Is Nothing Then: o_Libro.Close False
    If Not o_Excel Is Nothing Then: o_Excel.Quit
    Call ReleaseObjects(o_Excel, o_Libro, o_Hoja)
    If Err.Number <> 1004 Then MsgBox Err.Description, vbCritical
End Function

Public Function GenerarArchivoSICORE(OutputPath As String, CodigoLiquidacion As String, _
Periodo As String, Fecha As String, SignoDecimal As String) As Boolean

    Dim intNumeroArchivo        As Integer
    Dim strCodigoComprobante    As String
    Dim strFechaComprobante     As String
    Dim strNroComprobante       As String
    Dim strImporteComprobante   As String 'Buscar con recordset
    Dim strCodigoImpuesto       As String
    Dim strCodigoRegimen        As String
    Dim strCodigoOperacion      As String
    Dim strBaseCalculo          As String
    Dim strFechaRetencion       As String
    Dim strCodigoCondicion      As String
    Dim strRetSujetosSusp       As String
    Dim strImporteRetencion     As String
    Dim strPorcentajeExclusion  As String
    Dim strFechaBoletin         As String
    Dim strTipoDocumento        As String
    Dim strCUIL                 As String 'Buscar con recordset
    Dim strNroCertificado       As String
    Dim strDenominacionOrd      As String
    Dim strAcrecentamiento      As String
    Dim strCUITPaisRet          As String
    Dim strCUITOrdenante        As String
    Dim strCadenaCompleta       As String
    Dim SQL                     As String
    
    'Carga inicial de las variables repetitivas
    strCodigoComprobante = "07" 'Código para Recibo de Sueldo
    strFechaComprobante = Fecha
    strNroComprobante = "000000" & Right(Periodo, 4) & Left(Periodo, 2) & Space$(4) 'Código definido por Belén B.
    strImporteComprobante = Space$(12) & "0" & SignoDecimal & "00" 'Importe en 0
    strCodigoImpuesto = "787"
    strCodigoRegimen = "160"
    strCodigoOperacion = "1"
    strBaseCalculo = Space$(10) & "0" & SignoDecimal & "00" 'Importe en 0
    strFechaRetencion = Fecha
    strCodigoCondicion = "01" 'Responsables Inscriptos
    strRetSujetosSusp = "0"
    strPorcentajeExclusion = Space$(2) & "0" & SignoDecimal & "00" 'Importe en 0
    strFechaBoletin = Space$(10)
    strTipoDocumento = "86" 'Códigoo para CUIL
    strNroCertificado = String$(14, "0")
    strDenominacionOrd = Space$(30)
    strAcrecentamiento = "0"
    strCUITPaisRet = Space$(11)
    strCUITOrdenante = Space$(11)
    
    'Iniciamos la carga de los datos
    intNumeroArchivo = FreeFile
    Open OutputPath For Output As #intNumeroArchivo
    
    'Generamos un bucle
    Set rstListadoSlave = New ADODB.Recordset
    SQL = "Select CUIL, Retencion From AGENTES Inner Join LIQUIDACIONGANANCIAS4TACATEGORIA " _
        & "On AGENTES.PuestoLaboral = LIQUIDACIONGANANCIAS4TACATEGORIA.PuestoLaboral " _
        & "Where CodigoLiquidacion = '" & CodigoLiquidacion & "' " _
        & "And Retencion > 0 " _
        & "Order By Retencion Desc"
    rstListadoSlave.Open SQL, dbSlave, adOpenDynamic, adLockOptimistic
    If rstListadoSlave.BOF = False Then
        rstListadoSlave.MoveFirst
        While rstListadoSlave.EOF = False
            strImporteRetencion = FormatNumber(rstListadoSlave!Retencion, 2, , , vbFalse)
            strImporteRetencion = Left(strImporteRetencion, Len(strImporteRetencion) - 3) _
            & SignoDecimal & Right(strImporteRetencion, 2)
            strImporteRetencion = Space$(14 - Len(strImporteRetencion)) & strImporteRetencion
            strCUIL = rstListadoSlave!CUIL & Space$(9)
            strCadenaCompleta = strCodigoComprobante & strFechaComprobante & strNroComprobante _
                                & strImporteComprobante & strCodigoImpuesto & strCodigoRegimen _
                                & strCodigoOperacion & strBaseCalculo & strFechaRetencion _
                                & strCodigoCondicion & strRetSujetosSusp & strImporteRetencion _
                                & strPorcentajeExclusion & strFechaBoletin & strTipoDocumento _
                                & strCUIL & strNroCertificado & strDenominacionOrd _
                                & strAcrecentamiento & strCUITPaisRet & strCUITOrdenante
            Print #intNumeroArchivo, strCadenaCompleta
            rstListadoSlave.MoveNext
        Wend
    End If

    Close #intNumeroArchivo
    rstListadoSlave.Close
    Set rstListadoSlave = Nothing
    Unload ExportacionSICORE
    GenerarArchivoSICORE = True

End Function

Public Function GenerarArchivoLiquidacionHonorarios(OutputPath As String, Año As String, _
SignoDecimal As String) As Boolean

    Dim intNumeroArchivo        As Integer
    Dim strImporte              As String 'Buscar con recordset
    Dim strCadenaCompleta       As String
    Dim SQL                     As String
     
    'Iniciamos la carga de los datos
    intNumeroArchivo = FreeFile
    Open OutputPath For Output As #intNumeroArchivo
    
    'Generamos un bucle
    Set rstListadoSlave = New ADODB.Recordset
    SQL = "Select * From LIQUIDACIONHONORARIOS " _
        & "Where Year(Fecha) = '" & Año & "' " _
        & "Order By Comprobante Asc"
    rstListadoSlave.Open SQL, dbSlave, adOpenDynamic, adLockOptimistic
    If rstListadoSlave.BOF = False Then
        rstListadoSlave.MoveFirst
        While rstListadoSlave.EOF = False
            'Inicializamos la cadena (No hace falta, pero por las dudas...)
            strCadenaCompleta = ""
            'Cargamos Fecha
            strCadenaCompleta = Chr(34) & CStr(rstListadoSlave!Fecha) & Chr(34) & ","
            'Cargamos Proveedor
            strCadenaCompleta = strCadenaCompleta & Chr(34) & CStr(rstListadoSlave!Proveedor) & Chr(34) & ","
            'Cargamos Importe Sellos
            strImporte = FormatNumber(rstListadoSlave!Sellos, 2, , , vbFalse)
            strImporte = Left(strImporte, Len(strImporte) - 3) _
            & SignoDecimal & Right(strImporte, 2)
            strCadenaCompleta = strCadenaCompleta & Chr(34) & CStr(strImporte) & Chr(34) & ","
            'Cargamos Importe Seguro
            strImporte = FormatNumber(rstListadoSlave!Seguro, 2, , , vbFalse)
            strImporte = Left(strImporte, Len(strImporte) - 3) _
            & SignoDecimal & Right(strImporte, 2)
            strCadenaCompleta = strCadenaCompleta & Chr(34) & CStr(strImporte) & Chr(34) & ","
            'Cargamos Comprobante
            strCadenaCompleta = strCadenaCompleta & Chr(34) & CStr(rstListadoSlave!Comprobante) & Chr(34) & ","
            'Cargamos Tipo
            strCadenaCompleta = strCadenaCompleta & Chr(34) & CStr(rstListadoSlave!Tipo) & Chr(34) & ","
            'Cargamos Importe Bruto
            strImporte = FormatNumber(rstListadoSlave!MontoBruto, 2, , , vbFalse)
            strImporte = Left(strImporte, Len(strImporte) - 3) _
            & SignoDecimal & Right(strImporte, 2)
            strCadenaCompleta = strCadenaCompleta & Chr(34) & CStr(strImporte) & Chr(34) & ","
            'Cargamos Importe IIBB
            strImporte = FormatNumber(rstListadoSlave!IIBB, 2, , , vbFalse)
            strImporte = Left(strImporte, Len(strImporte) - 3) _
            & SignoDecimal & Right(strImporte, 2)
            strCadenaCompleta = strCadenaCompleta & Chr(34) & CStr(strImporte) & Chr(34) & ","
            'Cargamos Importe Libramiento Pago
            strImporte = FormatNumber(rstListadoSlave!LibramientoPago, 2, , , vbFalse)
            strImporte = Left(strImporte, Len(strImporte) - 3) _
            & SignoDecimal & Right(strImporte, 2)
            strCadenaCompleta = strCadenaCompleta & Chr(34) & CStr(strImporte) & Chr(34) & ","
            'Cargamos Importe Otra Retencion
            strImporte = FormatNumber(rstListadoSlave!OtraRetencion, 2, , , vbFalse)
            strImporte = Left(strImporte, Len(strImporte) - 3) _
            & SignoDecimal & Right(strImporte, 2)
            strCadenaCompleta = strCadenaCompleta & Chr(34) & CStr(strImporte) & Chr(34) & ","
            'Cargamos Importe Anticipo
            strImporte = FormatNumber(rstListadoSlave!Anticipo, 2, , , vbFalse)
            strImporte = Left(strImporte, Len(strImporte) - 3) _
            & SignoDecimal & Right(strImporte, 2)
            strCadenaCompleta = strCadenaCompleta & Chr(34) & CStr(strImporte) & Chr(34) & ","
            'Cargamos Importe Descuento
            strImporte = FormatNumber(rstListadoSlave!Descuento, 2, , , vbFalse)
            strImporte = Left(strImporte, Len(strImporte) - 3) _
            & SignoDecimal & Right(strImporte, 2)
            strCadenaCompleta = strCadenaCompleta & Chr(34) & CStr(strImporte) & Chr(34) & ","
            'Cargamos Actividad
            strCadenaCompleta = strCadenaCompleta & Chr(34) & CStr(rstListadoSlave!ACTIVIDAD) & Chr(34) & ","
            'Cargamos Partida
            strCadenaCompleta = strCadenaCompleta & Chr(34) & CStr(rstListadoSlave!PARTIDA) & Chr(34)
            
            'Cargamos el registro
            Print #intNumeroArchivo, strCadenaCompleta
            rstListadoSlave.MoveNext
        Wend
    End If

    Close #intNumeroArchivo
    rstListadoSlave.Close
    Set rstListadoSlave = Nothing
    Unload ExportacionSLAVE
    GenerarArchivoLiquidacionHonorarios = True

End Function

Public Function GenerarArchivoPrecarizados(OutputPath As String) As Boolean

    Dim intNumeroArchivo        As Integer
    Dim strImporte              As String 'Buscar con recordset
    Dim strCadenaCompleta       As String
    Dim SQL                     As String
     
    'Iniciamos la carga de los datos
    intNumeroArchivo = FreeFile
    Open OutputPath For Output As #intNumeroArchivo
    
    'Generamos un bucle
    Set rstListadoSlave = New ADODB.Recordset
    SQL = "Select * From PRECARIZADOS " _
        & "Order By Agentes Asc"
    rstListadoSlave.Open SQL, dbSlave, adOpenDynamic, adLockOptimistic
    If rstListadoSlave.BOF = False Then
        rstListadoSlave.MoveFirst
        While rstListadoSlave.EOF = False
            'Inicializamos la cadena (No hace falta, pero por las dudas...)
            strCadenaCompleta = ""
            'Cargamos Agentes
            strCadenaCompleta = Chr(34) & CStr(rstListadoSlave!AGENTES) & Chr(34) & ","
            'Cargamos Actividad
            strCadenaCompleta = strCadenaCompleta & Chr(34) & CStr(rstListadoSlave!ACTIVIDAD) & Chr(34) & ","
            'Cargamos Partida
            strCadenaCompleta = strCadenaCompleta & Chr(34) & CStr(rstListadoSlave!PARTIDA) & Chr(34)
            
            'Cargamos el registro
            Print #intNumeroArchivo, strCadenaCompleta
            rstListadoSlave.MoveNext
        Wend
    End If

    Close #intNumeroArchivo
    rstListadoSlave.Close
    Set rstListadoSlave = Nothing
    Unload ExportacionSLAVE
    GenerarArchivoPrecarizados = True

End Function

