VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form ListadoPrecarizados 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Precarizados"
   ClientHeight    =   7275
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   5415
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7275
   ScaleWidth      =   5415
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Caption         =   "Listado de Precarizados"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   4935
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   5175
      Begin MSMask.MaskEdBox mskEdicion 
         Height          =   285
         Left            =   1680
         TabIndex        =   5
         Top             =   2520
         Visible         =   0   'False
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   503
         _Version        =   393216
         BorderStyle     =   0
         Appearance      =   0
         BackColor       =   12648384
         MaxLength       =   12
         Mask            =   "##-##-##-###"
         PromptChar      =   "_"
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid dgPrecarizados 
         Height          =   4455
         Left            =   120
         TabIndex        =   4
         Top             =   360
         Width           =   4900
         _ExtentX        =   8652
         _ExtentY        =   7858
         _Version        =   393216
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Acciones Posibles"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   2055
      Left            =   240
      TabIndex        =   0
      Top             =   5160
      Width           =   5010
      Begin VB.CommandButton cmdExportar 
         BackColor       =   &H008080FF&
         Caption         =   "Exportar Precarizados"
         Height          =   375
         Left            =   1320
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   1560
         Width           =   2295
      End
      Begin VB.CommandButton cmdImportar 
         BackColor       =   &H008080FF&
         Caption         =   "Importar Precarizados"
         Enabled         =   0   'False
         Height          =   375
         Left            =   2640
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   960
         Width           =   2175
      End
      Begin VB.CommandButton cmdEliminar 
         BackColor       =   &H008080FF&
         Caption         =   "Eliminar Precarizado"
         Height          =   375
         Left            =   2640
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   360
         Width           =   2175
      End
      Begin VB.CommandButton cmdAgregar 
         BackColor       =   &H008080FF&
         Caption         =   "Agregar Precarizado"
         Height          =   375
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   360
         Width           =   2175
      End
      Begin VB.CommandButton cmdEditar 
         BackColor       =   &H008080FF&
         Caption         =   "Editar Precarizado"
         Height          =   375
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   960
         Width           =   2175
      End
   End
End
Attribute VB_Name = "ListadoPrecarizados"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
' Variable para la clase
Dim EditarGrilla As CEditarFlexGrid

Private Sub cmdAgregar_Click()

    CargaPrecarizado.Show
    Unload ListadoPrecarizados

End Sub

Private Sub cmdEditar_Click()

    EditarPrecarizado

End Sub

Private Sub cmdEliminar_Click()

    EliminarPrecarizado

End Sub

Private Sub cmdExportar_Click()

    Dim strNombreArchivo As String
   
    strNombreArchivo = "\Listado Factureros de SLAVE.xls"
        
    If Exportar_Excel(App.Path & strNombreArchivo, Me.dgPrecarizados) Then
        MsgBox " Datos exportados en " & App.Path, vbInformation
    End If

End Sub

Private Sub Form_Load()

    Call CenterMe(ListadoPrecarizados, 5500, 7700)
    'Nueva instancia
    Set EditarGrilla = New CEditarFlexGrid

    'Inicia ( se le envia el Flex y el text )
    EditarGrilla.Iniciar dgPrecarizados, mskEdicion

End Sub

'Option Explicit
'' Variable para la clase
'Dim EditarGrilla As CEditarFlexGrid
'
'Private Sub cmdAgregar_Click()
'    Dim X As Integer
'    If Validar = True Then
'        With dgLaboratorio
'            If rstDatosCargaHistorial.BOF = False Then
'                rstDatosCargaHistorial.MoveFirst
'                For X = 1 To (.Rows - 1)
'                    If Trim(.TextMatrix(X, 3)) <> "" Then
'                        Dim ValorBuscado As String
'                        ValorBuscado = Format(.TextMatrix(X, 0), "'&&&&&&&&&&&&&&&&&&&&&&&" _
'                        & "&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&" _
'                        & "&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&" _
'                        & "&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&'")
'                        rstDatosCargaHistorial.FindFirst "CodigoHistorial = " & ValorBuscado
'                        If rstDatosCargaHistorial.NoMatch = True Then
'                            rstDatosCargaHistorial.AddNew
'                            rstDatosCargaHistorial.Fields("NumeroIngreso") = rstDatosConsultaHistorial.Fields("NumeroIngreso")
'                            rstDatosCargaHistorial.Fields("CodigoHistorial") = .TextMatrix(X, 0)
'                        Else
'                            rstDatosCargaHistorial.Edit
'                        End If
'                            rstDatosCargaHistorial.Fields("FechaDesde") = txtFechaLaboratorio.Text
'                            rstDatosCargaHistorial.Fields("Dato") = .TextMatrix(X, 3)
'                            rstDatosCargaHistorial.Update
'                    End If
'                Next X
'                ValorBuscado = ""
'            Else
'                Call SetRecordset(rstCargaHistorial, "HISTORIAL")
'                For X = 1 To (.Rows - 1)
'                    If Trim(.TextMatrix(X, 3)) <> "" Then
'                        rstCargaHistorial.AddNew
'                        rstCargaHistorial.Fields("NumeroIngreso") = rstDatosConsultaHistorial.Fields("NumeroIngreso")
'                        rstCargaHistorial.Fields("CodigoHistorial") = .TextMatrix(X, 0)
'                        rstCargaHistorial.Fields("FechaDesde") = txtFechaLaboratorio.Text
'                        rstCargaHistorial.Fields("Dato") = .TextMatrix(X, 3)
'                        rstCargaHistorial.Update
'                    End If
'                Next X
'                Set rstCargaHistorial = Nothing
'            End If
'        End With
'        strBuscarPaciente = rstDatosConsultaHistorial.Fields("DNI")
'        Unload Laboratorio
'        ListadoConsultas.Show
'    End If
'    X = 0
'End Sub
'
'Private Sub cmdEliminarTodos_Click()
'    If rstDatosCargaHistorial.BOF = False Then
'        Dim Borrar As Integer
'        Borrar = MsgBox("Desea Borrar DEFINITIVAMENTE todos los datos de LABORATORIO de la presente CONSULTA?", vbQuestion + vbYesNo, "BORRANDO HISTORIAL")
'        If Borrar = 6 Then
'            With rstDatosCargaHistorial
'                .MoveFirst
'                Do Until .EOF = True
'                    .Delete
'                    .MoveNext
'                Loop
'            End With
'            ConfigurarLaboratorio
'            CargarLaboratorio
'        End If
'        Borrar = 0
'    Else
'        MsgBox "La consulta NO POSEE datos para eliminar", vbInformation + vbOKOnly, "IMPOSIBLE ELIMINAR"
'    End If
'End Sub
'
'Private Sub cmdEliminarUno_Click()
'    Dim X As Integer
'    X = dgLaboratorio.Row
'    If dgLaboratorio.TextMatrix(X, 0) <> "SECCION" Then
'        Dim ValorBuscado As String
'        ValorBuscado = Format(dgLaboratorio.TextMatrix(X, 0), "'&&&&&&&&&&&&&&&&&&&&&&&" _
'        & "&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&" _
'        & "&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&" _
'        & "&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&'")
'        rstDatosCargaHistorial.FindFirst "CodigoHistorial = " & ValorBuscado
'        If rstDatosCargaHistorial.NoMatch = True Then
'            MsgBox "El VALOR que intenta borrar se encuentra vacio" & vbCrLf & "Verifique si la fila seleccionada es la correcta", vbInformation + vbOKOnly, "IMPOSIBLE ELIMINAR"
'        Else
'            Dim Borrar As Integer
'            Borrar = MsgBox("Desea Borrar DEFINITIVAMENTE la VARIABLE: " & dgLaboratorio.TextMatrix(X, 1) & " ?", vbQuestion + vbYesNo, "BORRANDO HISTORIAL")
'            If Borrar = 6 Then
'                rstDatosCargaHistorial.Delete
'                ConfigurarLaboratorio
'                CargarLaboratorio
'            End If
'            Borrar = 0
'        End If
'        ValorBuscado = ""
'    End If
'    X = 0
'End Sub
'
'Private Sub Form_Load()
'    Call CenterMe(Laboratorio)
'    With Laboratorio
'        .Width = 4700
'        .Height = 9000
'    End With
'    With rstDatosConsultaHistorial
'        txtApellidoyNombre.Text = .Fields("Apellido") & ", " & .Fields("Nombre")
'        txtFechaConsulta.Text = .Fields("Fecha")
'        If rstDatosCargaHistorial.BOF = False Then
'            rstDatosCargaHistorial.MoveFirst
'            txtFechaLaboratorio.Text = rstDatosCargaHistorial.Fields("FechaDesde")
'        Else
'            txtFechaLaboratorio.Text = .Fields("Fecha")
'        End If
'    End With
'    ConfigurarLaboratorio
'    CargarLaboratorio
'    'Nueva instancia
'    Set EditarGrilla = New CEditarFlexGrid
'
'    'Inicia ( se le envia el Flex y el text )
'    EditarGrilla.Iniciar dgLaboratorio, txtEdicion
'End Sub
'
'Sub ConfigurarLaboratorio()
'    With dgLaboratorio
'        .Clear
'        .Cols = 4
'        .Rows = 2
'        .TextMatrix(0, 0) = "Código Historial"
'        .TextMatrix(0, 1) = "Descripción"
'        .TextMatrix(0, 2) = "TipoDatos"
'        .TextMatrix(0, 3) = "Datos"
'        .ColWidth(0) = 1
'        .ColWidth(1) = 2500
'        .ColWidth(2) = 1
'        .ColWidth(3) = 1250
'        .FixedCols = 3
'        .FocusRect = flexFocusHeavy
'        .HighLight = flexHighlightWithFocus
'        .AllowUserResizing = flexResizeColumns
'        .ColAlignment(0) = 1
'        .ColAlignment(1) = 1
'        .ColAlignment(2) = 1
'        .ColAlignment(3) = 1
'    End With
'End Sub
'
'Sub CargarLaboratorio()
'    Dim i As Integer
'    i = 1
'    dgLaboratorio.Rows = 2
'    dgLaboratorio.RowHeight(i) = 300
'    dgLaboratorio.TextMatrix(i, 0) = "SECCION"
'    dgLaboratorio.TextMatrix(i, 1) = "LABORATORIO"
'    dgLaboratorio.Rows = dgLaboratorio.Rows + 1
'    Call SetRecordset(rstListadoHistorialPrincipal, "SELECT * FROM TipoHistorial Where CodigoHistorial LIKE 'L###' AND Jerarquia = 'Principal' ORDER BY CodigoHistorial")
'    If rstListadoHistorialPrincipal.BOF = False Then
'        With rstListadoHistorialPrincipal
'            .MoveFirst
'            While .EOF = False
'                i = i + 1
'                dgLaboratorio.RowHeight(i) = 300
'                dgLaboratorio.TextMatrix(i, 0) = .Fields("CodigoHistorial")
'                dgLaboratorio.TextMatrix(i, 1) = " + " & .Fields("Descripcion")
'                dgLaboratorio.TextMatrix(i, 2) = .Fields("TipoDato")
'                dgLaboratorio.Rows = dgLaboratorio.Rows + 1
'                Call SetRecordset(rstListadoHistorialAccesorio, "SELECT * FROM TipoHistorial Where CodigoHistorial LIKE 'L###' AND Jerarquia = " & "'" & rstListadoHistorialPrincipal.Fields("Descripcion") & "'")
'                If rstListadoHistorialAccesorio.BOF = False Then
'                    With rstListadoHistorialAccesorio
'                        .MoveFirst
'                        While .EOF = False
'                            i = i + 1
'                            dgLaboratorio.RowHeight(i) = 300
'                            dgLaboratorio.TextMatrix(i, 0) = .Fields("CodigoHistorial")
'                            dgLaboratorio.TextMatrix(i, 1) = "     - " & .Fields("Descripcion")
'                            dgLaboratorio.TextMatrix(i, 2) = .Fields("TipoDato")
'                            dgLaboratorio.Rows = dgLaboratorio.Rows + 1
'                            .MoveNext
'                        Wend
'                    End With
'                End If
'                .MoveNext
'            Wend
'        End With
'    End If
'    dgLaboratorio.Rows = dgLaboratorio.Rows - 1
'    Set rstListadoHistorialPrincipal = Nothing
'    Set rstListadoHistorialAccesorio = Nothing
'    If rstDatosCargaHistorial.BOF = False Then
'        With rstDatosCargaHistorial
'            .MoveFirst
'            While .EOF = False
'            Dim X As Integer
'                For X = 1 To (dgLaboratorio.Rows - 1)
'                    If dgLaboratorio.TextMatrix(X, 0) = .Fields("CodigoHistorial") Then
'                        dgLaboratorio.TextMatrix(X, 3) = .Fields("Dato")
'                        Exit For
'                    End If
'                Next X
'                X = 0
'                .MoveNext
'            Wend
'        End With
'    End If
'End Sub
'
'Private Sub Form_Unload(Cancel As Integer)
'    strBuscarPaciente = rstDatosConsultaHistorial.Fields("DNI")
'    Set EditarGrilla = Nothing
'    Set rstDatosConsultaHistorial = Nothing
'    Set rstDatosCargaHistorial = Nothing
'    ListadoConsultas.Show
'End Sub
'
'Function Validar() As Boolean
'    Dim X As Integer
'    With dgLaboratorio
'        For X = 1 To (.Rows - 1)
'            If .TextMatrix(X, 0) = "SECCION" Then
'                .TextMatrix(X, 3) = ""
'            Else
'                Select Case .TextMatrix(X, 2)
'                Case Is = "Ninguno"
'                    If Trim(.TextMatrix(X, 3)) <> "" Then
'                        MsgBox "La Variable " & .TextMatrix(X, 1) & " no puedo guardar datos por ser su propiedad igual a NINGUNO", vbOKOnly, "IMPOSIBLE CARGAR DATOS"
'                        .TextMatrix(X, 3) = ""
'                    End If
'                Case Is = "Numero"
'                    If Trim(.TextMatrix(X, 3)) <> "" Then
'                        If EsNumeroNoVacio(.TextMatrix(X, 3), "30", .TextMatrix(X, 1)) = False Then
'                            Validar = False
'                            .SetFocus
'                            .Row = X
'                            .col = 3
'                            Exit Function
'                        End If
'                    End If
'                Case Is = "Texto"
'                    If Trim(.TextMatrix(X, 3)) <> "" Then
'                        If EsTextoNoVacio(.TextMatrix(X, 3), "30", .TextMatrix(X, 1)) = False Then
'                            Validar = False
'                            .SetFocus
'                            .Row = X
'                            .col = 3
'                            Exit Function
'                        End If
'                    End If
'                Case Is = "Fecha"
'                    If Trim(.TextMatrix(X, 3)) <> "" Then
'                        If EsFechaNoVacio(.TextMatrix(X, 3), "30", .TextMatrix(X, 1)) = False Then
'                            Validar = False
'                            .SetFocus
'                            .Row = X
'                            .col = 3
'                            Exit Function
'                        End If
'                    End If
'                End Select
'            End If
'        Next X
'    End With
'    If EsFechaNoVacio(txtFechaLaboratorio.Text, "20", "FECHA LABORATORIO") = False Then
'        txtFechaLaboratorio.SetFocus
'        Validar = False
'        Exit Function
'    End If
'    X = 0
'    Validar = True
'
'End Function

Private Sub Form_Unload(Cancel As Integer)

    Set EditarGrilla = Nothing

End Sub
