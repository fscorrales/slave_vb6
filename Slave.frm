VERSION 5.00
Begin VB.MDIForm Slave 
   BackColor       =   &H8000000C&
   Caption         =   "SLAVE"
   ClientHeight    =   3765
   ClientLeft      =   225
   ClientTop       =   855
   ClientWidth     =   5610
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Menu MnuArchivo 
      Caption         =   "&Archivo"
      Begin VB.Menu MnuCargaAgente 
         Caption         =   "Carga Agente"
         Begin VB.Menu MnuAgregarAgente 
            Caption         =   "Agregar Agente"
         End
         Begin VB.Menu MnuListadoAgentes 
            Caption         =   "Listado de Agentes"
         End
      End
      Begin VB.Menu MnuCargaConcepto 
         Caption         =   "Carga Concepto"
         Begin VB.Menu MnuAgregarConcepto 
            Caption         =   "Agregar Concepto"
         End
         Begin VB.Menu MnuListadoConceptos 
            Caption         =   "Listado de Conceptos"
         End
      End
      Begin VB.Menu MnuParestesco 
         Caption         =   "Carga Parentesco"
         Begin VB.Menu MnuAgregarParentesco 
            Caption         =   "Agregar Parentesco"
         End
         Begin VB.Menu MnuListadoParentesco 
            Caption         =   "Listado de Parentesco"
         End
      End
      Begin VB.Menu Line01 
         Caption         =   "-"
      End
      Begin VB.Menu MnuGanancias4taCategoria 
         Caption         =   "Gcias.4ta Categoría"
         Begin VB.Menu MnuEscalaAplicable 
            Caption         =   "Escala Aplicable"
         End
         Begin VB.Menu MnuLimitesDeducciones 
            Caption         =   "Límites Deducciones"
         End
      End
      Begin VB.Menu MnuSIRADIG 
         Caption         =   "SIRADIG"
         Begin VB.Menu MnuAgregarParentescoSIRADIG 
            Caption         =   "Agregar Parentesco"
         End
         Begin VB.Menu MnuListadoParentescoSIRADIG 
            Caption         =   "Listado Parentesco"
         End
         Begin VB.Menu Line02 
            Caption         =   "-"
         End
         Begin VB.Menu MnuAgregarDeduccionesSIRADIG 
            Caption         =   "Agregar Deducciones"
         End
         Begin VB.Menu MnuListadoDeduccionesSIRADIG 
            Caption         =   "Listado Deducciones"
         End
         Begin VB.Menu Line03 
            Caption         =   "-"
         End
         Begin VB.Menu MnuAgregarOtrasDeduccionesSIRADIG 
            Caption         =   "Agregar Otras Deducciones"
         End
         Begin VB.Menu MnuListadoOtrasDeduccionesSIRADIG 
            Caption         =   "Listado Otras Deducciones"
         End
      End
      Begin VB.Menu Line04 
         Caption         =   "-"
      End
      Begin VB.Menu MnuCargaPrecarizado 
         Caption         =   "Carga Precarizado"
         Begin VB.Menu MnuAgregarPrecarizado 
            Caption         =   "Agregar Precarizado"
         End
         Begin VB.Menu MnuListadoPrecarizados 
            Caption         =   "Listado de  Precarizados"
         End
         Begin VB.Menu Line05 
            Caption         =   "-"
         End
         Begin VB.Menu MnuImportarPrecarizados 
            Caption         =   "Importación SLAVE"
         End
      End
   End
   Begin VB.Menu MnuFamiliares 
      Caption         =   "&Familiares"
      Begin VB.Menu MnuListadoFamiliares 
         Caption         =   "Listado Familiares"
      End
      Begin VB.Menu MnuImportacionSISPER 
         Caption         =   "Importación SISPER (.csv)"
      End
   End
   Begin VB.Menu MnuGanancias 
      Caption         =   "&Ganancias"
      Begin VB.Menu MnuLiquidacionGanancias 
         Caption         =   "Liquidación"
         Begin VB.Menu MnuLiquidacionMensualGanancias 
            Caption         =   "Liquidación Mensual"
         End
         Begin VB.Menu MnuLiquidacionFinalGanancias 
            Caption         =   "Liquidación Anual / Final"
         End
      End
      Begin VB.Menu MnuListadoPorAgente 
         Caption         =   "Listado Por Agente"
      End
      Begin VB.Menu MnuListadoPorPeriodo 
         Caption         =   "Listado Por Período"
      End
      Begin VB.Menu MnuCargaDeducciones 
         Caption         =   "Carga Deducciones"
         Begin VB.Menu MnuDeduccionesPersonales 
            Caption         =   "Deducciones Personales (Viejo)"
         End
         Begin VB.Menu MnuDeduccionesGenerales 
            Caption         =   "Deducciones Generales (Viejo)"
         End
         Begin VB.Menu MnuF572 
            Caption         =   "F. 572 Presentados"
         End
         Begin VB.Menu Line06 
            Caption         =   "-"
         End
         Begin VB.Menu MnuImportarF572Web 
            Caption         =   "Importar F.572 web (.xml)"
         End
         Begin VB.Menu MnuMigrarDeducciones 
            Caption         =   "Migrar Deducciones"
         End
      End
   End
   Begin VB.Menu MnuSueldo 
      Caption         =   "&Sueldo"
      Begin VB.Menu MnuCodigoLiquidaciones 
         Caption         =   "Código Liquidaciones"
         Begin VB.Menu MnuAgregarCodigo 
            Caption         =   "Agregar Código"
         End
         Begin VB.Menu MnuListadoCodigos 
            Caption         =   "Listado de Códigos"
         End
      End
      Begin VB.Menu MnuLiquidacionesSISPER 
         Caption         =   "Liquidaciones SISPER"
         Begin VB.Menu MnuImportarSueldoSISPER 
            Caption         =   "Importar Sueldo"
         End
         Begin VB.Menu MnuIncorporarConceptoSISPER 
            Caption         =   "Incorporar Concepto"
         End
         Begin VB.Menu MnuCopiarLiquidacionSISPER 
            Caption         =   "Copiar Liquidación"
         End
         Begin VB.Menu MnuEliminarLiquidacionSISPER 
            Caption         =   "Eliminar Liquidación"
         End
         Begin VB.Menu MnuLiquidacionPruebaSISPER 
            Caption         =   "Liquidación Prueba"
         End
      End
      Begin VB.Menu Line07 
         Caption         =   "-"
      End
      Begin VB.Menu MnuReciboDeSueldo 
         Caption         =   "Recibo de Sueldo"
      End
   End
   Begin VB.Menu MnuHonorarios 
      Caption         =   "&Honorarios"
      Begin VB.Menu MnuImportarHonorarios 
         Caption         =   "Importación G.Fciera. (.csv)"
      End
      Begin VB.Menu MnuExportarSLAVE 
         Caption         =   "Exportación SLAVE"
      End
   End
   Begin VB.Menu MnuSIIF 
      Caption         =   "&SIIF"
      Begin VB.Menu MnuCargaComprobante 
         Caption         =   "Carga Comprobante"
         Begin VB.Menu MnuAgregarComprobante 
            Caption         =   "Agregar Comprobante"
         End
         Begin VB.Menu MnuListadoComprobantes 
            Caption         =   "Listado de Comprobantes"
         End
         Begin VB.Menu Line08 
            Caption         =   "-"
         End
         Begin VB.Menu MnuImportarComprobantesSIIF 
            Caption         =   "Importación SLAVE"
         End
      End
   End
End
Attribute VB_Name = "Slave"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub MDIForm_Load()

    Dim strPassword As String
    Dim datFechaLimite As Date
     
    datFechaLimite = #3/15/2019#
     
'    If Date < datFechaLimite Then
        Conectar
'    End If
    
    strPassword = InputBox("Escriba la contraseña correcta o ingrese sin ella en forma restringida", "Contraseña ADMINISTRADOR")
    If strPassword <> "kanou25" Then
        MnuListadoFamiliares.Enabled = False
        MnuImportacionSISPER.Enabled = False
        MnuLiquidacionMensualGanancias.Enabled = False
        MnuLiquidacionFinalGanancias.Enabled = False
        MnuListadoPorAgente.Enabled = False
        MnuListadoPorPeriodo.Enabled = False
        MnuDeduccionesPersonales.Enabled = False
        MnuDeduccionesGenerales.Enabled = False
        MnuAgregarCodigo = False
        MnuListadoCodigos = False
        MnuMigrarDeducciones = False
        MnuImportarSueldoSISPER = False
        MnuIncorporarConceptoSISPER = False
        MnuCopiarLiquidacionSISPER = False
        MnuEliminarLiquidacionSISPER = False
        MnuLiquidacionPruebaSISPER = False
        MnuReciboDeSueldo = False
        MnuImportarF572Web = False
        MnuImportarComprobantesSIIF = False
        MnuImportarPrecarizados = False
        MnuAgregarParentescoSIRADIG = False
        MnuListadoParentescoSIRADIG = False
        MnuAgregarDeduccionesSIRADIG = False
        MnuListadoDeduccionesSIRADIG = False
        MnuAgregarOtrasDeduccionesSIRADIG = False
        MnuListadoOtrasDeduccionesSIRADIG = False
    End If

End Sub

Private Sub MDIForm_Unload(Cancel As Integer)

    VaciarTodasLasVariables
    Desconectar

End Sub

Private Sub MnuAgregarAgente_Click()

    CargaAgente.Show

End Sub

Private Sub MnuAgregarCodigo_Click()

    CargaCodigoLiquidacion.Show

End Sub

Private Sub MnuAgregarComprobante_Click()

    Autocarga.Show
    ConfigurardgAutocarga
    CargardgAutocarga

End Sub

Private Sub MnuAgregarConcepto_Click()

    CargaConcepto.Show

End Sub

Private Sub MnuAgregarParentesco_Click()

    CargaParentesco.Show

End Sub

Private Sub MnuAgregarParentescoSIRADIG_Click()

    strCargaCodigoSIRADIG = "Parentesco"
    CargaCodigoSIRADIG.Show
    
End Sub

Private Sub MnuAgregarPrecarizado_Click()

    CargaPrecarizado.Show

End Sub

Private Sub MnuCopiarLiquidacionSISPER_Click()

    CopiarLiquidacionSisper.Show
    CargarCmbCopiarLiquidacionSisper

End Sub

Private Sub MnuDeduccionesGenerales_Click()

    ListadoDeduccionesGenerales.Show
    Call ConfigurardgAgentes(ListadoDeduccionesGenerales)
    Call CargardgAgentes(ListadoDeduccionesGenerales)
    ConfigurardgDeduccionesGenerales
    Call CargardgDeduccionesGenerales(ListadoDeduccionesGenerales.dgAgentes.TextMatrix(1, 2))

End Sub

Private Sub MnuDeduccionesPersonales_Click()

    ListadoDeduccionesPersonales.Show
    Call ConfigurardgAgentes(ListadoDeduccionesPersonales)
    Call CargardgAgentes(ListadoDeduccionesPersonales)
    ConfigurardgDeduccionesPersonales
    Call CargardgDeduccionesPersonales(ListadoDeduccionesPersonales.dgAgentes.TextMatrix(1, 2))

End Sub

Private Sub MnuAgregarDeduccionesSIRADIG_Click()

    strCargaCodigoSIRADIG = "Deducciones"
    CargaCodigoSIRADIG.Show

End Sub

Private Sub MnuEliminarLiquidacionSISPER_Click()

    ListadoLiquidacionesSISPER.Show
    ConfigurardgLiquidacionSISPER
    CargardgLiquidacionSISPER

End Sub

Private Sub MnuEscalaAplicable_Click()

    ListadoEscalaGanancias.Show
    ConfigurardgNormasEscalaGanancias
    CargardgNormasEscalaGanancias
    ConfigurardgEscalaGanancias
    CargardgEscalaGanancias (ListadoEscalaGanancias.dgNormasEscalaGanancias.TextMatrix(1, 0))

End Sub



Private Sub MnuExportarSLAVE_Click()

    With ExportacionSLAVE
        .Show
        .txtAno.Text = Year(Now())
        .txtDecimal = "."
    End With

End Sub

Private Sub MnuF572_Click()

    ListadoF572.Show
    ConfigurardgPresentacionesF572
    ConfigurardgCargasDeFamiliaF572
    ConfigurardgDeduccionesGeneralesF572

End Sub

Private Sub MnuImportacionSISPER_Click()

    ImportacionPadronFamiliares.Show

End Sub

Private Sub MnuImportarComprobantesSIIF_Click()

    bolImportandoLiquidacionHonorariosSLAVE = True
    ImportacionLiquidacionHonorarios.Show

End Sub

Private Sub MnuImportarF572Web_Click()

    ImportacionF572Web.Show

End Sub

Private Sub MnuImportarHonorarios_Click()

    bolImportandoLiquidacionHonorariosSLAVE = False
    bolPrecarizadosSLAVE = False
    ImportacionLiquidacionHonorarios.Show

End Sub

Private Sub MnuImportarPrecarizados_Click()

    bolImportandoPrecarizadosSLAVE = True
    ImportacionLiquidacionHonorarios.Show

End Sub

Private Sub MnuImportarSueldoSISPER_Click()

    ImportacionLiquidacionSueldo.Show
    CargarCmbImportacionLiquidacionSueldo

End Sub

Private Sub MnuIncorporarConceptoSISPER_Click()

    IncorporarConceptoSueldo.Show
    CargarCmbIncorporarConceptoSueldo

End Sub

Private Sub MnuLimitesDeducciones_Click()

    ListadoLimitesDeducciones.Show
    ConfigurardgLimitesDeducciones
    CargardgLimitesDeducciones
    
End Sub

Private Sub MnuLiquidacionFinalGanancias_Click()

    LiquidacionFinalGanancias.Show

End Sub

Private Sub MnuLiquidacionMensualGanancias_Click()

    ListadoLiquidacionGanancias.Show
    ConfigurardgCodigosLiquidacionesGanancias
    CargardgCodigosLiquidacionesGanancias
    ConfigurardgAgentesRetenidos
    Call CargardgAgentesRetenidos(ListadoLiquidacionGanancias.dgCodigosLiquidacionesGanancias.TextMatrix(1, 0))

End Sub

Private Sub MnuLiquidacionPruebaSISPER_Click()

    LiquidacionPruebaSISPER.Show
    CargarCmbLiquidacionPruebaSisper

End Sub

Private Sub MnuListadoAgentes_Click()
    
    ListadoAgentes.Show
    Call ConfigurardgAgentes(ListadoAgentes)
    Call CargardgAgentes(ListadoAgentes)
    
End Sub

Private Sub MnuListadoCodigos_Click()

    ListadoCodigoLiquidaciones.Show
    ConfigurardgCodigoLiquidacion
    CargardgCodigoLiquidacion
    
End Sub

Private Sub MnuListadoComprobantes_Click()

    Dim x As Integer
    
    With ListadoComprobantesSIIF
        .Show
        ConfigurardgComprobantesSIIF
        Call CargardgComprobantesSIIF(, Year(Now()))
        .txtFecha.Text = Year(Now())
        x = .dgListadoComprobante.Row
        ConfigurardgImputacion
        CargardgImputacion (.dgListadoComprobante.TextMatrix(x, 0))
        ConfigurardgRetencion
        CargardgRetencion (.dgListadoComprobante.TextMatrix(x, 0))
        x = 0
    End With

End Sub

Private Sub MnuListadoConceptos_Click()

    ListadoConceptos.Show
    ConfigurardgConceptos
    CargardgConceptos

End Sub

Private Sub MnuListadoDeduccionesSIRADIG_Click()

    strListadoCodigoSIRADIG = "Deducciones"
    ListadoCodigosSIRADIG.Show
    Call ConfigurardgCodigosSIRADIG
    Call CargardgCodigosSIRADIG(strListadoCodigoSIRADIG)

End Sub

Private Sub MnuListadoFamiliares_Click()

    ListadoFamiliares.Show
    Call ConfigurardgAgentes(ListadoFamiliares)
    Call CargardgAgentes(ListadoFamiliares)
    ConfigurardgFamiliares
    Call CargardgFamiliares(ListadoFamiliares.dgAgentes.TextMatrix(1, 2))

End Sub

Private Sub MnuListadoOtrasDeduccionesSIRADIG_Click()

    strListadoCodigoSIRADIG = "OtrasDeducciones"
    ListadoCodigosSIRADIG.Show
    Call ConfigurardgCodigosSIRADIG
    Call CargardgCodigosSIRADIG(strListadoCodigoSIRADIG)

End Sub

Private Sub MnuListadoParentesco_Click()

    ListadoParentesco.Show
    ConfigurardgParentesco
    CargardgParentesco

End Sub

Private Sub MnuListadoParentescoSIRADIG_Click()
    
    strListadoCodigoSIRADIG = "Parentesco"
    ListadoCodigosSIRADIG.Show
    Call ConfigurardgCodigosSIRADIG
    Call CargardgCodigosSIRADIG(strListadoCodigoSIRADIG)

End Sub

Private Sub MnuListadoPorAgente_Click()

    ResumenAnualGanancias.Show
    ConfigurardgLiquidacionMensualGanancias (12)

End Sub

Private Sub MnuListadoPrecarizados_Click()

    ListadoPrecarizados.Show
    Call ConfigurardgPrecarizados
    Call CargardgPrecarizados

End Sub

Private Sub MnuAgregarOtrasDeduccionesSIRADIG_Click()

    strCargaCodigoSIRADIG = "OtrasDeducciones"
    CargaCodigoSIRADIG.Show
    
End Sub

Private Sub MnuMigrarDeducciones_Click()

    MigrarDeducciones.Show
    CargarCmbMigrarDeducciones

End Sub

Private Sub MnuReciboDeSueldo_Click()

    ReciboDeSueldo.Show
    ConfigurardgHaberesLiquidados
    ConfigurardgDescuentosLiquidados

End Sub
