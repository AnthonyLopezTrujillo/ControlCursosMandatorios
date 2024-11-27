/*********************************/
/************ TRIGGER ************/
/*********************************/

function procesoDiarioControlCursos(tipoProceso) {
  if (tipoProceso === TIPO_PROCESO_ONB) {
    ejecutarProcesoOnboarding();
  } else if (tipoProceso === TIPO_PROCESO_REG) {
    ejecutarProcesoRegular();
  }
}

/*********************************/
/*********************************/
/***** FUNCIONES A EJECUTAR ******/
/*********************************/
/*********************************/

function ejecutarProcesoOnboarding() {
  var dataControlColaboradores = obtenerControlColaboradores(TIPO_PROCESO_ONB);
  var dataAltasWorkday = obtenerAltasWorkday(TIPO_PROCESO_ONB);
  var mapaColaboradoresWorkday = convertirListaAColaboradoresMap(dataAltasWorkday);
  var dataItinerarios = obtenerDataItinerarios(TIPO_PROCESO_ONB);
  var dataGruposFormacion = obtenerGruposFormacion();
  var dataItinerariosDetalle = obtenerItinerariosConSusCursos(TIPO_PROCESO_ONB);
  var listaItinerariosDetalle = Object.keys(dataItinerariosDetalle);
  var grupoCursos = obtenerGrupoCursos(TIPO_PROCESO_ONB);
  var dataEncabezadosCursos = obtenerEncabezadosCursos(TIPO_PROCESO_ONB);

  
  console.log("dataControlColaboradores:: "+ dataControlColaboradores.length);
  console.log("dataAltasWorkday:: "+ dataAltasWorkday.length);
  console.log("mapaColaboradores:: "+ mapaColaboradoresWorkday.size);
  console.log("dataItinerarios:: "+ dataItinerarios.length);
  console.log("dataGruposFormacion:: "+ dataGruposFormacion.length);
  console.log("dataItinerariosDetalle:: "+ JSON.stringify(dataItinerariosDetalle));
  console.log("listaItinerariosDetalle:: " + listaItinerariosDetalle);
  console.log("grupoCursos:: "+ JSON.stringify(grupoCursos));
  console.log("dataEncabezadosCursos:: "+ dataEncabezadosCursos);

  //SE AÑADE ESTA MODIFICACION
  var dataAltasProcesoDos = obtenerAltasWorkday(TIPO_PROCESO_REG);
  var mapaColaboradoresWorkdayProcesoDos = convertirListaAColaboradoresMap(dataAltasProcesoDos);

  // Filtrar los colaboradores que no tienen el estado de eliminados
  var dataControlColaboradoresActivos = dataControlColaboradores.filter(function(item){ 
    return item[SHEET_CONTROL_COLAB_ESTADO_MATRICULA_IDX] !== ESTADO_INSCRIPCION_ELIMINADO;
  });
  console.log("dataControlColaboradoresActivos:: " + dataControlColaboradoresActivos.length);

  //HASTA AQUI
  
  /** PROCESO 1 */
  procesarDataNuevasAltas(dataControlColaboradores, dataAltasWorkday, dataItinerarios, dataGruposFormacion, dataItinerariosDetalle, dataEncabezadosCursos,TIPO_PROCESO_ONB);

  /** PROCESO 2 */
  // Verificar colaboradores no existentes y correos ---> SE MODIFICO LA DATA DE CONTROL A SOLO FILAS SIN ESTADO ELIMINADO- PARA QUE NO SE VUELVA A VALIDAR
  procesarColaboradoresSinEmail(dataControlColaboradoresActivos, mapaColaboradoresWorkdayProcesoDos,TIPO_PROCESO_ONB);

  /** PROCESO 3 */
  // Obtener colaboradores pendientes de control cursos
  var listaColaboradoresPendienteControl = obtenerDataControlColaboradoresPendientes(TIPO_PROCESO_ONB);
  // Verificar matricula de colaboradores a ITONS
  procesarVerificacionColaboradoresMatriculados(listaColaboradoresPendienteControl,listaItinerariosDetalle,TIPO_PROCESO_ONB);

  /** PROCESO 4 */
  // Lista de colaboradores con cursos pendientes
  var listaPendientesCursos = obtenerDataControlCursosPendientes(TIPO_PROCESO_ONB);
  // Verificar cumplimiento por curso
  procesarVerificacionCumplimientoCursos(listaPendientesCursos,dataEncabezadosCursos,grupoCursos,TIPO_PROCESO_ONB);

  /** PROCESO 5 */
  var listaColaboradoresPendienteControl = obtenerDataControlColaboradoresPendientes(TIPO_PROCESO_ONB);
  procesarVerificacionCumplimientoFinalCursos(listaColaboradoresPendienteControl,dataEncabezadosCursos,TIPO_PROCESO_ONB);
  
}

function inicializarProcesoRegular(codigoProcesoRegular){

  console.log("INICIALIZANDO PROCESO:: " + codigoProcesoRegular);

  var nuevasAltasControl = [];
  var nuevasAltasDetallado = [];

  var dataGruposFormacion = obtenerGruposFormacion();
  var dataItinerariosDetalle = obtenerItinerariosConSusCursos(TIPO_PROCESO_REG);
  var dataEncabezadosCursos = obtenerEncabezadosCursos(TIPO_PROCESO_REG,codigoProcesoRegular,dataItinerariosDetalle);
  var listaItinerarios = obtenerItinerariosPorCodigoProceso(codigoProcesoRegular);
  var fechaInicioProceso = listaItinerarios[0][SHEET_ITINERARIO_FECHA_INI_IDX];
  var fechaFinProceso = listaItinerarios[0][SHEET_ITINERARIO_FECHA_FIN_IDX];
  var dataAltasWorkday = obtenerAltasWorkday(TIPO_PROCESO_REG); 

  console.log("fechaInicioProceso: "+fechaInicioProceso);
  console.log("fechaFinProceso: "+fechaFinProceso);
  console.log("dataAltasWorkday:: "+dataAltasWorkday.length);

  var todasFechasIguales = true;

  var hojaColab = SS_CONTROL_CURSOS_REGULARES_BBVA.getSheetByName('Colab_' + codigoProcesoRegular);
  if (!hojaColab) {
      hojaColab = SS_CONTROL_CURSOS_REGULARES_BBVA.insertSheet('Colab_' + codigoProcesoRegular);
  }
  var hojaDetalle = SS_CONTROL_CURSOS_REGULARES_BBVA.getSheetByName('Detalle_' + codigoProcesoRegular);
  if (!hojaDetalle) {
      hojaDetalle = SS_CONTROL_CURSOS_REGULARES_BBVA.insertSheet('Detalle_' + codigoProcesoRegular);
  }

  hojaColab.getRange(1, 1, 1, LISTA_ENCABEZADO_NUEVA_HOJA_CONTROL.length).setValues([LISTA_ENCABEZADO_NUEVA_HOJA_CONTROL]);

  var encabezadosDetalle = LISTA_ENCABEZADO_NUEVA_HOJA_DETALLE.concat(dataEncabezadosCursos);
  hojaDetalle.getRange(1, 1, 1, encabezadosDetalle.length).setValues([encabezadosDetalle]);

  for (var i = 0; i < listaItinerarios.length; i++) {
    var itemItinerario = listaItinerarios[i];
    if (convertirDateAString(itemItinerario[SHEET_ITINERARIO_FECHA_INI_IDX]) != convertirDateAString(fechaInicioProceso)) {
      todasFechasIguales = false;
      break;
    }
  }

  if (todasFechasIguales){
    var fechaLimitePrograma = convertirStringADate(fechaFinProceso);

    console.log("entrando en fechas iguales::");
    var dataAltaAntesCorte = obtenerDataAltaAntesCorte(dataAltasWorkday, fechaInicioProceso);
    console.log("dataAltaAntesCorte: "+dataAltaAntesCorte.length);

    for(var i = 0; i < dataAltaAntesCorte.length; i++){
      var itemDataAltas = dataAltaAntesCorte[i];
      var itinerarioID = obtenerProgramaParaColaborador(listaItinerarios, itemDataAltas, TIPO_PROCESO_REG, codigoProcesoRegular,"");
      var codigoPrograma = itinerarioID ? itinerarioID[SHEET_ITINERARIO_CODIGO_IDX] : "NO ENCONTRADO";

      nuevasAltasControl.push(guardarDatosNuevaAltaEnLista(itemDataAltas, codigoPrograma, dataGruposFormacion, TIPO_PROCESO_REG,"",fechaLimitePrograma));
      nuevasAltasDetallado.push(guardarDatosNuevaAltaEnListaDetallado(itemDataAltas, codigoPrograma, dataItinerariosDetalle, dataEncabezadosCursos, TIPO_PROCESO_REG));
    }

    console.log("Nuevas Altas Control: ", nuevasAltasControl.length);
    console.log("Nuevas Altas Detallado: ", nuevasAltasDetallado.length);

    if (nuevasAltasControl.length > 0) {
      hojaColab.getRange(hojaColab.getLastRow() + 1, 1, nuevasAltasControl.length, nuevasAltasControl[0].length).setValues(nuevasAltasControl);
    }
    if (nuevasAltasDetallado.length > 0) {
      hojaDetalle.getRange(hojaDetalle.getLastRow() + 1, 1, nuevasAltasDetallado.length, nuevasAltasDetallado[0].length).setValues(nuevasAltasDetallado);
    }

  } else {
    console.log("Las fechas no son iguales");
  }
}

function ejecutarProcesoRegular() {
  var procesosRegularesActivos = obtenerProcesoRegularActivo();
  console.log("procesos a realizar:: " + procesosRegularesActivos);
  var dataAltasWorkday = obtenerAltasWorkday(TIPO_PROCESO_REG);
  var mapaColaboradoresWorkday = convertirListaAColaboradoresMap(dataAltasWorkday);
  console.log( "dataWorkday:: " + mapaColaboradoresWorkday.size);
  var dataItinerariosDetalle = obtenerItinerariosConSusCursos(TIPO_PROCESO_REG);
  var listaItinerariosDetalle = Object.keys(dataItinerariosDetalle);
  var grupoCursos = obtenerGrupoCursos(TIPO_PROCESO_REG);

  procesosRegularesActivos.forEach(function(procesoRegularActivo) {
    console.log("procesoRegularActivo:: " + procesoRegularActivo);

    var dataControlColaboradores = obtenerControlColaboradores(TIPO_PROCESO_REG, procesoRegularActivo);
    console.log("dataControlColaboradores:: " + dataControlColaboradores.length);

    var dataEncabezadosCursos = obtenerEncabezadosCursos(TIPO_PROCESO_REG, procesoRegularActivo, dataItinerariosDetalle);
    console.log("dataEncabezadosCursos:: " + dataEncabezadosCursos);

    /** PROCESO 2 */
    // Verificar colaboradores no existentes y correos
    procesarColaboradoresSinEmail(dataControlColaboradores, mapaColaboradoresWorkday, TIPO_PROCESO_REG, procesoRegularActivo);

    /** PROCESO 3 */
    // Obtener colaboradores pendientes de control cursos
    var listaColaboradoresPendienteControl = obtenerDataControlColaboradoresPendientes(TIPO_PROCESO_REG, procesoRegularActivo);
    // Verificar matricula de colaboradores a programa
    procesarVerificacionColaboradoresMatriculados(listaColaboradoresPendienteControl, listaItinerariosDetalle, TIPO_PROCESO_REG, procesoRegularActivo);

    /** PROCESO 4 */
    // Lista de colaboradores con cursos pendientes
    var listaPendientesCursos = obtenerDataControlCursosPendientes(TIPO_PROCESO_REG, procesoRegularActivo);
    // Verificar cumplimiento por curso
    procesarVerificacionCumplimientoCursos(listaPendientesCursos, dataEncabezadosCursos, grupoCursos, TIPO_PROCESO_REG, procesoRegularActivo);

    /** PROCESO 5 */
    listaColaboradoresPendienteControl = obtenerDataControlColaboradoresPendientes(TIPO_PROCESO_REG, procesoRegularActivo);
    procesarVerificacionCumplimientoFinalCursos(listaColaboradoresPendienteControl, dataEncabezadosCursos, TIPO_PROCESO_REG, procesoRegularActivo);
  });
}



/*********************************/
/*********************************/
/***** FUNCIONES PRINCIPALES *****/
/*********************************/
/*********************************/

/** PROCESO 1:: GUARDAR NUEVAS ALTAS EN HOJA CONTROL Y DETALLADO _ ESTA FUNCION SOLO SE USA EN ONBOARDING*/
function procesarDataNuevasAltas(dataControlColaboradores, dataAltasWorkday, dataParametria, dataGruposFormacion, dataItinerariosDetalle,dataEncabezadosCursos, tipoProceso){

  var nuevasAltasControl = [];
  var nuevasAltasDetallado = [];

  for (var i in dataAltasWorkday) {
    var itemNuevaAlta = dataAltasWorkday[i];
    var colaboradorID = itemNuevaAlta[SHEET_ALTAS_REGISTRO_IDX];
    var fechaIngreso = convertirDateAString(itemNuevaAlta[SHEET_ALTAS_FECHA_INGRESO_IDX]);
    
    var existe = validarSiAltaEstaEnBaseControl(dataControlColaboradores, colaboradorID, fechaIngreso, tipoProceso);

    // Si no existe en la base de control lo proceso
    if (!existe) {

      // Proceso solo si es de BBVA o de subsidiarias validas
      var empresa = itemNuevaAlta[SHEET_ALTAS_EMPRESA_IDX];
      if(empresa==EMPRESA_NOMBRE_BBVA || LISTA_SUSBSIDIARIAS_ONBOARDING.includes(empresa)){
      
        var fechaRegistroAltaWD = itemNuevaAlta[SHEET_ALTAS_FECHA_FIN_ALTA_IDX];
        var fechaPrimerDia = obtenerFechaPrimerDia(fechaIngreso, fechaRegistroAltaWD);
        var codigoProceso = ""; // PEND_DJB
        var programa = obtenerProgramaParaColaborador(dataParametria, itemNuevaAlta, tipoProceso, codigoProceso, fechaPrimerDia);
        var codigoPrograma = programa ? programa[SHEET_ITINERARIO_CODIGO_IDX] : "NO ENCONTRADO";
        
        nuevasAltasControl.push(guardarDatosNuevaAltaEnLista(itemNuevaAlta, codigoPrograma, dataGruposFormacion, tipoProceso, fechaPrimerDia));
        nuevasAltasDetallado.push(guardarDatosNuevaAltaEnListaDetallado(itemNuevaAlta, codigoPrograma, dataItinerariosDetalle, dataEncabezadosCursos, tipoProceso));

      }
    }
  }

    // Añadir todas las filas a las hojas de una sola vez
  if (nuevasAltasControl.length > 0) {
    if (tipoProceso == TIPO_PROCESO_ONB) {
      SHEET_CONTROL_COLABORADORES.getRange(SHEET_CONTROL_COLABORADORES.getLastRow() + 1, 1, nuevasAltasControl.length, nuevasAltasControl[0].length).setValues(nuevasAltasControl);
    } else if (tipoProceso == TIPO_PROCESO_REG) {
      SHEET_CONTROL_COLABORA_REGULARES_BBVA.getRange(SHEET_CONTROL_COLABORA_REGULARES_BBVA.getLastRow() + 1, 1, nuevasAltasControl.length, nuevasAltasControl[0].length).setValues(nuevasAltasControl);
    }
  }

  if (nuevasAltasDetallado.length > 0) {
    if (tipoProceso == TIPO_PROCESO_ONB) {
      SHEET_CONTROL_DETALLADO_BBVA.getRange(SHEET_CONTROL_DETALLADO_BBVA.getLastRow() + 1, 1, nuevasAltasDetallado.length, nuevasAltasDetallado[0].length).setValues(nuevasAltasDetallado);
    } else if (tipoProceso == TIPO_PROCESO_REG) {
      SHEET_CONTROL_DETALLADO_REGULARES_BBVA.getRange(SHEET_CONTROL_DETALLADO_REGULARES_BBVA.getLastRow() + 1, 1, nuevasAltasDetallado.length, nuevasAltasDetallado[0].length).setValues(nuevasAltasDetallado);
    }
  }

  SpreadsheetApp.flush();
}

/** PROCESO 2:: CAMBIAR ESTADO A ELIMINADO LSO COLABORADORES QUE YA NO SE ENCUENTREN (HOJA CONTROL, NO HOJA DETALLE)*/
function procesarColaboradoresSinEmail(dataControlColaboradores, dataAltasWorkdayMap, tipoProceso,procesoRegularActivo) {
    for (var i = 0; i < dataControlColaboradores.length; i++) {
        var itemAltaActual = dataControlColaboradores[i];
        var altaActualColaboradorID = itemAltaActual[SHEET_CONTROL_COLAB_COLABORADOR_ID_IDX];
        var filaColaboradorActual = itemAltaActual[SHEET_CONTROL_COLAB_FILA_IDX];

        var itemColaboradorNuevo = dataAltasWorkdayMap.get(altaActualColaboradorID);

        if (!itemColaboradorNuevo) {
            // Caso: Colaborador ya no está en la lista, marcar como eliminado
            //guardarControlColaboradorEstadoMatricula(filaColaboradorActual, ESTADO_INSCRIPCION_ELIMINADO, new Date(), tipoProceso,procesoRegularActivo);
            guardarEstadoEliminadoControlColaborador(filaColaboradorActual, ESTADO_INSCRIPCION_ELIMINADO, new Date(), tipoProceso,procesoRegularActivo);
        } else if (!itemAltaActual[SHEET_CONTROL_COLAB_COLABORADOR_EMAIL_IDX]) {
            // Caso: Actualizar email si está vacío
            var emailColaboradorNuevo = itemColaboradorNuevo[SHEET_ALTAS_EMAIL_IDX];
            guardarControlEmailColaborador(filaColaboradorActual, emailColaboradorNuevo, tipoProceso,procesoRegularActivo);
        }
    }
}

/** PROCESO 3:: SE VERIFICA EL ESTADO SI HA CULMINADO EL ITINERARIO O NO */
function procesarVerificacionColaboradoresMatriculados(listaColaboradoresPendienteControl,listaItinerariosDetalle,tipoProceso,procesoRegularActivo){

  // Procesar matriculados por itinerario
  for(var i in listaItinerariosDetalle){
    var itinerarioID = listaItinerariosDetalle[i];
    // Obtener pendientes matricula de itinerario
    var listaPendientesMatricula = obtenerPendientesMatriculaPorItinerario(listaColaboradoresPendienteControl, itinerarioID);
    console.log("listaPendientesMatricula "+itinerarioID+":: "+listaPendientesMatricula.length);
    if(listaPendientesMatricula!=null && listaPendientesMatricula.length>0){
      // Obtener archivo del itinerario
      var itinerarioFileID = obtenerFilePorInicioNombre(DRIVE_FOLDER_ITINERARIO_COLABORADORES_ID, itinerarioID);
      console.log("itinerarioID="+itinerarioID+" - fileID="+itinerarioFileID);
      if(itinerarioFileID!=null){
        // Obtener data de itinerario de colaboradores de cornerstone
        var listaCornerstoneItinerarioColaborador = obtenerDataCornerstoneItinerario(itinerarioID, itinerarioFileID,tipoProceso);
        // Verificar estado de matriculas del itinerario
        verificarEstadoMatriculaDeItinerario(listaPendientesMatricula, listaCornerstoneItinerarioColaborador,tipoProceso,procesoRegularActivo);
      }
    }
  }
}

/** PROCESO 4:: SE VERIFICA EL ESTADO POR CURSO */
function procesarVerificacionCumplimientoCursos(listaPendientesCursos, encabezadosControlDetalle, grupoCursos, tipoProceso, procesoRegularActivo) {
  // Obtener el detalle de itinerarios basado en el tipo de proceso
  var dataItinerariosDetalle = tipoProceso === TIPO_PROCESO_ONB 
    ? obtenerItinerariosConSusCursos(TIPO_PROCESO_ONB) 
    : obtenerItinerariosConSusCursos(TIPO_PROCESO_REG);
  
  console.log("dataItinerariosDetalle:: "+ Object.keys(dataItinerariosDetalle))

  // Definir la hoja en base al tipo de proceso
  var hoja = tipoProceso === TIPO_PROCESO_ONB 
    ? SHEET_CONTROL_DETALLADO_BBVA 
    : SS_CONTROL_CURSOS_REGULARES_BBVA.getSheetByName(PREFIJO_HOJA_REGULAR_DETALLADO + procesoRegularActivo);

  for (var grupo in grupoCursos) {
    var cursosGrupo = grupoCursos[grupo]; // Lista de cursos por grupo
    var grupoFileID = obtenerFilePorGrupoID(grupo);
    console.log("Procesando grupo: " + grupo + " con cursos: " + cursosGrupo + " - fileID: " + grupoFileID);
    
    if (grupoFileID != null) {
      var listaCornerstoneCursos = obtenerDataCornerstoneGrupo(grupoFileID);
      console.log("Datos obtenidos para el grupo " + grupo + ": " + listaCornerstoneCursos.length + " registros.");

      // Procesar cada curso dentro del grupo
      cursosGrupo.forEach(function(cursoID) {   
        var columnaCurso = obtenerColumnaControlDelCurso(cursoID, encabezadosControlDetalle);
        if (columnaCurso != -1) {
          listaPendientesCursos.forEach(function(itemColaborador) {
            var estadoCurso = itemColaborador[columnaCurso];
            var filaRegistro = itemColaborador[SHEET_CONTROL_DET_FILA_IDX];
            var colaboradorID = itemColaborador[SHEET_CONTROL_DET_COLABORADOR_ID_IDX];
            
            // Obtener el itinerario asignado y validar que esté en dataItinerariosDetalle
            var itonAsignado = itemColaborador[SHEET_CONTROL_DET_ITINERARIO_ID_IDX];
            var cursosAsignadosITON = dataItinerariosDetalle[itonAsignado];

            if (!cursosAsignadosITON) {
              //console.log("No se encontró el ITON: " + itonAsignado + " en dataItinerariosDetalle. No se realiza la actualización.");
              return; 
            }

            // Validar si el cursoID está asignado al ITON del colaborador
            if (cursosAsignadosITON.includes(cursoID)) {
              if (estadoCurso !== ESTADO_CURSO_TERMINADO && estadoCurso !== ESTADO_CURSO_NO_APLICA && estadoCurso !== "") {
                var itemEstadoCursoCornerstone = obtenerEstadoCurso(itemColaborador, cursoID, listaCornerstoneCursos);
                if (itemEstadoCursoCornerstone != null) {
                  var nuevoEstado = itemEstadoCursoCornerstone[SHEET_CORNERS_ITINERARIO_ESTADO_IDX];
                  if (nuevoEstado != estadoCurso) {
                    updateFila(filaRegistro, columnaCurso, nuevoEstado, hoja);
                    //console.log("Actualizado curso: " + cursoID + " para colaborador: " + colaboradorID);
                  }
                }
              }
            }
          });
        }
      });
    }
  }
}

/** PROCESO 5:: COMPROBAR ESTADO HOJAS */
function procesarVerificacionCumplimientoFinalCursos(listaColaboradoresPendienteControl,encabezadosControlDetallado,tipoProceso,procesoRegularActivo){

  console.log("encabezadosCursos:: " + encabezadosControlDetallado);

  // Lista de colaboradores con cursos pendientes
  var listaColaboradoresPendientesCursos = obtenerDataControlCursosPendientes(tipoProceso,procesoRegularActivo);
  console.log("listaColaboradoresPendientesCursos: " + listaColaboradoresPendientesCursos.length);
  
  // Verificar cursos por colaborador
  for (var i in listaColaboradoresPendientesCursos) {
    var indFinalizaCursos = true; 
    var itemPendiente = listaColaboradoresPendientesCursos[i];

    // Verificacion de cumplimiento de cursos x colaborador
    for (var j = 0; j < encabezadosControlDetallado.length; j++) {
      var estadoCurso = itemPendiente[j+SHEET_ITINERARIO_DET_CURSO_DELIMITADOR_ID_IDX];
      if (estadoCurso === ESTADO_CURSO_PENDIENTE || estadoCurso === ESTADO_CURSO_EN_PROCESO || estadoCurso === ESTADO_CURSO_INSCRIT0 ||  estadoCurso === ESTADO_CURSO_EMPEZADO || estadoCurso === "") {
        indFinalizaCursos = false;
        break;
      }
    }

    // Si el colaborador ha finalizado todos los cursos
    if (indFinalizaCursos) {
      // Obtener datos del caso
      var filaControlDetallado = itemPendiente[SHEET_CONTROL_DET_FILA_IDX];
      var colaboradorID = itemPendiente[SHEET_CONTROL_DET_COLABORADOR_ID_IDX];
      var itinerarioID = itemPendiente[SHEET_CONTROL_DET_ITINERARIO_ID_IDX];

      // Actualizar estado en control detallado
      guardarControlDetalladoEstadoCursos(filaControlDetallado, ESTADO_CURSO_FINAL_TERMINADO,tipoProceso,procesoRegularActivo);

      // Actualizar estado en control general
      var itemControlColaborador = obtenerControlColaboradorPendiente(listaColaboradoresPendienteControl, colaboradorID, itinerarioID);
      if (itemControlColaborador != null) {
        var filaControlColaborador = itemControlColaborador[SHEET_CONTROL_COLAB_FILA_IDX];
        guardarControlColaboradorEstadoCursos(filaControlColaborador, ESTADO_CURSO_FINAL_TERMINADO,tipoProceso,procesoRegularActivo);
      }
    }
  }
}

/************************************/
/********* FUNCIONES CORREOS ********/
/************************************/
/*
function procesarRecordatoriosControlCursos(){
  //var asuntoCorreo="", imagenBanner=null, columnaAviso=0;
  var listaColabRecordatorio3 = new Array();
  var listaColabVencidos = new Array(); // Lista para RRLL

  /* CORREOS A INDIVIDUALES A COLABORADORES */
  // Obtener banners
  /*
  var bannerRecordatorio1 = obtenerImagenDrive(FOLDER_IMAGENES_CONTROL_CURSOS, IMG_NOMBRE_RECORDATORIO_1);
  var bannerRecordatorio2 = obtenerImagenDrive(FOLDER_IMAGENES_CONTROL_CURSOS, IMG_NOMBRE_RECORDATORIO_2);
  var bannerRecordatorio3 = obtenerImagenDrive(FOLDER_IMAGENES_CONTROL_CURSOS, IMG_NOMBRE_RECORDATORIO_3);

  // Obtener lista cursos pendientes de recordatorio
  var listaPendientesRecordatorio = obtenerDataControlPendientesRecordatorio(TIPO_PROCESO_ONB);

  if (listaPendientesRecordatorio != null) {
    console.log("Lista Pendientes Recordatorio Onboarding:: "+listaPendientesRecordatorio.length);
    listaColabRecordatorio3 = procesarColaboradoresRecordatorios(listaPendientesRecordatorio, bannerRecordatorio1, bannerRecordatorio2, bannerRecordatorio3);
  }

  /* CORREOS A ADVISORS */
  // Enviar correo de 20 dias a Advisors
  /*
  if(listaColabRecordatorio3.length>0){
    enviarCorreoEscalaAdvisors(listaColabRecordatorio3,TIPO_PROCESO_ONB);
    // Registrar envio de correos a advisors
    for(var i in listaColabRecordatorio3){
      var itemRecordatorio = listaColabRecordatorio3[i];
      var filaControlColaborador = itemRecordatorio[SHEET_CONTROL_COLAB_FILA_IDX];
      guardarControlEnvioRecordatorio(filaControlColaborador, SHEET_CONTROL_COLAB_ESCALA_ADVISOR_IDX);
    }
  }

}*/

function procesarRecordatoriosControlCursos() {
  //var asuntoCorreo="", imagenBanner=null, columnaAviso=0;
  var listaColabRecordatorio3 = new Array();
  var listaColabVencidos = new Array(); // Lista para RRLL

  /* CORREOS A INDIVIDUALES A COLABORADORES */
  // Obtener banners
  var bannerRecordatorio1 = obtenerImagenDrive(FOLDER_IMAGENES_CONTROL_CURSOS, IMG_NOMBRE_RECORDATORIO_1);
  var bannerRecordatorio2 = obtenerImagenDrive(FOLDER_IMAGENES_CONTROL_CURSOS, IMG_NOMBRE_RECORDATORIO_2);
  var bannerRecordatorio3 = obtenerImagenDrive(FOLDER_IMAGENES_CONTROL_CURSOS, IMG_NOMBRE_RECORDATORIO_3);

  // Obtener lista cursos pendientes de recordatorio
  var listaPendientesRecordatorio = obtenerDataControlPendientesRecordatorio(TIPO_PROCESO_ONB);
  if (listaPendientesRecordatorio !== null && listaPendientesRecordatorio.length > 0) {
    console.log("Lista Pendientes Recordatorio Onboarding:: " + listaPendientesRecordatorio.length);
    listaColabRecordatorio3 = procesarColaboradoresRecordatorios(listaPendientesRecordatorio, bannerRecordatorio1, bannerRecordatorio2, bannerRecordatorio3);
  }

  /* CORREOS A ADVISORS */
  // Enviar correo de 20 dias a Advisors
  if (listaColabRecordatorio3 != null || listaColabRecordatorio3.length > 0) {
    console.log("Lista Advisors:: " + listaColabRecordatorio3.length);
    enviarCorreoEscalaAdvisors(listaColabRecordatorio3, TIPO_PROCESO_ONB);
    // Registrar envio de correos a advisors
    for (var i = 0; i < listaColabRecordatorio3.length; i++) {
      var itemRecordatorio = listaColabRecordatorio3[i];
      var filaControlColaborador = itemRecordatorio[SHEET_CONTROL_COLAB_FILA_IDX];
      guardarEnvioCorreoRRLL(filaControlColaborador, SHEET_CONTROL_COLAB_ESCALA_ADVISOR_IDX);
    }
  }

  // Obtener lista para RRLL
  var listaPendientesRRLL = obtenerDataOnboardingRRLL();
  if (listaPendientesRRLL != null || listaPendientesRRLL.length > 0) {
    console.log("Lista Pendientes RRLL:: " + listaPendientesRRLL.length);
    listaColabVencidos = procesarColaboradoresRRLL(listaPendientesRRLL);
  }

  /* CORREOS A RRLL */
  // Enviar correo de mas de 31 dias a RRLL
  if (listaColabVencidos !== null && listaColabVencidos.length > 0) {
    enviarCorreoEscalaRRLL(listaColabVencidos, TIPO_PROCESO_ONB);
    // Registrar envio de correos a advisors
    for (var i = 0; i < listaColabVencidos.length; i++) {
      var itemRecordatorio = listaColabVencidos[i];
      var filaControlColaborador = itemRecordatorio[SHEET_CONTROL_COLAB_FILA_IDX];
      guardarEnvioCorreoRRLL(filaControlColaborador, SHEET_CONTROL_COLAB_ESCALA_RRLL_IDX);
    }
  }



}

function procesarRecordatoriosControlRegulares() {
  var asuntoCorreo = "", imagenBanner = null, columnaAviso = 0;
  var listaColabRecordatorioAdvisors = new Array();
  var listaColabRecordatorioRRLL = new Array();

  // Obtener banners
  var bannerRecordatorio1 = obtenerImagenDrive(FOLDER_IMAGENES_CONTROL_CURSOS, IMG_NOMBRE_REGULAR_RECORDATORIO_1);
  var bannerRecordatorio2 = obtenerImagenDrive(FOLDER_IMAGENES_CONTROL_CURSOS, IMG_NOMBRE_REGULAR_RECORDATORIO_2);
  var bannerRecordatorio3 = obtenerImagenDrive(FOLDER_IMAGENES_CONTROL_CURSOS, IMG_NOMBRE_REGULAR_RECORDATORIO_3);

  //OBTENER PROCESOS ACTIVOS EJM:: *REG-2024-1* *REG-2024-2*
  var procesosRegularesActivos = obtenerProcesoRegularActivo();
  console.log("Procesos a ejecutar:: " + procesosRegularesActivos);
  procesosRegularesActivos.forEach(function (procesoRegularActivo) {
    console.log("Ejecutando:: " + procesoRegularActivo);

    var listaPendientesRecordatorio = obtenerDataControlPendientesRecordatorio(TIPO_PROCESO_REG, procesoRegularActivo);
    console.log("listaPendientesRecordatorio=" + listaPendientesRecordatorio.length);
    

    var listaPendientesRecordatorioRRLL = obtenerDataControlPendientesRecordatorioRRLL(procesoRegularActivo);
    console.log("listaPendientesRecordatorioRRLL=" + listaPendientesRecordatorioRRLL.length);


    // Enviar recordatorio por colaborador
    for (var i = 0; i < listaPendientesRecordatorio.length; i++) {
      var itemRecordatorio = listaPendientesRecordatorio[i];
      var diasTranscurridos = itemRecordatorio[SHEET_CONTROL_COLAB_DIAS_TRANSCURRIDOS_IDX];
      var avisoRecordatorio1 = itemRecordatorio[SHEET_CONTROL_COLAB_RECORDATORIO_1_IDX];
      var avisoRecordatorio2 = itemRecordatorio[SHEET_CONTROL_COLAB_RECORDATORIO_2_IDX];
      var avisoRecordatorio3 = itemRecordatorio[SHEET_CONTROL_COLAB_RECORDATORIO_3_IDX];
      var areaColaborador = itemRecordatorio[SHEET_CONTROL_COLAB_AREA_IDX];
      var avisoRecordatorioAdvisor = itemRecordatorio[SHEET_CONTROL_COLAB_ESCALA_ADVISOR_IDX];
      var avisoRecordatorioAdvisor2 = itemRecordatorio[SHEET_CONTROL_COLAB_AVISO_ADVISOR_2_REGULAR_IDX];
      var envioRecordatorio = false;

      // Obtener por tipo de recordatorio
      if (DIAS_RECORDATORIO_REGULAR_COLAB_1 == diasTranscurridos && avisoRecordatorio1 == "") {
        envioRecordatorio = true;
        asuntoCorreo = "¡Vamos, tú puedes! Termina tu itinerario";
        imagenBanner = bannerRecordatorio1;
        columnaAviso = SHEET_CONTROL_COLAB_RECORDATORIO_1_IDX;

      } else if (DIAS_RECORDATORIO_REGULAR_COLAB_2 == diasTranscurridos && (avisoRecordatorio2 == "" || avisoRecordatorioAdvisor == "")) {
        envioRecordatorio = true;
        asuntoCorreo = "¡Se acaba el tiempo! Finaliza tus cursos mandatorios";
        imagenBanner = bannerRecordatorio2;
        columnaAviso = SHEET_CONTROL_COLAB_RECORDATORIO_2_IDX;
        if (AREA_BANCA_COMERCIAL === areaColaborador || AREA_BANCA_EMPRESA === areaColaborador) {
          listaColabRecordatorioAdvisors.push(itemRecordatorio);
        }
      } else if (DIAS_NOTIFICACION_REGULAR_ADVISOR_2 == diasTranscurridos && avisoRecordatorioAdvisor2 == "") {
        if (AREA_BANCA_COMERCIAL === areaColaborador || AREA_BANCA_EMPRESA === areaColaborador) {
          listaColabRecordatorioAdvisors.push(itemRecordatorio);
        }
      } else if ((DIAS_RECORDATORIO_REGULAR_COLAB_3 == diasTranscurridos || DIAS_RECORDATORIO_REGULAR_COLAB_3 - 1 == diasTranscurridos || DIAS_RECORDATORIO_REGULAR_COLAB_3 -2 == diasTranscurridos) && avisoRecordatorio3 == "") {
        envioRecordatorio = true;
        asuntoCorreo = "El 86% ya completó sus mandatorios";
        imagenBanner = bannerRecordatorio3;
        columnaAviso = SHEET_CONTROL_COLAB_RECORDATORIO_3_IDX;
      }

      // Enviar recordatorios individuales
      if (envioRecordatorio === true) {
        // Enviar correo recordatorio a colaborador
        enviarCorreoRecordatorio(itemRecordatorio, diasTranscurridos, asuntoCorreo, imagenBanner); // TEST_DJB
        // Registrar envio de correo en base
        var filaControlColaborador = itemRecordatorio[SHEET_CONTROL_COLAB_FILA_IDX];
        guardarControlEnvioRecordatorio(filaControlColaborador, columnaAviso, procesoRegularActivo);
      }
    }

    /* CORREOS A ADVISORS */
    // Enviar correo a Advisors
    if (listaColabRecordatorioAdvisors != null && listaColabRecordatorioAdvisors.length > 0) {
      enviarCorreoEscalaAdvisors(listaColabRecordatorioAdvisors, TIPO_PROCESO_REG);
      // Registrar envio de correos a advisors
      for (var i = 0; i < listaColabRecordatorioAdvisors.length; i++) {
        var itemRecordatorio = listaColabRecordatorioAdvisors[i];
        var diasTranscurridosItem = itemRecordatorio[SHEET_CONTROL_COLAB_DIAS_TRANSCURRIDOS_IDX];
        var filaControlColaborador = itemRecordatorio[SHEET_CONTROL_COLAB_FILA_IDX];
        if (diasTranscurridosItem == DIAS_RECORDATORIO_REGULAR_COLAB_2) {
          var columnaFecha = SHEET_CONTROL_COLAB_ESCALA_ADVISOR_IDX;
        } else if (diasTranscurridosItem == DIAS_NOTIFICACION_REGULAR_ADVISOR_2) {
          var columnaFecha = SHEET_CONTROL_COLAB_AVISO_ADVISOR_2_REGULAR_IDX;
        }
        guardarControlEnvioRecordatorio(filaControlColaborador, columnaFecha, procesoRegularActivo);
      }
    }

    if (listaPendientesRecordatorioRRLL != null && listaPendientesRecordatorioRRLL.length > 0) {
      for (var i = 0; i < listaPendientesRecordatorioRRLL.length; i++) {
        var itemRecordatorio = listaPendientesRecordatorioRRLL[i];
        var diasTranscurridos = itemRecordatorio[SHEET_CONTROL_COLAB_DIAS_TRANSCURRIDOS_IDX];
        var areaColaborador = itemRecordatorio[SHEET_CONTROL_COLAB_AREA_IDX];
        var avisoRecordatorioRRLL = itemRecordatorio[SHEET_CONTROL_COLAB_ESCALA_RRLL_IDX];
        if (diasTranscurridos >= DIAS_NOTIFICACION_REGULAR_RRLL && avisoRecordatorioRRLL == "") {
          listaColabRecordatorioRRLL.push(itemRecordatorio);
        }
      }

    }
    console.log("listaColabRecordatorioRRLL:: " + listaColabRecordatorioRRLL.length)

    if (listaColabRecordatorioRRLL != null && listaColabRecordatorioRRLL.length > 0) {
      enviarCorreoEscalaRRLL(listaColabRecordatorioRRLL);
      for (var i = 0; i < listaColabRecordatorioRRLL.length; i++) {
        var itemRecordatorio = listaColabRecordatorioRRLL[i];
        var filaControlColaborador = itemRecordatorio[SHEET_CONTROL_COLAB_FILA_IDX];
        guardarControlEnvioRecordatorio(filaControlColaborador, SHEET_CONTROL_COLAB_ESCALA_RRLL_IDX, procesoRegularActivo);
      }
    }

  });


}

/*
function procesarRecordatoriosControlRegulares(){
  var asuntoCorreo="", imagenBanner=null, columnaAviso=0;
  var listaColabRecordatorioAdvisors = new Array();
  var listaColabRecordatorioRRLL = new Array();

  // Obtener banners
  var bannerRecordatorio1 = obtenerImagenDrive(FOLDER_IMAGENES_CONTROL_CURSOS, IMG_NOMBRE_REGULAR_RECORDATORIO_1);
  var bannerRecordatorio2 = obtenerImagenDrive(FOLDER_IMAGENES_CONTROL_CURSOS, IMG_NOMBRE_REGULAR_RECORDATORIO_2);
  var bannerRecordatorio3 = obtenerImagenDrive(FOLDER_IMAGENES_CONTROL_CURSOS, IMG_NOMBRE_REGULAR_RECORDATORIO_3);

  //OBTENER PROCESOS ACTIVOS EJM:: *REG-2024-1* *REG-2024-2*
  var procesosRegularesActivos = obtenerProcesoRegularActivo();
  console.log("Procesos a ejecutar:: " + procesosRegularesActivos);
  procesosRegularesActivos.forEach(function(procesoRegularActivo){
    console.log("Ejecutando:: " + procesoRegularActivo);

    var listaPendientesRecordatorio = obtenerDataControlPendientesRecordatorio(TIPO_PROCESO_REG,procesoRegularActivo);
    if(listaPendientesRecordatorio!=null){
      console.log("listaPendientesRecordatorio="+listaPendientesRecordatorio.length);
    }

    var listaPendientesRecordatorioRRLL = obtenerDataControlPendientesRecordatorioRRLL(procesoRegularActivo);
    if(listaPendientesRecordatorioRRLL!=null){
      console.log("listaPendientesRecordatorioRRLL="+listaPendientesRecordatorioRRLL.length);
    }

    // Enviar recordatorio por colaborador
    for(var i in listaPendientesRecordatorio){
      var itemRecordatorio = listaPendientesRecordatorio[i];
      var diasTranscurridos = itemRecordatorio[SHEET_CONTROL_COLAB_DIAS_TRANSCURRIDOS_IDX];
      var avisoRecordatorio1 = itemRecordatorio[SHEET_CONTROL_COLAB_RECORDATORIO_1_IDX];
      var avisoRecordatorio2 = itemRecordatorio[SHEET_CONTROL_COLAB_RECORDATORIO_2_IDX];
      var avisoRecordatorio3 = itemRecordatorio[SHEET_CONTROL_COLAB_RECORDATORIO_3_IDX];
      var areaColaborador = itemRecordatorio[SHEET_CONTROL_COLAB_AREA_IDX];
      var avisoRecordatorioAdvisor = itemRecordatorio[SHEET_CONTROL_COLAB_ESCALA_ADVISOR_IDX];
      var avisoRecordatorioAdvisor2 = itemRecordatorio[SHEET_CONTROL_COLAB_AVISO_ADVISOR_2_REGULAR_IDX];
      var envioRecordatorio = false;

      // Obtener por tipo de recordatorio
      if(DIAS_RECORDATORIO_REGULAR_COLAB_1==diasTranscurridos && avisoRecordatorio1==""){
        envioRecordatorio = true;
        asuntoCorreo = "¡Vamos, tú puedes! Termina tu itinerario";
        imagenBanner = bannerRecordatorio1;
        columnaAviso= SHEET_CONTROL_COLAB_RECORDATORIO_1_IDX;

      }else if(DIAS_RECORDATORIO_REGULAR_COLAB_2==diasTranscurridos && avisoRecordatorio2=="" && avisoRecordatorioAdvisor == ""){
        envioRecordatorio = true;
        asuntoCorreo = "¡Se acaba el tiempo! Finaliza tus cursos mandatorios";
        imagenBanner = bannerRecordatorio2;
        columnaAviso= SHEET_CONTROL_COLAB_RECORDATORIO_2_IDX;
        if(AREA_BANCA_COMERCIAL===areaColaborador || AREA_BANCA_EMPRESA === areaColaborador){
          listaColabRecordatorioAdvisors.push(itemRecordatorio);
        }
      }else if(DIAS_NOTIFICACION_REGULAR_ADVISOR_2 == diasTranscurridos && avisoRecordatorioAdvisor2 == ""){
        if(AREA_BANCA_COMERCIAL===areaColaborador || AREA_BANCA_EMPRESA === areaColaborador){
          listaColabRecordatorioAdvisors.push(itemRecordatorio);
        }
      }else if(DIAS_RECORDATORIO_REGULAR_COLAB_3==diasTranscurridos && avisoRecordatorio3==""){
        envioRecordatorio = true;
        asuntoCorreo = "El 86% ya completó sus mandatorios";
        imagenBanner = bannerRecordatorio3;
        columnaAviso= SHEET_CONTROL_COLAB_RECORDATORIO_3_IDX;
      }

      // Enviar recordatorios individuales
      if(envioRecordatorio === true){
        // Enviar correo recordatorio a colaborador
        enviarCorreoRecordatorio(itemRecordatorio, diasTranscurridos, asuntoCorreo, imagenBanner); // TEST_DJB
        // Registrar envio de correo en base
        var filaControlColaborador = itemRecordatorio[SHEET_CONTROL_COLAB_FILA_IDX];
        guardarControlEnvioRecordatorio(filaControlColaborador, columnaAviso,procesoRegularActivo);
      }
    }

    // Enviar correo a Advisors
    if(listaColabRecordatorioAdvisors.length>0){
      enviarCorreoEscalaAdvisors(listaColabRecordatorioAdvisors,TIPO_PROCESO_REG);
      // Registrar envio de correos a advisors
      for(var i in listaColabRecordatorioAdvisors){
        var itemRecordatorio = listaColabRecordatorioAdvisors[i];
        var diasTranscurridosItem = itemRecordatorio[SHEET_CONTROL_COLAB_DIAS_TRANSCURRIDOS_IDX];
        var filaControlColaborador = itemRecordatorio[SHEET_CONTROL_COLAB_FILA_IDX];
        if(diasTranscurridosItem == DIAS_RECORDATORIO_REGULAR_COLAB_2){
          var columnaFecha = SHEET_CONTROL_COLAB_ESCALA_ADVISOR_IDX;
        }else if(diasTranscurridosItem == DIAS_NOTIFICACION_REGULAR_ADVISOR_2){
          var columnaFecha = SHEET_CONTROL_COLAB_AVISO_ADVISOR_2_REGULAR_IDX;
        }
        guardarControlEnvioRecordatorio(filaControlColaborador, columnaFecha,procesoRegularActivo);
      }
    }

    for(var i in listaPendientesRecordatorioRRLL){
      var itemRecordatorio = listaPendientesRecordatorioRRLL[i];
      var diasTranscurridos = itemRecordatorio[SHEET_CONTROL_COLAB_DIAS_TRANSCURRIDOS_IDX];
      var areaColaborador = itemRecordatorio[SHEET_CONTROL_COLAB_AREA_IDX];
      var avisoRecordatorioRRLL = itemRecordatorio[SHEET_CONTROL_COLAB_ESCALA_RRLL_IDX];
      if(diasTranscurridos >= DIAS_NOTIFICACION_REGULAR_RRLL && avisoRecordatorioRRLL==""){
        listaColabRecordatorioRRLL.push(itemRecordatorio);
      }
    }

    console.log( "listaColabRecordatorioRRLL:: "+  listaColabRecordatorioRRLL.length)

    if(listaColabRecordatorioRRLL.length > 0){
      enviarCorreoEscalaRRLL(listaColabRecordatorioRRLL);
      for(var i in listaColabRecordatorioRRLL){
        var itemRecordatorio = listaColabRecordatorioRRLL[i];
        var filaControlColaborador = itemRecordatorio[SHEET_CONTROL_COLAB_FILA_IDX];
          guardarControlEnvioRecordatorio(filaControlColaborador, SHEET_CONTROL_COLAB_ESCALA_RRLL_IDX,procesoRegularActivo);
      }
    }
  });

}
*/
