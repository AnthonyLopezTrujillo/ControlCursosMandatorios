// Data Colaboradores Workday (Activos y futuros)
const SS_ALTAS_BBVA_ID = "";
const SS_ALTAS_BBVA = SpreadsheetApp.openById(SS_ALTAS_BBVA_ID);
const SHEET_COLABORADORES_BBVA_NOMBRE = "ColaboradoresOnb";
/** TEMPORAL */
const SHEET_COLABORADORES_BBVA_ACTUALES_NOMBRE = "ColaboradoresActual";
const SHEET_COLABORADORES_BBVA = SS_ALTAS_BBVA.getSheetByName(SHEET_COLABORADORES_BBVA_NOMBRE);

// Data Parametria Cursos
const SS_PARAMETRIA_CURSOS_ID = "";
const SS_PARAMETRIA_CURSOS_BBVA = SpreadsheetApp.openById(SS_PARAMETRIA_CURSOS_ID);
const SHEET_ITINERARIO_BBVA_NOMBRE = "Itinerario";
const SHEET_ITINERARIO_DETALLE_BBVA_NOMBRE = "ItinerarioDetalleOnboarding";
const SHEET_ITINERARIO_DETALLE_REGULAR_BBVA_NOMBRE = "ItinerarioDetalleRegular";

const SHEET_PARAMETRIA_ITINERARIO_DETALLE = SS_PARAMETRIA_CURSOS_BBVA.getSheetByName(SHEET_ITINERARIO_DETALLE_BBVA_NOMBRE);
const SHEET_PARAMETRIA_ITINERARIO_DETALLE_REGULAR = SS_PARAMETRIA_CURSOS_BBVA.getSheetByName(SHEET_ITINERARIO_DETALLE_REGULAR_BBVA_NOMBRE);


// Data Control Cursos Mandatorios Onboarding
const SS_CONTROL_CURSOS_BBVA_ID = "";
const SS_CONTROL_CURSOS_BBVA = SpreadsheetApp.openById(SS_CONTROL_CURSOS_BBVA_ID);
const SHEET_CONTROL_COLABORADORES_3M_NOMBRE = "ControlColaboradores3M";
const SHEET_CONTROL_COLABORADORES_3M = SS_CONTROL_CURSOS_BBVA.getSheetByName(SHEET_CONTROL_COLABORADORES_3M_NOMBRE);
const SHEET_CONTROL_COLABORADORES_NOMBRE = "ControlColaboradores";
const SHEET_CONTROL_COLABORADORES = SS_CONTROL_CURSOS_BBVA.getSheetByName(SHEET_CONTROL_COLABORADORES_NOMBRE);
const SHEET_CONTROL_DETALLADO_NOMBRE = "ControlDetallado";
const SHEET_CONTROL_DETALLADO_BBVA = SS_CONTROL_CURSOS_BBVA.getSheetByName(SHEET_CONTROL_DETALLADO_NOMBRE);
const SHEET_GRUPOS_FORMACION_NOMBRE = "GruposFormacion";


// Data Control Cursos Mandatorios Regulares
const SS_CONTROL_CURSOS_REGULARES_BBVA_ID = "";
const SS_CONTROL_CURSOS_REGULARES_BBVA = SpreadsheetApp.openById(SS_CONTROL_CURSOS_REGULARES_BBVA_ID);
const SHEET_CONTROL_DETALLADO_REGULARES_BBVA = SS_CONTROL_CURSOS_REGULARES_BBVA.getSheetByName(SHEET_CONTROL_DETALLADO_NOMBRE);
const SHEET_CONTROL_COLABORA_REGULARES_BBVA = SS_CONTROL_CURSOS_REGULARES_BBVA.getSheetByName(SHEET_CONTROL_COLABORADORES_NOMBRE);


// Datos Cornerstone
const DRIVE_FOLDER_ITINERARIO_COLABORADORES_ID = "";
const DRIVE_FOLDER_CURSOS_COLABORADORES_ID = "";




/** DATA CONTROL CURSOS */
function obtenerControlColaboradores(tipoProceso,procesoRegularActivo){

  var query = "SELECT * ";

  if(tipoProceso==TIPO_PROCESO_ONB){
    var resultado = executeQuery(SS_CONTROL_CURSOS_BBVA_ID, SHEET_CONTROL_COLABORADORES_NOMBRE, query, "A1:AD", false);

  }else if(tipoProceso==TIPO_PROCESO_REG){

    var resultado = executeQuery(SS_CONTROL_CURSOS_REGULARES_BBVA_ID, PREFIJO_HOJA_REGULAR+procesoRegularActivo, query, "A1:U", false);
  };
  return resultado;
}


/** DATA WORKDAY - para ambos procesos sacan la misma data nueva*/
function obtenerAltasWorkday(tipoProceso){

  //DONDE P - REGISTRO NO ESTE VACIO
  var query = "SELECT * WHERE C <> ''";

  if(tipoProceso===TIPO_PROCESO_ONB){
    var resultado = executeQuery(SS_ALTAS_BBVA_ID, SHEET_COLABORADORES_BBVA_NOMBRE, query, "A1:U", false);
  }else if(tipoProceso===TIPO_PROCESO_REG){
    var resultado = executeQuery(SS_ALTAS_BBVA_ID, SHEET_COLABORADORES_BBVA_ACTUALES_NOMBRE, query, "A1:U", false);
  }

  return resultado;
}

function obtenerItinerariosConSusCursos(tipoProceso) {
  try {
    var hojaGrupoCursos;
    if (tipoProceso === TIPO_PROCESO_ONB) {
      hojaGrupoCursos = SpreadsheetApp.openById(SS_PARAMETRIA_CURSOS_ID).getSheetByName(SHEET_ITINERARIO_DETALLE_BBVA_NOMBRE);
    } else if (tipoProceso === TIPO_PROCESO_REG) {
      hojaGrupoCursos = SpreadsheetApp.openById(SS_PARAMETRIA_CURSOS_ID).getSheetByName(SHEET_ITINERARIO_DETALLE_REGULAR_BBVA_NOMBRE);
    } else {
      throw new Error("Tipo de proceso no reconocido: " + tipoProceso);
    }

    if (!hojaGrupoCursos) {
      throw new Error("No se encontró la hoja correspondiente para el tipo de proceso: " + tipoProceso);
    }

    var rangoDatos = hojaGrupoCursos.getDataRange().getValues();
    var grupoCursos = {}; // objeto, para almacenar grupo y curso

    for (var i = 0; i < rangoDatos[0].length; i++) {
      var nombreGrupo = rangoDatos[0][i]; // encabezado de grupo
      grupoCursos[nombreGrupo] = [];
      for (var j = 1; j < rangoDatos.length; j++) {
        if (rangoDatos[j][i]) {
          grupoCursos[nombreGrupo].push(rangoDatos[j][i]);
        }
      }
    }

    return grupoCursos;
  } catch (error) {
    console.error("Error en obtenerGrupoCursos: " + error.message);
    throw error;
  }
}

function obtenerGrupoCursos(tipoProceso) {
  try {
    var hojaGrupoCursos;
    if (tipoProceso === TIPO_PROCESO_ONB) {
      hojaGrupoCursos = SpreadsheetApp.openById(SS_PARAMETRIA_CURSOS_ID).getSheetByName("GruposCursosOnb");
    } else if (tipoProceso === TIPO_PROCESO_REG) {
      hojaGrupoCursos = SpreadsheetApp.openById(SS_PARAMETRIA_CURSOS_ID).getSheetByName("GruposCursosReg");
    } else {
      throw new Error("Tipo de proceso no reconocido: " + tipoProceso);
    }

    if (!hojaGrupoCursos) {
      throw new Error("No se encontró la hoja correspondiente para el tipo de proceso: " + tipoProceso);
    }

    var rangoDatos = hojaGrupoCursos.getDataRange().getValues();
    var grupoCursos = {}; // objeto, para almacenar grupo y curso

    for (var i = 0; i < rangoDatos[0].length; i++) {
      var nombreGrupo = rangoDatos[0][i]; // encabezado de grupo
      grupoCursos[nombreGrupo] = [];
      for (var j = 1; j < rangoDatos.length; j++) {
        if (rangoDatos[j][i]) {
          grupoCursos[nombreGrupo].push(rangoDatos[j][i]);
        }
      }
    }

    return grupoCursos;
  } catch (error) {
    console.error("Error en obtenerGrupoCursos: " + error.message);
    throw error;
  }
}


/** DATA DE GRUPO PERTENECIENTE POR FECHA DE INGRESO */
function obtenerGruposFormacion(){
  // Elaborar query
  var query = "SELECT * ";
  // Ejecutar query 
  var resultado = executeQuery(SS_PARAMETRIA_CURSOS_ID, SHEET_GRUPOS_FORMACION_NOMBRE, query, "A1:C", false);

  return resultado;
}

function obtenerEncabezadosCursos(tipoProceso,codigoProcesoRegular,dataItinerariosDetalle){
  if(tipoProceso == TIPO_PROCESO_ONB){
      var encabezadosControlDetalle = SHEET_CONTROL_DETALLADO_BBVA.getRange(1, SHEET_ITINERARIO_DET_CURSO_DELIMITADOR_ID_IDX + 1, 1, SHEET_CONTROL_DETALLADO_BBVA.getLastColumn()-4).getValues()[0];
  }else if(tipoProceso == TIPO_PROCESO_REG){
      var query = "SELECT A WHERE G = '"+TIPO_PROCESO_REG+"' AND I = '"+codigoProcesoRegular +"' AND E = '"+ TIPO_ESTADO_ACTIVO + "'";
      var itons = executeQuery(SS_PARAMETRIA_CURSOS_ID, SHEET_ITINERARIO_BBVA_NOMBRE, query, "A1:I", false);
      var encabezadosControlDetalle = obtenerCursosUnicos(itons,dataItinerariosDetalle);

  }

  return encabezadosControlDetalle;
}


/** DATA DE ITINERARIOS */
function obtenerDataItinerarios(tipoProceso){
  var query = "SELECT * WHERE G = '"+tipoProceso+"' AND E = '"+TIPO_ESTADO_ACTIVO+"'";
  if(tipoProceso == TIPO_PROCESO_ONB){
    var resultado = executeQuery(SS_PARAMETRIA_CURSOS_ID, SHEET_ITINERARIO_BBVA_NOMBRE, query, "A1:I", false);
  }else if(tipoProceso == TIPO_PROCESO_REG){
    var resultado = executeQuery(SS_PARAMETRIA_CURSOS_ID, SHEET_ITINERARIO_BBVA_NOMBRE, query, "A1:I", false);
  }
  return resultado;
}

function obtenerProcesoRegularActivo(){
  var query = "SELECT I WHERE G = '"+ TIPO_PROCESO_REG+"'AND E = '"+ TIPO_ESTADO_ACTIVO+"'";
  var resultado = executeQuery(SS_PARAMETRIA_CURSOS_ID, SHEET_ITINERARIO_BBVA_NOMBRE, query, "A1:I", false);

  //quitar duplicados
  var procesosUnicos = Array.from(new Set(resultado.flat()));

  return procesosUnicos;
}

function obtenerDataOnboardingRRLL(){
  var query = "SELECT * WHERE K = '" + ESTADO_INSCRIPCION_REGISTRADO + "' AND AA = '' AND ( V >= " + DIAS_NOTIFICACION_RRLL + ")";
  var resultado = executeQuery(SS_CONTROL_CURSOS_BBVA_ID, SHEET_CONTROL_COLABORADORES_NOMBRE,query, "A1:AA", false);

  return resultado;
}

/** DATA CURSOS POR ITINERARIOS HOJA GUIA */
function obtenerItinerariosDetalle(tipoProceso) {

  if(tipoProceso == TIPO_PROCESO_ONB){
    var rangoDatos = SHEET_PARAMETRIA_ITINERARIO_DETALLE.getDataRange();
    var resultado = rangoDatos.getValues();

  }else if(tipoProceso == TIPO_PROCESO_REG){
    var rangoDatos = SHEET_PARAMETRIA_ITINERARIO_DETALLE_REGULAR.getDataRange();
    var resultado = rangoDatos.getValues();
  }
  
  return resultado;
}

function construirMapaItinerarios(listaCursosItinerarios) {
  var mapaItinerarios = {};

  // Recorrer la primera fila para encontrar los codigos de los itinerarios
  for (var i = 0; i < listaCursosItinerarios[0].length; i++) {
    var codigoPrograma = listaCursosItinerarios[0][i];
    var cursos = [];

    // Recopilar todos los cursos del itinerario
    for (var j = 1; j < listaCursosItinerarios.length; j++) {
      if (listaCursosItinerarios[j][i]) {
        cursos.push(listaCursosItinerarios[j][i]);
      }
    }

    // Agregar el itinerario y sus cursos al mapa
    mapaItinerarios[codigoPrograma] = cursos;
  }
  return mapaItinerarios;
}

function obtenerEncabezadosItons(tipoProceso){
  var encabezadosItons = obtenerItinerariosDetalle(tipoProceso);

  encabezadosItons = encabezadosItons[0];

  return encabezadosItons;

}




/** GUARDAR NUEVOS INGRESOS EN HOJA CONTROL */
function guardarDatosNuevaAltaEnLista(itemNuevaAlta, codigoPrograma, listaGruposFormacion, tipoProceso, fechaPrimerDia,fechaLimitePrograma) {
  
  if(tipoProceso == TIPO_PROCESO_REG){
    var fechaIngresoColaborador = convertirStringADate(itemNuevaAlta[SHEET_ALTAS_FECHA_INGRESO_IDX]);

    //MODIFICAR FECHA LIMITE A LA FECHA FIN DEL ITINERARIO
    //var fechaIngresoColaborador = convertirStringADate(itemNuevaAlta[SHEET_ALTAS_FECHA_INGRESO_IDX]);
    //var fechaLimitePrograma = obtenerFechaMasNDias(fechaIngresoColaborador, CANT_DIAS_LIMITE_PROGRAMA);

    var nombrePreferido = itemNuevaAlta[SHEET_ALTAS_NOMBRE_PREFERIDO_IDX];
    var nombreNormal = itemNuevaAlta[SHEET_ALTAS_NOMBRE_COMPLETO_IDX];
    var nombreElegido = nombrePreferido == "" ? nombreNormal : nombrePreferido;
    
    var datosNuevaAlta = [
      "=ROW()",
      new Date(),
      itemNuevaAlta[SHEET_ALTAS_REGISTRO_IDX],
      nombreElegido,         //itemNuevaAlta[SHEET_ALTAS_NOMBRE_COMPLETO_IDX],
      itemNuevaAlta[SHEET_ALTAS_EMAIL_IDX],
      fechaIngresoColaborador,
      itemNuevaAlta[SHEET_ALTAS_AREA_CONVERSION_KEIJI_IDX],
      itemNuevaAlta[SHEET_ALTAS_UNIDAD_IDX],
      itemNuevaAlta[SHEET_ALTAS_SUBUNIDAD_IDX],
      codigoPrograma,
      ESTADO_INSCRIPCION_PENDIENTE,
      null, // fecha inscripción programa
      fechaLimitePrograma,
      ESTADO_CURSO_FINAL_PENDIENTE,
      null, // fecha término cursos
      null, // PENDIENTE
      itemNuevaAlta[SHEET_ALTAS_MANAGER_EMAIL_IDX],
      itemNuevaAlta[SHEET_ALTAS_EMPRESA_IDX],
      itemNuevaAlta[SHEET_ALTAS_SUPERVISORY_IDX],
      itemNuevaAlta[SHEET_ALTAS_CENTRO_COSTOS_IDX],
      itemNuevaAlta[SHEET_ALTAS_JOB_PROFILE_IDX],
      null, null, null // Fecha avisos
    ];
  }else if(tipoProceso = TIPO_PROCESO_ONB){
    fechaPrimerDia = convertirStringADate(fechaPrimerDia);
    var fechaIngresoColaborador = convertirStringADate(itemNuevaAlta[SHEET_ALTAS_FECHA_INGRESO_IDX]);
    var fechaFinAlta = convertirStringADate(itemNuevaAlta[SHEET_ALTAS_FECHA_FIN_ALTA_IDX]);

    //SE CAMBIA LA FECHA + N DIAS, PORQUE AHORA EL TIEMPO LIMITE ES EN BASE A LA FECHA DE INSCRIPCION, NO EN LA FECHA DE INGRESO
    //var fechaLimitePrograma = obtenerFechaMasNDias(fechaPrimerDia, CANT_DIAS_LIMITE_PROGRAMA);
    var grupoFormacion = obtenerGrupoFormacion(fechaPrimerDia, listaGruposFormacion);

    var nombrePreferido = itemNuevaAlta[SHEET_ALTAS_NOMBRE_PREFERIDO_IDX];
    var nombreNormal = itemNuevaAlta[SHEET_ALTAS_NOMBRE_COMPLETO_IDX];
    var nombreElegido = nombrePreferido == "" ? nombreNormal : nombrePreferido;
    
    var datosNuevaAlta = [
      "=ROW()",
      new Date(),
      itemNuevaAlta[SHEET_ALTAS_REGISTRO_IDX],
      nombreElegido,         //itemNuevaAlta[SHEET_ALTAS_NOMBRE_COMPLETO_IDX],
      itemNuevaAlta[SHEET_ALTAS_EMAIL_IDX],
      fechaIngresoColaborador,
      itemNuevaAlta[SHEET_ALTAS_AREA_CONVERSION_KEIJI_IDX],
      itemNuevaAlta[SHEET_ALTAS_UNIDAD_IDX],
      itemNuevaAlta[SHEET_ALTAS_SUBUNIDAD_IDX],
      codigoPrograma,
      ESTADO_INSCRIPCION_PENDIENTE,
      null, // fecha inscripción programa
      null, //fecha limite culminacion programa
      ESTADO_CURSO_FINAL_PENDIENTE,
      null, // fecha término cursos
      grupoFormacion, // PENDIENTE
      itemNuevaAlta[SHEET_ALTAS_MANAGER_EMAIL_IDX],
      itemNuevaAlta[SHEET_ALTAS_EMPRESA_IDX],
      itemNuevaAlta[SHEET_ALTAS_SUPERVISORY_IDX],
      itemNuevaAlta[SHEET_ALTAS_CENTRO_COSTOS_IDX],
      itemNuevaAlta[SHEET_ALTAS_JOB_PROFILE_IDX],
      null, null, null,null,null,null,null,null, // Fecha avisos
      fechaPrimerDia, fechaFinAlta 
    ];

  }
  

  return datosNuevaAlta;
}

function guardarDatosNuevaAltaEnListaDetallado(itemNuevaAlta, codigoPrograma, mapaItinerarios, encabezados, tipoProceso) {
  var colaboradorID = itemNuevaAlta[SHEET_ALTAS_REGISTRO_IDX];
  var cursosItinerario = mapaItinerarios[codigoPrograma] || null;

  var nuevoRegistro = Array(encabezados.length + 4).fill('');
  nuevoRegistro[SHEET_CONTROL_DET_FILA_IDX] = "=ROW()";
  nuevoRegistro[SHEET_CONTROL_DET_COLABORADOR_ID_IDX] = colaboradorID;
  nuevoRegistro[SHEET_CONTROL_DET_ITINERARIO_ID_IDX] = codigoPrograma;
  nuevoRegistro[SHEET_CONTROL_DET_ESTADO_IDX] = ESTADO_CURSO_FINAL_PENDIENTE;

  if (cursosItinerario != null) {
    for (var i = 0; i < encabezados.length; i++) {
      var curso = encabezados[i];
      if (cursosItinerario.indexOf(curso) !== -1) {
        nuevoRegistro[i+4] = ESTADO_CURSO_PENDIENTE;
      } else {
        nuevoRegistro[i+4] = ESTADO_CURSO_NO_APLICA;
      }
    }
  }

  return nuevoRegistro;
}



/*
function obtenerColaboradorNuevaAlta(listaNuevasAltas, colaboradorID){
  var nuevaAltaColaborador = null;
  for(var i in listaNuevasAltas){
    var itemNuevaAlta = listaNuevasAltas[i];
    var nuevaAltaColaboradorID = itemNuevaAlta[SHEET_ALTAS_REGISTRO_IDX];
    if(nuevaAltaColaboradorID==colaboradorID){
      nuevaAltaColaborador = itemNuevaAlta;
      break;
    }
  }
  return nuevaAltaColaborador;
}
*/
/*
function obtenerColaboradorNuevaAlta(mapaColaboradores, colaboradorID) {
  return mapaColaboradores.get(colaboradorID) || null;
}
*/

function guardarControlEmailColaborador(filaRegistro, emailColaborador,tipoProceso,procesoRegularActivo){
  if(tipoProceso == TIPO_PROCESO_ONB){
    SHEET_CONTROL_COLABORADORES.getRange(filaRegistro, SHEET_CONTROL_COLAB_COLABORADOR_EMAIL_IDX+1).setValue(emailColaborador); 
  } else if(tipoProceso == TIPO_PROCESO_REG){
    var hojaControlRegular = SS_CONTROL_CURSOS_REGULARES_BBVA.getSheetByName(PREFIJO_HOJA_REGULAR + procesoRegularActivo);
    hojaControlRegular.getRange(filaRegistro, SHEET_CONTROL_COLAB_COLABORADOR_EMAIL_IDX+1).setValue(emailColaborador); 
  }
}


function obtenerDataControlColaboradoresPendientes(tipoProceso,procesoRegularActivo) {

  var query = "SELECT * where K<>'"+ESTADO_INSCRIPCION_TERMINADO+"' OR N='"+ESTADO_CURSO_FINAL_PENDIENTE+"'";
  if(tipoProceso == TIPO_PROCESO_ONB){
    var resultado = executeQuery(SS_CONTROL_CURSOS_BBVA_ID, SHEET_CONTROL_COLABORADORES_NOMBRE, query, "A1:AE", false);
  }else if(tipoProceso == TIPO_PROCESO_REG){
    var resultado = executeQuery(SS_CONTROL_CURSOS_REGULARES_BBVA_ID, PREFIJO_HOJA_REGULAR + procesoRegularActivo, query, "A1:V", false);
  }
  return resultado;
}


function guardarControlColaboradorEstadoMatricula(filaRegistro, estadoItinerario, fechaInscripcion, tipoProceso,procesoRegularActivo){
  if(tipoProceso == TIPO_PROCESO_ONB){
    SHEET_CONTROL_COLABORADORES.getRange(filaRegistro, SHEET_CONTROL_COLAB_ESTADO_MATRICULA_IDX+1, 1, 2).setValues([[estadoItinerario, fechaInscripcion]]);  
  }else if(tipoProceso == TIPO_PROCESO_REG){
    var hojaControlRegular = SS_CONTROL_CURSOS_REGULARES_BBVA.getSheetByName(PREFIJO_HOJA_REGULAR + procesoRegularActivo);
    hojaControlRegular.getRange(filaRegistro,SHEET_CONTROL_COLAB_ESTADO_MATRICULA_IDX+1,1,2 ).setValues([[estadoItinerario, fechaInscripcion]]);
  }
}


function guardarEstadoEliminadoControlColaborador(filaRegistro, estadoItinerario, fechaActual, tipoProceso, procesoRegularActivo) {
  if (tipoProceso == TIPO_PROCESO_ONB) {
    SHEET_CONTROL_COLABORADORES.getRange(filaRegistro, SHEET_CONTROL_COLAB_ESTADO_MATRICULA_IDX + 1,1,2).setValues([[estadoItinerario, ""]]);
    SHEET_CONTROL_COLABORADORES.getRange(filaRegistro, SHEET_CONTROL_COLAB_FECHA_REGISTRO_IDX + 1).setValue(fechaActual);
  } else if (tipoProceso == TIPO_PROCESO_REG) {
    var hojaControlRegular = SS_CONTROL_CURSOS_REGULARES_BBVA.getSheetByName(PREFIJO_HOJA_REGULAR + procesoRegularActivo);
    hojaControlRegular.getRange(filaRegistro, SHEET_CONTROL_COLAB_ESTADO_MATRICULA_IDX + 1,1,2).setValues([[estadoItinerario, ""]]);
    hojaControlRegular.getRange(filaRegistro, SHEET_CONTROL_COLAB_FECHA_REGISTRO_IDX  + 1).setValue(fechaActual);
  }
}



function obtenerEncabezadosItinerarios(tipoProceso){
  if(tipoProceso == TIPO_PROCESO_ONB){
    var encabezadosItinerarioDetalle = SHEET_PARAMETRIA_ITINERARIO_DETALLE.getRange(1, 1, 1, SHEET_PARAMETRIA_ITINERARIO_DETALLE.getLastColumn()).getValues()[0];
    console.log("encabezadosItinerarioDetalle: "+encabezadosItinerarioDetalle);

  }else if(tipoProceso == TIPO_PROCESO_REG){
    var encabezadosItinerarioDetalle = SHEET_PARAMETRIA_ITINERARIO_DETALLE_REGULAR.getRange(1, 1, 1, SHEET_PARAMETRIA_ITINERARIO_DETALLE_REGULAR.getLastColumn()).getValues()[0];
    console.log("encabezadosItinerarioDetalle: "+encabezadosItinerarioDetalle);
  }

  return encabezadosItinerarioDetalle;
}


function obtenerPendientesMatriculaPorItinerario(listaDataPendienteControl, itinerarioID){
  var listaResultado = new Array();
  for(var i in listaDataPendienteControl){
    var itemData = listaDataPendienteControl[i];
    var itemEstadoMatricula = itemData[SHEET_CONTROL_COLAB_ESTADO_MATRICULA_IDX];
    var itemItinerarioID = itemData[SHEET_CONTROL_COLAB_ITINERARIO_CODIGO_IDX].trim();
    if(ESTADO_INSCRIPCION_TERMINADO!=itemEstadoMatricula && itinerarioID==itemItinerarioID){
      listaResultado.push(itemData);
    }
  }
  return listaResultado;
}

/** DATA CORNERSTONE */
function obtenerDataCornerstoneItinerario(itinerarioID, itinerarioFileID,tipoProceso){

  var query = "SELECT * WHERE B='"+itinerarioID+"'";
  if(tipoProceso == TIPO_PROCESO_ONB){
    var nombreHojaItinerario = itinerarioID+SHEET_CORNERSTONE_SUFIJO;
  }else if(tipoProceso == TIPO_PROCESO_REG){
    var nombreHojaItinerario = itinerarioID+SHEET_CORNERSTONE_SUFIJO_REGULAR;
  }
  
  var resultado = executeQuery(itinerarioFileID, nombreHojaItinerario, query, "A1:F", false);
  return resultado;
}


function obtenerDataControlCursosPendientes(tipoProceso,procesoRegularActivo) {
  var query = "SELECT * where D='"+ESTADO_CURSO_FINAL_PENDIENTE+"'";

  if(tipoProceso == TIPO_PROCESO_ONB){
    var resultado = executeQuery(SS_CONTROL_CURSOS_BBVA_ID, SHEET_CONTROL_DETALLADO_NOMBRE, query, "A1:AY", false);
    console.log("resultado="+resultado.length);

  }else if(tipoProceso == TIPO_PROCESO_REG){
    var resultado = executeQuery(SS_CONTROL_CURSOS_REGULARES_BBVA_ID, PREFIJO_HOJA_REGULAR_DETALLADO + procesoRegularActivo, query, "A1:Z", false);
    console.log("resultado="+resultado.length);
  }
  return resultado;
}


function obtenerDataCornerstoneGrupo(grupoFileID) {

    var abrirHoja = SpreadsheetApp.openById(grupoFileID);
    var hoja = abrirHoja.getSheets()[0];

    // Obtener el rango de datos desde la fila 10 hasta la última fila y última columna
    var primeraFila = 10;
    var ultimaFila = hoja.getLastRow();
    var ultimaColumna = hoja.getLastColumn();
    var dataRange = hoja.getRange(primeraFila, 1, ultimaFila - primeraFila + 1, ultimaColumna);
    var data = dataRange.getValues();

    // Retorna todos los datos desde la fila 10 (excluyendo los encabezados de la fila 9)
    return data;
}


function guardarControlColaboradorEstadoCurso(filaRegistro, cursoID, estadoCurso, tipoProceso) {
    // Obtener los encabezados de la hoja de control detallado
    var hoja = tipoProceso === TIPO_PROCESO_ONB ? SHEET_CONTROL_DETALLADO_BBVA : SHEET_CONTROL_DETALLADO_REGULARES_BBVA;
    var encabezados = hoja.getRange(1, 1, 1, hoja.getLastColumn()).getValues()[0];

    // Encontrar el índice de la columna del curso
    var sheetControlCursoIdx = encabezados.indexOf(cursoID);

    // Si se encuentra el índice, actualizar el estado del curso
    if (sheetControlCursoIdx != -1) {
        hoja.getRange(filaRegistro, sheetControlCursoIdx + 1).setValue(estadoCurso);
    }
}

function cargarDatosDeTodosLosGrupos(grupoCursos) {
    var datosGrupos = {};
    for (var grupo in grupoCursos) {
        var cursosGrupo = grupoCursos[grupo];
        var cursoFileID = obtenerFilePorGrupoID(grupo);
        if (cursoFileID != null) {
            var listaCornerstoneCursos = obtenerDataCornerstoneGrupo(cursoFileID);
            cursosGrupo.forEach(function(cursoID) {
                datosGrupos[cursoID] = listaCornerstoneCursos;
            });
        }
    }
    return datosGrupos;
}

function updateFila(filaRegistro, columnaCurso, nuevoEstado, hoja) {
  if (filaRegistro > 0 && columnaCurso > 0) {
    hoja.getRange(filaRegistro, columnaCurso + 1).setValue(nuevoEstado);
  }else{
    console.log(`Fila o columna inválida: fila ${filaRegistro}, columna ${columnaCurso}`);
  }
}

function guardarControlDetalladoEstadoCursos(filaRegistro, estadoCursos,tipoProceso,procesoRegularActivo){
  if(tipoProceso == TIPO_PROCESO_ONB){
    SHEET_CONTROL_DETALLADO_BBVA.getRange(filaRegistro, SHEET_CONTROL_DET_ESTADO_IDX+1).setValue(estadoCursos);  
  }else if(tipoProceso == TIPO_PROCESO_REG){
    var hojaControlRegular = SS_CONTROL_CURSOS_REGULARES_BBVA.getSheetByName(PREFIJO_HOJA_REGULAR_DETALLADO + procesoRegularActivo);
    hojaControlRegular.getRange(filaRegistro, SHEET_CONTROL_DET_ESTADO_IDX+1).setValue(estadoCursos); 

  }
}

function guardarControlColaboradorEstadoCursos(filaRegistro, estadoCursos,tipoProceso,procesoRegularActivo){
  if(tipoProceso == TIPO_PROCESO_ONB){
    SHEET_CONTROL_COLABORADORES.getRange(filaRegistro, SHEET_CONTROL_COLAB_ESTADO_CURSOS_IDX+1, 1, 2).setValues([[estadoCursos, new Date()]]); 
  }else if(tipoProceso == TIPO_PROCESO_REG){
    var hojaControlRegular = SS_CONTROL_CURSOS_REGULARES_BBVA.getSheetByName(PREFIJO_HOJA_REGULAR + procesoRegularActivo);
    hojaControlRegular.getRange(filaRegistro, SHEET_CONTROL_COLAB_ESTADO_CURSOS_IDX+1, 1, 2).setValues([[estadoCursos, new Date()]]);
  } 
}

//Proceso solo para regular
function obtenerItinerariosPorCodigoProceso(codigoProcesoRegular){
  var query = "SELECT * WHERE I = '"+ codigoProcesoRegular + "' AND E = '"+ TIPO_ESTADO_ACTIVO + "'" ;
  var resultado = executeQuery(SS_PARAMETRIA_CURSOS_ID,SHEET_ITINERARIO_BBVA_NOMBRE,query,"A1:I",false);

  return resultado;
}
/*
function obtenerDataControlPendientesRecordatorio(tipoProceso,procesoRegularActivo) {
  if(tipoProceso == TIPO_PROCESO_ONB){
    var query = "SELECT * where K='"+ESTADO_INSCRIPCION_REGISTRADO+"' and (V="+DIAS_RECORDATORIO_COLAB_1+" or V="+DIAS_RECORDATORIO_COLAB_2+" or V="+DIAS_RECORDATORIO_COLAB_3+")";
    var resultado = executeQuery(SS_CONTROL_CURSOS_BBVA_ID, SHEET_CONTROL_COLABORADORES_NOMBRE, query, "A1:AA", false);
  }
  else if(tipoProceso == TIPO_PROCESO_REG){

    var query = "SELECT * WHERE K = '" + ESTADO_INSCRIPCION_REGISTRADO + "' AND ( V = " + DIAS_RECORDATORIO_REGULAR_COLAB_1 + " OR V = " + DIAS_RECORDATORIO_REGULAR_COLAB_2 + " OR V = " + DIAS_RECORDATORIO_REGULAR_COLAB_3 + " OR V = " + DIAS_NOTIFICACION_REGULAR_ADVISOR_2 + ")";
    var resultado = executeQuery(SS_CONTROL_CURSOS_REGULARES_BBVA_ID, PREFIJO_HOJA_REGULAR + procesoRegularActivo, query, "A1:AD", false);
  }
  return resultado;
}*/

function obtenerDataControlPendientesRecordatorio(tipoProceso,procesoRegularActivo) {
  if(tipoProceso == TIPO_PROCESO_ONB){
    var query = "SELECT * where K='"+ESTADO_INSCRIPCION_REGISTRADO+"' and (V="+DIAS_RECORDATORIO_COLAB_1+" or V="+DIAS_RECORDATORIO_COLAB_2+" or V="+DIAS_RECORDATORIO_COLAB_3+")";
    var resultado = executeQuery(SS_CONTROL_CURSOS_BBVA_ID, SHEET_CONTROL_COLABORADORES_NOMBRE, query, "A1:AA", false);
  }
  else if(tipoProceso == TIPO_PROCESO_REG){

    //var query = "SELECT * WHERE K = '" + ESTADO_INSCRIPCION_REGISTRADO + "' AND ( V = " + DIAS_RECORDATORIO_REGULAR_COLAB_1 + " OR V = " + DIAS_RECORDATORIO_REGULAR_COLAB_2 + " OR V = " + DIAS_RECORDATORIO_REGULAR_COLAB_3 + " OR V = " + DIAS_NOTIFICACION_REGULAR_ADVISOR_2 + ")";

    var query =  "SELECT * WHERE K = '" + ESTADO_INSCRIPCION_REGISTRADO + "' AND ( V = " + DIAS_RECORDATORIO_REGULAR_COLAB_1 + " OR V = " + DIAS_RECORDATORIO_REGULAR_COLAB_2 + " OR V = " + DIAS_RECORDATORIO_REGULAR_COLAB_3 + " OR V = " + DIAS_NOTIFICACION_REGULAR_ADVISOR_2 + " OR V = " + DIAS_RECORDATORIO_REGULAR_COLAB_28 +" OR V = "  +DIAS_RECORDATORIO_REGULAR_COLAB_29 +")";
    var resultado = executeQuery(SS_CONTROL_CURSOS_REGULARES_BBVA_ID, PREFIJO_HOJA_REGULAR + procesoRegularActivo, query, "A1:AD", false);
  }
  return resultado;
}

function obtenerDataControlPendientesRecordatorioRRLL(procesoRegularActivo){
  var query = "SELECT * WHERE K = '" + ESTADO_INSCRIPCION_REGISTRADO + "' AND ( V >= "+DIAS_NOTIFICACION_REGULAR_RRLL+")";
  var resultado = executeQuery(SS_CONTROL_CURSOS_REGULARES_BBVA_ID, PREFIJO_HOJA_REGULAR + procesoRegularActivo, query, "A1:AD", false);

  return resultado;
}

function guardarControlEnvioRecordatorio(filaRegistro, columnaAviso,procesoRegularActivo){
  var hojaControlRegular = SS_CONTROL_CURSOS_REGULARES_BBVA.getSheetByName(PREFIJO_HOJA_REGULAR + procesoRegularActivo);
  hojaControlRegular.getRange(filaRegistro, columnaAviso+1).setValue(new Date());  
}

function obtenerImagenDrive(nombreCarpeta, nombreArchivo) {
  var archivo;
  var carpetas = DriveApp.getFoldersByName(nombreCarpeta);
  while (carpetas.hasNext()) {
    var carpeta = carpetas.next();
    var archivos = carpeta.getFilesByName(nombreArchivo);
    while (archivos.hasNext()) {
      archivo = archivos.next().getAs("image/jpeg");
    }
  }

  return archivo;
}

function guardarEnvioCorreoRRLL(filaRegistro, columnaAviso){
  SHEET_CONTROL_COLABORADORES.getRange(filaRegistro, columnaAviso+1).setValue(new Date());  
}

/****************************/
/***** QUERYS GENERALES *****/
/****************************/

// Query generico para busquedas en archivos
function executeQuery(spreadSheetId, sheetName, queryFormula, range, showHeader) {
  var lastRow = SpreadsheetApp.openById(spreadSheetId).getSheetByName(sheetName).getLastRow();
  var qvizURL = 'https://docs.google.com/spreadsheets/d/' + spreadSheetId
    + '/gviz/tq?tqx=out:json&headers=1&sheet=' + sheetName
    + '&range=' + range + lastRow
    + '&tq=' + encodeURIComponent(queryFormula);
  //console.log("qvizURL=" + qvizURL);
  var ret = UrlFetchApp.fetch(qvizURL, { headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() } }).getContentText();
  var resp = JSON.parse(ret.replace("/*O_o*/", "").replace("google.visualization.Query.setResponse(", "").slice(0, -2));
  var data = resp.table.rows.map(row => {
    return row.c.map(cols => {
      return cols === null ? '' : cols.f !== undefined ? cols.f : (cols.v === null ? '' : cols.v);
    });
  });
  if (showHeader) {
    var header = resp.table.cols.map(col => {
      return col.label;
    });
    return [header].concat(data);
  } else {
    return data;
  }
}
