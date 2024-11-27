// Indices SHEET Nuevas Altas Workday
const SHEET_ALTAS_POSICION_IDX = 1;
const SHEET_ALTAS_REGISTRO_IDX = 2;
const SHEET_ALTAS_NOMBRE_COMPLETO_IDX = 3;
const SHEET_ALTAS_NOMBRE_PREFERIDO_IDX = 4;
const SHEET_ALTAS_EMAIL_IDX = 5;
const SHEET_ALTAS_FECHA_INGRESO_IDX = 6;
const SHEET_ALTAS_JOB_PROFILE_IDX = 7;
const SHEET_ALTAS_AREA_IDX = 8;
const SHEET_ALTAS_UNIDAD_IDX = 9;
const SHEET_ALTAS_SUBUNIDAD_IDX = 10;
const SHEET_ALTAS_CENTRO_COSTOS_IDX = 11;
const SHEET_ALTAS_SUPERVISORY_IDX = 12;
const SHEET_ALTAS_MANAGER_NOMBRE_IDX = 13;
const SHEET_ALTAS_MANAGER_EMAIL_IDX = 14;
const SHEET_ALTAS_EMPRESA_IDX = 15;
const SHEET_ALTAS_TIPO_POSICION_IDX = 16;
const SHEET_ALTAS_WORKFORCE_PLANNING_IDX = 17;
const SHEET_ALTAS_SUPERVISORY_9_IDX = 18;
const SHEET_ALTAS_AREA_CONVERSION_KEIJI_IDX = 19;
const SHEET_ALTAS_FECHA_FIN_ALTA_IDX = 20;


// Indices Sheet Itinerario
const SHEET_ITINERARIO_CODIGO_IDX = 0;
const SHEET_ITINERARIO_FECHA_INI_IDX = 2;
const SHEET_ITINERARIO_FECHA_FIN_IDX = 3;
const SHEET_ITINERARIO_TIPO_PROGRAMA_IDX = 4;
const SHEET_ITINERARIO_TIPO_ESTADO_IDX = 4; /** INACTIVO, ACTIVO */
const SHEET_ITINERARIO_TIPO_COLECTIVO_IDX = 5; /**SEDE 1, SEDE 2, RED 1, RED 2, ...ETC */
const SHEET_ITINERARIO_TIPO_PROCESO_IDX = 6;
const SHEET_ITINERARIO_DETALLE_COLECTIVO_IDX = 7; 
const SHEET_ITINERARIO_CODIGO_PROCESO_IDX = 8;

// Indices Sheet Itinerario Detalle
const SHEET_ITINERARIO_DET_ITINERARIO_ID_IDX = 0;
const SHEET_ITINERARIO_DET_CURSO_ID_IDX = 1;
const SHEET_ITINERARIO_DET_CURSO_DELIMITADOR_ID_IDX = 4;

// Indices Sheet Control Colaboradores
const SHEET_CONTROL_COLAB_FILA_IDX = 0;
const SHEET_CONTROL_COLAB_FECHA_REGISTRO_IDX = 1;
const SHEET_CONTROL_COLAB_COLABORADOR_ID_IDX = 2;
const SHEET_CONTROL_COLAB_COLABORADOR_NOMBRE_IDX = 3;
const SHEET_CONTROL_COLAB_COLABORADOR_EMAIL_IDX = 4;
const SHEET_CONTROL_COLAB_FECHA_INGRESO_IDX = 5;
const SHEET_CONTROL_COLAB_AREA_IDX = 6;
const SHEET_CONTROL_COLAB_ITINERARIO_CODIGO_IDX = 9;
const SHEET_CONTROL_COLAB_ESTADO_MATRICULA_IDX = 10;
const SHEET_CONTROL_COLAB_ESTADO_CURSOS_IDX = 13;
const SHEET_CONTROL_COLAB_EMAIL_MANAGER_IDX = 16;
const SHEET_CONTROL_COLAB_DIAS_TRANSCURRIDOS_IDX = 21;
const SHEET_CONTROL_COLAB_RECORDATORIO_1_IDX = 22;
const SHEET_CONTROL_COLAB_RECORDATORIO_2_IDX = 23;
const SHEET_CONTROL_COLAB_RECORDATORIO_3_IDX = 24;
const SHEET_CONTROL_COLAB_ESCALA_ADVISOR_IDX = 25;
const SHEET_CONTROL_COLAB_ESCALA_RRLL_IDX = 26;
const SHEET_CONTROL_COLAB_FECHA_1ER_DIA_IDX =29;
const SHEET_CONTROL_COLAB_FECHA_REGISTRO_WD_IDX = 30;

const SHEET_CONTROL_COLAB_AVISO_ADVISOR_2_REGULAR_IDX = 29;
// Indices Sheet Control Detallado cursos
const SHEET_CONTROL_DET_FILA_IDX = 0;
const SHEET_CONTROL_DET_COLABORADOR_ID_IDX = 1;
const SHEET_CONTROL_DET_ITINERARIO_ID_IDX = 2;
const SHEET_CONTROL_DET_ESTADO_IDX = 3;
const SHEET_CONTROL_DET_CURSO_1_IDX = 4;
const SHEET_CONTROL_DET_CURSO_2_IDX = 5;
const SHEET_CONTROL_DET_CURSO_3_IDX = 6;
const SHEET_CONTROL_DET_CURSO_4_IDX = 7;
const SHEET_CONTROL_DET_CURSO_5_IDX = 8;
const SHEET_CONTROL_DET_CURSO_6_IDX = 9;
const SHEET_CONTROL_DET_CURSO_7_IDX = 10;
const SHEET_CONTROL_DET_CURSO_8_IDX = 11;
const SHEET_CONTROL_DET_CURSO_9_IDX = 12;
const SHEET_CONTROL_DET_CURSO_10_IDX = 13;
const SHEET_CONTROL_DET_CURSO_11_IDX = 14;
const SHEET_CONTROL_DET_CURSO_12_IDX = 15;
const SHEET_CONTROL_DET_CURSO_13_IDX = 16;
const SHEET_CONTROL_DET_CURSO_14_IDX = 17;
const SHEET_CONTROL_DET_CURSO_15_IDX = 18;
const SHEET_CONTROL_DET_CURSO_16_IDX = 19;
const SHEET_CONTROL_DET_CURSO_17_IDX = 20;
const SHEET_CONTROL_DET_CURSO_18_IDX = 21;
const SHEET_CONTROL_DET_CURSO_19_IDX = 22;
const SHEET_CONTROL_DET_CURSO_20_IDX = 23;
const SHEET_CONTROL_DET_CURSO_21_IDX = 24;
const SHEET_CONTROL_DET_CURSO_22_IDX = 25;
const SHEET_CONTROL_DET_CURSO_23_IDX = 26;
const SHEET_CONTROL_DET_CURSO_24_IDX = 27;
const SHEET_CONTROL_DET_CURSO_25_IDX = 28;
/*
const SHEET_CONTROL_DET_CURSO_26_IDX = 29;
const SHEET_CONTROL_DET_CURSO_27_IDX = 30;
const SHEET_CONTROL_DET_CURSO_28_IDX = 31;
const SHEET_CONTROL_DET_CURSO_29_IDX = 32;
const SHEET_CONTROL_DET_CURSO_30_IDX = 33;
const SHEET_CONTROL_DET_CURSO_31_IDX = 34;
const SHEET_CONTROL_DET_CURSO_32_IDX = 35; */

//Header archivos grupos cursos
const SHEET_GRUPO_ID = "Nombre de usuario";
const SHEET_CUON_ID = "Código Objeto de formación";
const SHEET_ESTADO_ID = "Expediente - Estado";

// Data de Grupos de Formacion
const SHEET_GRUPO_FORMACION_ID_IDX = 0;
const SHEET_GRUPO_FORMACION_FECHA_INI_IDX = 1;
const SHEET_GRUPO_FORMACION_FECHA_FIN_IDX = 2;

// Data Cornerstone Itinerario Colaboradores
const SHEET_CORNERS_ITINERARIO_COLABORADOR_ID_IDX = 0;
const SHEET_CORNERS_ITINERARIO_CURSO_ID_IDX = 1;
const SHEET_CORNERS_ITINERARIO_FECHA_INSC_IDX = 3;
const SHEET_CORNERS_ITINERARIO_ESTADO_IDX = 5;


// Estados inscripcion programas
const ESTADO_INSCRIPCION_PENDIENTE = "NO MATRICULADO";
const ESTADO_INSCRIPCION_REGISTRADO = "INSCRITO";
const ESTADO_INSCRIPCION_TERMINADO = "REALIZADO";
const ESTADO_INSCRIPCION_ELIMINADO = "ELIMINADO";

// Estados de cumplimiento de cursos
const ESTADO_CURSO_FINAL_PENDIENTE = "PENDIENTE";
const ESTADO_CURSO_FINAL_TERMINADO = "REALIZADO";

// Estado individual de curso
const ESTADO_CURSO_NO_APLICA = "N.A.";
const ESTADO_CURSO_PENDIENTE = "No matriculado";
const ESTADO_CURSO_REGISTRADO = "Pendiente";
const ESTADO_CURSO_EN_PROCESO = "Pendiente";
const ESTADO_CURSO_TERMINADO = "Realizado";
const ESTADO_CURSO_INSCRIT0 = "Inscrito";
const ESTADO_CURSO_EMPEZADO = "Empezado";

// Datos de control
const CANT_DIAS_LIMITE_PROGRAMA = 30;

// Nombres empresas BBVA
const EMPRESA_NOMBRE_BBVA="BBVA PERU";

// Supervisories asesores
const AREA_BANCA_COMERCIAL = "BANCA COMERCIAL";
const AREA_BANCA_EMPRESA = "BANCA EMPRESA Y CORPORATIVA";
const AREA_CLIENT_SOLUTION = "CLIENT SOLUTIONS PERU";
const AREA_CIB_PERU_I = "CIB PERU I";

const SUB_UNIDAD_WORKFORCE_ASESORES = "ORGANIZATION, WORKFORCE & SERVICING";
const CENTRO_COSTOS_WORKFORCE_NO_ASESORES = "ORGANIZATION & AGILE";

// Tipos proceso
const TIPO_PROCESO_ONB = "ONBOARDING";
const TIPO_PROCESO_REG = "REGULAR";

// Tipos colectivo
const TIPO_COLECTIVO_SEDE = "SEDE";
const TIPO_COLECTIVO_RED = "RED";
const TIPO_COLECTIVO_COMERCIAL = "COMERCIAL";
const TIPO_SUBSIDIARIA = "SUBSIDIARIA";
const TIPO_COLECTIVO_RIESGOS = "RIESGOS";
const TIPO_COLECTIVO_NO_VALIDO = "NO VALIDO";

// Tipos estado
const TIPO_ESTADO_ACTIVO = "ACTIVO";
const TIPO_ESTADO_INACTIVO = "INACTIVO";

// Fechas
const FECHA_INICIO_RED_SEDE1_2024 = "1/01/2024"
const FECHA_FIN_RED_SEDE1_2024 = "30/06/2024"

const FECHA_INICIO_RED_SEDE2_2024 = "1/07/2024"
const FECHA_FIN_RED_SEDE2_2024 = "31/12/2024"


// Datos de archivos
const SHEET_CORNERSTONE_SUFIJO = "_ONBOARDING";
const SHEET_CORNERSTONE_SUFIJO_REGULAR = "_REGULAR";

const SHEET_CORNERSTONE_PREFIJO_ARCHIVO_GRUPO_ONBOARDING = "ONB_G";
const SHEET_CORNERSTONE_PREFIJO_ARCHIVO_GRUPO_REGULAR = "REG_G";


// Estados Cornerstone 
const CORNERSTONE_ESTADO_ITINERARIO_TERMINADO = "Realizado";
const CORNERSTONE_ESTADO_CURSO_TERMINADO = "Realizado";

// Obtener imagenes de recordatorios
const FOLDER_IMAGENES_CONTROL_CURSOS = "Imagenes_Cursos";
const IMG_NOMBRE_RECORDATORIO_1 = "banner_recordario_onb_dia_5.jpg";
const IMG_NOMBRE_RECORDATORIO_2 = "banner_recordario_onb_dia_15.jpg";
const IMG_NOMBRE_RECORDATORIO_3 = "banner_recordario_onb_dia_20.jpg";

// Dias recordatorio
const DIAS_RECORDATORIO_COLAB_1 = 5; 
const DIAS_RECORDATORIO_COLAB_2 = 15;
const DIAS_RECORDATORIO_COLAB_3 = 20;
const DIAS_NOTIFICACION_ADVISOR = 20;
const DIAS_NOTIFICACION_RRLL = 31;

//DIAS RECORDATORIO REGULAR
const DIAS_RECORDATORIO_REGULAR_COLAB_1 = 1;
const DIAS_RECORDATORIO_REGULAR_COLAB_2 = 7;
const DIAS_RECORDATORIO_REGULAR_COLAB_3 = 30;
const DIAS_NOTIFICACION_REGULAR_ADVISOR_1 = 7;
const DIAS_NOTIFICACION_REGULAR_ADVISOR_2 = 14;
const DIAS_NOTIFICACION_REGULAR_RRLL = 31;

//IMAGENES RECORDATORIOS REGULARES
const IMG_NOMBRE_REGULAR_RECORDATORIO_1 = "banner_1_dia_regular.jpg";
const IMG_NOMBRE_REGULAR_RECORDATORIO_2 = "banner_7_dias_regular.jpg";
const IMG_NOMBRE_REGULAR_RECORDATORIO_3 = "banner_30_dias_regular.jpg";



//PREFIJOS HOJAS
const PREFIJO_HOJA_REGULAR = "Colab_";
const PREFIJO_HOJA_REGULAR_DETALLADO = "Detalle_";


const LISTA_ITINERARIOS_URLS = {
  'ITON135452':'https://bbva...'
  'ITON138116':'https://bbva...'
  'ITON137892':'https://bbva...'
  'ITON137871':'https://bbva...'
  'ITON137890':'https://bbva...'
  'ITON138289':'https://bbva...'
}
// Correos manager no validos
const LISTA_EMAILS_MANGER_NO_VALIDOS = [];

//LISTA AREA ITON RED 2024-1 ONB
const LISTA_AREA_ITON_RED = []

//LISTA AREA ITON RED 2024-2 ONB
const LISTA_AREA_ITON_COMERCIAL_ONG = ["CLIENT SOLUTIONS PERU","BANCA EMPRESA Y CORPORATIVA","BANCA COMERCIAL","CIB PERU","WORKFORCE"]

//Listas areas, subareas especificas
const LISTA_AREAS_PERFIL_COMERCIAL_BBVA_2024_1 = ["BANCA COMERCIAL", "CIB PERU", "BANCA EMPRESA Y CORPORATIVA", "CLIENT SOLUTIONS PERU", "ORGANIZATION, WORKFORCE & SERVICING","WORKFORCE"];

const LISTA_SUSBSIDIARIAS_ONBOARDING = ["CONTINENTAL  TITULIZADORA","SOCIEDAD ADMINISTRADORA FONDOS","CONTINENTAL BOLSA SOCIEDAD AGE"];

//SUBSIDIARIAS
const LISTA_EMPRESAS_SUBSIDIARIAS = ["INMUEBLES Y RECUPERACIONES","BBVA CONSUMER FINANCE EDPYME","SOCIEDAD ADMINISTRADORA FONDOS","CONTINENTAL BOLSA SOCIEDAD AGE","FUNDACION PERU","OPPLUS LIMA","COMERCIALIZADORA CORPORATIVA","FORUM DISTRIBUIDORA DEL PERU","BBVA PERÚ HOLDING S.A.C.","CONTINENTAL TITULIZADORA"];

const LISTA_PERFIL_COMERCIAL_SUBSIDIARIA_2024_1 = ["COMERCIALIZADORA CORPORATIVA"];
const LISTA_PERFIL_RIESGOS_SUBSIDIARIA_2024_1 = ["SOCIEDAD ADMINISTRADORA FONDOS", "CONTINENTAL BOLSA SOCIEDAD AGE", "CONTINENTAL  TITULIZADORA"];
const LISTA_PERFIL_SEDE_SUBSIDIARIA_2024_1 = ["INMUEBLES Y RECUPERACIONES", "BBVA CONSUMER FINANCE EDPYME", "FUNDACION PERU", "OPPLUS LIMA", "FORUM DISTRIBUIDORA DEL PERU", "BBVA PERÚ HOLDING S.A.C."];

//SUBSIDIARIAS 2024-2 REGULAR MODULO 2 - ITON137892
const LISTA_REGULAR_EMPRESAS_2024_2_MODULO_2 = ["CONTINENTAL TITULIZADORA", "CONTINENTAL BOLSA SOCIEDAD AGE", "SOCIEDAD ADMINISTRADORA FONDOS"];





const LISTA_ADVISORS_RED = "AESCUDEROWHU@BBVA.COM, VALERIA.VINASDEL@BBVA.COM, KARINA.HJARLES@BBVA.COM, MARIACLAUDIA.VELAZCO@BBVA.COM, MARIANO.GONZALES@BBVA.COM, LLARA@BBVA.COM,ANNIE.SANDOVAL@BBVA.COM, KARLA.SALAZAR@BBVA.COM";

// Lista de advisors
const LISTA_ADVISORS_RED_POR_DIVISION = {
  'DIVISION NORTE': '@BBVA.COM',
  'DIVISION LIMA NORTE': '@BBVA.COM', 
  'DIVISION LIMA OESTE': '@BBVA.COM',

  'DIVISION CENTRO ORIENTE': '@BBVA.COM',
  'DIVISION LIMA ESTE': @BBVA.COM',
  'DIVISION LIMA CENTRO': '@BBVA.COM',

  'DIVISION SUR': '@BBVA.COM',
  'DIVISION MIRAFLORES': '@BBVA.COM',
  'DIVISION LIMA SUR': '@BBVA.COM',
}

const CORREOS_BEC_REGULAR = "@BBVA.COM";
const CORREOS_RRLL_REGULAR = "@BBVA.COM";


// Configuracion de correo
const MAIL_SENDER_EMAIL = "@bbva.com";
const MAIL_SENDER_NAME = '';
const MAIL_COPIA_TM = "";


//Conf Regulares Listas
 const LISTA_ENCABEZADO_NUEVA_HOJA_CONTROL = [
    "Fila", "Fecha registro", "Registro colab", "Nombre colaborador", "Correo colaborador", 
    "Fecha ingreso", "Area colaborador", "Unidad colaborador", "Sub unidad", "Codigo Programa", 
    "Estado inscripción", "Fecha inscripcion", "Fecha limite", "Estado cursos", "Fecha termino cursos", 
    "Grupo Ingreso", "Manager email", "Empresa colaborador", "Supervisory", "Centro Costo", 
    "Job profile", '=ARRAYFORMULA(SI(FILA(A1:A)=1;"Dias transcurridos";SI(L1:L = "";"";REDONDEAR(HOY()-L1:L;0))))', "Aviso dia 1", "Aviso dia 7", "Aviso dia 30", "Aviso Advisors", "Aviso RRLL",'=ARRAYFORMULA(SI(FILA(A1:A)=1;"Password";SI(A1:A<>"";BUSCARV(H1:H;' + "'N3 Password'" + '!C:D;2;0);"") ))', '=ARRAYFORMULA(SI(FILA(A1:A)=1;"Password N2";SI(A1:A<>"";BUSCARV(G1:G;' + "'N2 Password'" + '!A:B;2;0);"") ))',"Aviso Advisors 2"];

const LISTA_ENCABEZADO_NUEVA_HOJA_DETALLE = ["Fila", "Colaborador ID", "Codigo Programa", "Estado cursos"];

const CODIGO_PROCESO_REGULAR_2024_1 = "REG-2024-1";
const CODIGO_PROCESO_REGULAR_2024_2 = "REG-2024-2";

/** ASEGURA QUE LAS FECHAS ESTEN COMO FECHA Y NO COMO TEXTO 
 *  Formato fecha dd/MM/yyyy
*/
function convertirDateAString(fechaDate) {
  var fechaString = null;
  if (fechaDate instanceof Date) {
    var dia = fechaDate.getUTCDate();
    var mes = fechaDate.getMonth() + 1;
    var anio = fechaDate.getFullYear();
    fechaString = dia + "/" + mes + "/" + anio;
  } else {
    fechaString = fechaDate;
  }
  return fechaString;
}

/** VALIDAR SI LA DATA NUEVA WORKDAY ESTA EN BASE CONTROL DE CURSOS  */
function validarSiAltaEstaEnBaseControl(listaAltasActuales, colaboradorID, nuevaAltaFechaIngreso,tipoProceso){
  var existeColaborador = false;
  for(var i in listaAltasActuales){
    var itemAlta = listaAltasActuales[i];
    var altaRegistro = itemAlta[SHEET_CONTROL_COLAB_COLABORADOR_ID_IDX];
    var altaFechaIngreso = convertirDateAString(itemAlta[SHEET_CONTROL_COLAB_FECHA_INGRESO_IDX]) ;
    if(altaRegistro.toUpperCase()===colaboradorID.toUpperCase() && nuevaAltaFechaIngreso===altaFechaIngreso){
      existeColaborador = true;
      break;
    }
  }
  return existeColaborador;
}


/** OBTENER EL CODIGO DE ITINERARIO DE EL COLABORADOR  */
function obtenerProgramaParaColaborador(listaItinerarios, itemNuevaAlta,tipoProceso, codigoProceso,fechaPrimerDia){

  // Obtener datos de la nueva alta
  //var colaboradorID = itemNuevaAlta[SHEET_ALTAS_REGISTRO_IDX];
  var empresaColaborador = itemNuevaAlta[SHEET_ALTAS_EMPRESA_IDX];
  //var fechaIngreso = itemNuevaAlta[SHEET_ALTAS_FECHA_INGRESO_IDX];
  var area = itemNuevaAlta[SHEET_ALTAS_AREA_CONVERSION_KEIJI_IDX];
  var subUnidad = itemNuevaAlta[SHEET_ALTAS_SUBUNIDAD_IDX];
  var centroCostos = itemNuevaAlta[SHEET_ALTAS_CENTRO_COSTOS_IDX];

  if(tipoProceso == TIPO_PROCESO_ONB){
    // Obtener tipo colectivo e itinerario
    var itinerariosEnRangoColaborador = obtenerItinerariosEnRangoColaborador(listaItinerarios, fechaPrimerDia);
    var codigoProceso = obtenerCodigoDelProceso(itinerariosEnRangoColaborador);
    var tipoColectivo = obtenerTipoColectivo(empresaColaborador, area, subUnidad, centroCostos, tipoProceso, codigoProceso);
    var itemItinerario = obtenerItinerario(itinerariosEnRangoColaborador, tipoColectivo);

  }else if(tipoProceso == TIPO_PROCESO_REG){
    var tipoColectivo = obtenerTipoColectivo(empresaColaborador, area, subUnidad,centroCostos,tipoProceso, codigoProceso);
    var itemItinerario = obtenerItinerario(listaItinerarios, tipoColectivo);
  }

  return itemItinerario;
}

/** VALIDAR FECHAS EN RANGOS ESPECIFICOS   DD/MM/YYYY <= X <= DD/MM/YYYY */
function validarFechaEnRango(fecha, fechaInicio, fechaFin) {
  //console.log("validarFechaEnRango:: fecha="+fecha+" - fechaInicio="+fechaInicio+" - fechaFin="+fechaFin);
  var resultado1= convertirStringADate(fecha) >= convertirStringADate(fechaInicio);
  var resultado2= convertirStringADate(fecha) <= convertirStringADate(fechaFin);
  //console.log("mayor a fecha inicio: "+resultado1);
  //console.log("menor a fecha fin: "+resultado2);
  var resultado = resultado1 && resultado2;
  return resultado;
}

/** CONVERTIR DE STRING A FECHA */
function convertirStringADate(fechaString) {
  var fechaDate = null;
  //console.log("fechaingresada\n::" + fechaString);
  if (fechaString instanceof Date) {
    fechaDate = fechaString;
  } else {
    if (fechaString != "") {
      fechaString = fechaString.split('/');
      try {
        fechaDate = new Date(fechaString[2], fechaString[1] - 1, fechaString[0]);
        
      } catch (e) { console.log("error fecha:: " + fechaString); }
    }
    //console.log("fecha.convertida\n::" + fechaDate);
  }
  return fechaDate;
}


/** OBTENER LISTAS DE ITINERARIOSS */
/*
function obtenerItinerariosPorTipo(listaItinerarios, tipoItinerario){
  var listaResultado = new Array();
  for(var i in listaItinerarios){
    var itemItinerario = listaItinerarios[i];
    var itemTipoItinerario = itemItinerario[SHEET_ITINERARIO_TIPO_COLECTIVO_IDX];
    var itemTipoValidacion = itemItinerario[SHEET_ITINERARIO_TIPO_ESTADO_IDX];

    if(tipoItinerario==itemTipoItinerario && itemTipoValidacion == TIPO_ESTADO_ACTIVO ){
      listaResultado.push(itemItinerario);
    }
  }
  return listaResultado;
}
*/

/** SUMAR DIAS A UNA FECHA COMO PARAMETROS */
function obtenerFechaMasNDias(fechaIngresada, diasAdicionales) {
  var nuevaFecha = new Date(fechaIngresada.getTime());
  nuevaFecha.setDate(fechaIngresada.getDate() + diasAdicionales);
  return nuevaFecha;
}

/** ASIGAR GRUPO EN BASE A LA FECHA DE INGRESO */
function obtenerGrupoFormacion(fechaIngreso, listaGruposFormacion){
  var grupoFormacion = "";
  fechaIngreso = convertirStringADate(fechaIngreso);
  for(var i in listaGruposFormacion){
    var itemGrupo = listaGruposFormacion[i];
    var fechaInicio = convertirStringADate(itemGrupo[SHEET_GRUPO_FORMACION_FECHA_INI_IDX]);
    var fechaFin = convertirStringADate(itemGrupo[SHEET_GRUPO_FORMACION_FECHA_FIN_IDX]);
    var existeEnRango = validarFechaEnRango(fechaIngreso, fechaInicio, fechaFin);
    if(existeEnRango){
      grupoFormacion = itemGrupo[SHEET_GRUPO_FORMACION_ID_IDX];
    }
  }
  return grupoFormacion;
}


function obtenerFilePorInicioNombre(folderId, nameFileInit) {
  //console.log("folderId="+folderId+" - nameFileInit="+nameFileInit);
  // Obtener la carpeta por su ID
  var folder = DriveApp.getFolderById(folderId);
  // Obtener todos los archivos en la carpeta
  var files = folder.getFiles();
  
  // Iterar sobre los archivos y buscar el archivo por nombre
  while (files.hasNext()) {
    var file = files.next();
    var fileName = file.getName();
    var fileNameList = fileName.split("_");

    if (fileNameList[0] === nameFileInit && file.getMimeType() === MimeType.GOOGLE_SHEETS) {
      // Si se encuentra el archivo, devolverlo
      return file.getId();
    }
  }
}

/** SE VERIFICA EL ESTADO SI HA CULMINADO EL ITINERARIO O NO */
function verificarEstadoMatriculaDeItinerario(listaPendientesMatricula, listaCornerstoneItinerarioColaborador, tipoProceso,procesoRegularActivo){
  for(var i in listaPendientesMatricula){
    var itemPendiente = listaPendientesMatricula[i];
    var pendienteColaboradorID = itemPendiente[SHEET_CONTROL_COLAB_COLABORADOR_ID_IDX];
    var filaBaseControl = itemPendiente[SHEET_CONTROL_COLAB_FILA_IDX];
    for(var j in listaCornerstoneItinerarioColaborador){
      var itemCornerstone = listaCornerstoneItinerarioColaborador[j];
      var cornersColaboradorID = itemCornerstone[SHEET_CORNERS_ITINERARIO_COLABORADOR_ID_IDX];
      var cornersFechaInscripcion = itemCornerstone[SHEET_CORNERS_ITINERARIO_FECHA_INSC_IDX];
      if(pendienteColaboradorID===cornersColaboradorID){
        var cornersEstadoItinerario = itemCornerstone[SHEET_CORNERS_ITINERARIO_ESTADO_IDX];
        var estadoInscripcion = ESTADO_INSCRIPCION_REGISTRADO;
        if(cornersEstadoItinerario===ESTADO_CURSO_TERMINADO){
          estadoInscripcion = ESTADO_INSCRIPCION_TERMINADO;
        }
        guardarControlColaboradorEstadoMatricula(filaBaseControl, estadoInscripcion, cornersFechaInscripcion,tipoProceso,procesoRegularActivo);
        break;
      }
    }
  }
}

/** SE VERIFICA EL ESTADO SI HA CULMINADO EL CURSO O NO */
function verificarEstadoCursosDeColaborador(cursoID, listaPendientesCursos, listaCornerstoneCursosColaborador,tipoProceso){
  for(var i in listaPendientesCursos){
    var itemPendiente = listaPendientesCursos[i];
    var pendienteColaboradorID = itemPendiente[SHEET_CONTROL_DET_COLABORADOR_ID_IDX];
    var pendienteEstado = itemPendiente[SHEET_CONTROL_DET_ESTADO_IDX];
    var filaBaseControl = itemPendiente[SHEET_CONTROL_DET_FILA_IDX];
    if(ESTADO_CURSO_NO_APLICA!=pendienteEstado && ESTADO_CURSO_TERMINADO!=pendienteEstado){
      for(var j in listaCornerstoneCursosColaborador){
        var itemCornerstone = listaCornerstoneCursosColaborador[j];
        var cornersColaboradorID = itemCornerstone[SHEET_CORNERS_ITINERARIO_COLABORADOR_ID_IDX];
        var cornersEstadoCurso = itemCornerstone[SHEET_CORNERS_ITINERARIO_ESTADO_IDX];
        if(pendienteColaboradorID===cornersColaboradorID){
          if(CORNERSTONE_ESTADO_CURSO_TERMINADO!=cornersEstadoCurso){
            cornersEstadoCurso = ESTADO_CURSO_EN_PROCESO;
          }
          guardarControlColaboradorEstadoCurso(filaBaseControl, cursoID, cornersEstadoCurso,tipoProceso);
          break;
        }
      }
    }
    
  }
}


function obtenerControlColaboradorPendiente(listaColaboradoresPendienteControl, colaboradorID, itinerarioID){
  var itemResultado = null;
  for(var i in listaColaboradoresPendienteControl){
    var itemControl = listaColaboradoresPendienteControl[i];
    var itemColaboradorID = itemControl[SHEET_CONTROL_COLAB_COLABORADOR_ID_IDX];
    var itemItinerarioID = itemControl[SHEET_CONTROL_COLAB_ITINERARIO_CODIGO_IDX];
    var itemEstadoCursos = itemControl[SHEET_CONTROL_COLAB_ESTADO_CURSOS_IDX];
    if(itemColaboradorID===colaboradorID && itemItinerarioID==itinerarioID && itemEstadoCursos==ESTADO_CURSO_FINAL_PENDIENTE){
      itemResultado = itemControl;
      break;
    }
  }
  return itemResultado;
}

/** OBTENER ITINERARIOS EN RANGO DEL COLABORADOR */
function obtenerItinerariosEnRangoColaborador(listaItinerarios, fechaIngreso) {
  var itinerariosEnRango = [];
  
  for (var i = 0; i < listaItinerarios.length; i++) {
    var itinerario = listaItinerarios[i];
    var fechaItemInicio = itinerario[SHEET_ITINERARIO_FECHA_INI_IDX]; 
    var fechaItemFin = itinerario[SHEET_ITINERARIO_FECHA_FIN_IDX];    

    if (validarFechaEnRango(fechaIngreso, fechaItemInicio, fechaItemFin)) {
      itinerariosEnRango.push(itinerario);
    }
  }
  
  return itinerariosEnRango;
}


/** OBTENER EL TIPO DE COLECTIVO */
function obtenerTipoColectivo(empresa, area, subUnidad, centroCostosID, tipoProceso, codigoProceso) {
  var tipoColectivo = "";
  if(tipoProceso == TIPO_PROCESO_ONB){

    if(empresa==EMPRESA_NOMBRE_BBVA){

      if(codigoProceso=="ONB-2024-1"){
        tipoColectivo = TIPO_COLECTIVO_SEDE;
        var esEquipoWorkforce = verificarSiEsWorkforce(subUnidad,centroCostosID);
        if(esEquipoWorkforce || LISTA_AREA_ITON_RED.includes(area)){
          tipoColectivo = TIPO_COLECTIVO_RED;
        }

      }else{ // if(codigoProceso=="ONB-2024-2" || codigoProceso=="ONB-2024-3"){
        tipoColectivo = TIPO_COLECTIVO_SEDE;
        var esEquipoWorkforce = verificarSiEsWorkforce(subUnidad,centroCostosID);
        if(esEquipoWorkforce || LISTA_AREA_ITON_COMERCIAL_ONG.includes(area)){
          tipoColectivo = TIPO_COLECTIVO_COMERCIAL;
        }

      }

    } else {
      if(LISTA_SUSBSIDIARIAS_ONBOARDING.includes(empresa)){
        tipoColectivo = TIPO_SUBSIDIARIA;
      }

    }

  }else if(tipoProceso == TIPO_PROCESO_REG){

    //if(codigoProcesoRegular == CODIGO_PROCESO_REGULAR_2024_1){
      tipoColectivo = TIPO_COLECTIVO_SEDE;

      if (EMPRESA_NOMBRE_BBVA == empresa) {
        var esEquipoWorkforce = verificarSiEsWorkforce(subUnidad,centroCostosID);
        if(esEquipoWorkforce || LISTA_AREAS_PERFIL_COMERCIAL_BBVA_2024_1.includes(area)){
          tipoColectivo = TIPO_COLECTIVO_COMERCIAL;
        }
      }else if(LISTA_EMPRESAS_SUBSIDIARIAS.includes(empresa)){
        if(LISTA_PERFIL_COMERCIAL_SUBSIDIARIA_2024_1.includes(empresa)){
          tipoColectivo = TIPO_COLECTIVO_COMERCIAL;
        }else if(LISTA_PERFIL_RIESGOS_SUBSIDIARIA_2024_1.includes(empresa)){
          tipoColectivo = TIPO_COLECTIVO_RIESGOS;
        }else if(LISTA_PERFIL_SEDE_SUBSIDIARIA_2024_1.includes(empresa)){
          tipoColectivo = TIPO_COLECTIVO_SEDE;
        }
      }else{
        tipoColectivo = TIPO_COLECTIVO_NO_VALIDO;
      }

    /*}else if(codigoProcesoRegular == CODIGO_PROCESO_REGULAR_2024_2){
      tipoColectivo = TIPO_COLECTIVO_SEDE;
      if(LISTA_REGULAR_EMPRESAS_2024_2_MODULO_2.includes(empresa)){
        tipoColectivo = TIPO_COLECTIVO_RIESGOS;
      }
    }*/

  }

  return tipoColectivo;
}


/** OBTENER EL ITINERARIO */
function obtenerItinerario(itinerariosEnRangoColaborador, tipoColectivo) {
  for (var i = 0; i < itinerariosEnRangoColaborador.length; i++) {
    var itinerario = itinerariosEnRangoColaborador[i];
    var itinerarioTipoColectivo = itinerario[SHEET_ITINERARIO_TIPO_COLECTIVO_IDX];

    if (itinerarioTipoColectivo === tipoColectivo) {
      return itinerario;
    }
  }

  return null;
}

/** VERIFICAR SI ES WORKFORCE */
function verificarSiEsWorkforce(subUnidad, centroCostosID) {
  if (subUnidad === SUB_UNIDAD_WORKFORCE_ASESORES && centroCostosID !== CENTRO_COSTOS_WORKFORCE_NO_ASESORES) {
    return true;
  }
  return false;
}






function obtenerNombreDeNombreCompleto(nombreCompleto){
  var nombre="";
  var nombreLista = nombreCompleto.split(",");
  nombre = nombreLista[1].trim();
  return nombre;
}

function convertirNombreAPrimeroMayuscula(nombreAConvertir){
  var nombreFinalArray = nombreAConvertir.split(" ");
  var nombreFinal = primeraLetraMayuscula(nombreFinalArray[0]);
  for (var i = 1; i < nombreFinalArray.length; i++) {
    nombreFinal = nombreFinal + ' ' + primeraLetraMayuscula(nombreFinalArray[i]);
  }
  return nombreFinal;
}

function primeraLetraMayuscula(cadena) {
  var nuevaCadena = '';
  if (cadena != '') {
    var primeraLetra = cadena[0].toUpperCase();
    var ultimasLetras = cadena.substring(1).toLowerCase();
    nuevaCadena = primeraLetra + ultimasLetras;
  }
  return nuevaCadena;
}

function obtenerURLCornerstonePorItinerario(itinerarioID){
  var urlItinerario = "";
  urlItinerario = LISTA_ITINERARIOS_URLS[itinerarioID];
  return urlItinerario;
}


function convertirListaAColaboradoresMap(listaNuevasAltas) {
  var mapaColaboradores = new Map();
  for (var i in listaNuevasAltas) {
    var itemNuevaAlta = listaNuevasAltas[i];
    var nuevaAltaColaboradorID = itemNuevaAlta[SHEET_ALTAS_REGISTRO_IDX];
    mapaColaboradores.set(nuevaAltaColaboradorID, itemNuevaAlta);
  }
  return mapaColaboradores;
}


function obtenerFilePorGrupoID(grupoID) {
    console.log("Buscando archivos para el grupo: " + grupoID);
    var folder = DriveApp.getFolderById(DRIVE_FOLDER_CURSOS_COLABORADORES_ID);
    var files = folder.getFiles();
    while (files.hasNext()) {
        var file = files.next();
        if (file.getName().startsWith(grupoID)) {
            console.log("Archivo correspondiente encontrado: " + file.getName());
            return file.getId();
        }
    }
    console.log("No se encontró archivo para el grupo: " + grupoID);
    return null;
}


function obtenerFilePorCursoID(cursoID, grupoCursos, tipoProceso) {
    var grupoCorrespondiente = null;

    // Buscar en qué grupo se encuentra el cursoID
    for (var grupo in grupoCursos) {
        if (grupoCursos[grupo].includes(cursoID)) {
            grupoCorrespondiente = grupo;
            break;
        }
    }

    if (grupoCorrespondiente) {
        // Determinar el prefijo según el tipo de proceso
        var prefijoArchivo = tipoProceso === TIPO_PROCESO_ONB ? "ONB_G" : "REG_G";
        var nombreArchivo = prefijoArchivo + grupoCorrespondiente.replace(prefijoArchivo, "");

        // Buscar el archivo correspondiente en la carpeta
        var folder = DriveApp.getFolderById(DRIVE_FOLDER_CURSOS_COLABORADORES_ID);
        var files = folder.getFiles();

        while (files.hasNext()) {
            var file = files.next();
            var fileName = file.getName();

            if (fileName.startsWith(nombreArchivo)) {
                return file.getId(); // Retorna el ID del archivo
            }
        }
    }
    return null; // Si no se encuentra el archivo correspondiente
}

function obtenerColumnaControlDelCurso(cursoID, encabezadosControlDetalle) {
    for (var i = 0; i < encabezadosControlDetalle.length; i++) {
        if (encabezadosControlDetalle[i] === cursoID) {
            return i + 4; 
        }
    }

    return -1;
}

function obtenerEstadoCurso(itemColaborador, cursoID, listaCornerstoneCursos) {
    var colaboradorID = itemColaborador[SHEET_CONTROL_DET_COLABORADOR_ID_IDX];
    for (var i = 0; i < listaCornerstoneCursos.length; i++) {
        var itemCornerstone = listaCornerstoneCursos[i];
        var cornerstoneColaboradorID = itemCornerstone[SHEET_CORNERS_ITINERARIO_COLABORADOR_ID_IDX];
        var cornerstoneCursoID = itemCornerstone[SHEET_CORNERS_ITINERARIO_CURSO_ID_IDX];
        
        if (cornerstoneColaboradorID == colaboradorID && cornerstoneCursoID == cursoID) {
            return itemCornerstone;
        }
    }
    return null;
}

function obtenerDataAltaAntesCorte(dataAltasWorkday, fechaInicioProceso) {
  console.log("fechaInicioProceso: "+fechaInicioProceso);

  fechaInicioProceso = convertirStringADate(fechaInicioProceso);

    var dataFiltrada = [];
    for (var i = 0; i < dataAltasWorkday.length; i++) {
        var itemAlta = dataAltasWorkday[i];
        var fechaAlta = convertirStringADate(itemAlta[SHEET_ALTAS_FECHA_INGRESO_IDX]);
        //console.log("fechaAlta: "+fechaAlta);
        if (fechaAlta <= fechaInicioProceso) {
            dataFiltrada.push(itemAlta);
        }
    }
    return dataFiltrada;
}

function obtenerCursosUnicos(itons, dataItinerariosDetalle) {
  var cursosUnicos = [];

  itons.forEach(function(itinerarioArray) {
    var itinerarioID = itinerarioArray[0]; // Extrae el ITON ID de la lista
    var cursos = dataItinerariosDetalle[itinerarioID]; // Obtén los CUON para ese ITON

    if (cursos) {
      cursos.forEach(function(curso) {
        if (cursosUnicos.indexOf(curso) === -1) { // Verifica si el curso ya está en la lista
          cursosUnicos.push(curso); // Si no está, lo agrega a la lista
        }
      });
    }
  });

  return cursosUnicos;
}

function obtenerFechaPrimerDia(fechaIngreso, fechaRegistroAltaWD){
  var fechaPrimerDia = fechaIngreso;

  if(convertirStringADate(fechaIngreso) < convertirStringADate(fechaRegistroAltaWD)){
    fechaPrimerDia = fechaRegistroAltaWD;
  }

  return fechaPrimerDia;
}

function obtenerCodigoDelProceso(itinerariosEnRangoColaborador){
  var codigoProceso = "";
  for(var i in itinerariosEnRangoColaborador){
    var itemAltaItinerario = itinerariosEnRangoColaborador[i];
    codigoProceso = itemAltaItinerario[SHEET_ITINERARIO_CODIGO_PROCESO_IDX];
  }
  return codigoProceso;
}


/*
function procesarColaboradoresRecordatorios(listaPendientesRecordatorio, bannerRecordatorio1, bannerRecordatorio2, bannerRecordatorio3) {
  var listaColabRecordatorio3 = [];
  var asuntoCorreo = "", imagenBanner = null, columnaAviso = 0;

  // Enviar recordatorio por colaborador
  for (var i in listaPendientesRecordatorio) {
    var itemRecordatorio = listaPendientesRecordatorio[i];
    var diasTranscurridos = itemRecordatorio[SHEET_CONTROL_COLAB_DIAS_TRANSCURRIDOS_IDX];
    var avisoRecordatorio1 = itemRecordatorio[SHEET_CONTROL_COLAB_RECORDATORIO_1_IDX];
    var avisoRecordatorio2 = itemRecordatorio[SHEET_CONTROL_COLAB_RECORDATORIO_2_IDX];
    var avisoRecordatorio3 = itemRecordatorio[SHEET_CONTROL_COLAB_RECORDATORIO_3_IDX];
    var areaColaborador = itemRecordatorio[SHEET_CONTROL_COLAB_AREA_IDX];
    var envioRecordatorio = false;

    // Obtener por tipo de recordatorio
    if (DIAS_RECORDATORIO_COLAB_1 == diasTranscurridos && avisoRecordatorio1 == "") {
      envioRecordatorio = true;
      asuntoCorreo = "Quedan 25 días para finalizar tu itinerario de Onboarding";
      imagenBanner = bannerRecordatorio1;
      columnaAviso = SHEET_CONTROL_COLAB_RECORDATORIO_1_IDX;

    } else if (DIAS_RECORDATORIO_COLAB_2 == diasTranscurridos && avisoRecordatorio2 == "") {
      envioRecordatorio = true;
      asuntoCorreo = "Quedan 15 días para finalizar tu itinerario de Onboarding";
      imagenBanner = bannerRecordatorio2;
      columnaAviso = SHEET_CONTROL_COLAB_RECORDATORIO_2_IDX;

    } else if (DIAS_RECORDATORIO_COLAB_3 == diasTranscurridos && avisoRecordatorio3 == "") {
      envioRecordatorio = true;
      asuntoCorreo = "¡Continúa tu Onboarding BBVA!";
      imagenBanner = bannerRecordatorio3;
      columnaAviso = SHEET_CONTROL_COLAB_RECORDATORIO_3_IDX;

      // Si el colaborador pertenece al área de Banca Comercial, agregarlo a la lista para Advisors
      if (AREA_BANCA_COMERCIAL === areaColaborador) {
        listaColabRecordatorio3.push(itemRecordatorio);
      }
    }

    // Enviar recordatorios individuales
    if (envioRecordatorio === true) {
      // Enviar correo recordatorio a colaborador
      enviarCorreoRecordatorio(itemRecordatorio, diasTranscurridos, asuntoCorreo, imagenBanner);
      // Registrar envio de correo en base
      var filaControlColaborador = itemRecordatorio[SHEET_CONTROL_COLAB_FILA_IDX];
      guardarControlEnvioRecordatorio(filaControlColaborador, columnaAviso);
    }
  }

  return listaColabRecordatorio3;
}
*/

function procesarColaboradoresRecordatorios(listaPendientesRecordatorio, bannerRecordatorio1, bannerRecordatorio2, bannerRecordatorio3) {
  var listaColabRecordatorio3 = [];
  var asuntoCorreo = "", imagenBanner = null, columnaAviso = 0;

  // Enviar recordatorio por colaborador
  for (var i = 0; i < listaPendientesRecordatorio.length; i++) {
    var itemRecordatorio = listaPendientesRecordatorio[i];
    var diasTranscurridos = itemRecordatorio[SHEET_CONTROL_COLAB_DIAS_TRANSCURRIDOS_IDX];
    var avisoRecordatorio1 = itemRecordatorio[SHEET_CONTROL_COLAB_RECORDATORIO_1_IDX];
    var avisoRecordatorio2 = itemRecordatorio[SHEET_CONTROL_COLAB_RECORDATORIO_2_IDX];
    var avisoRecordatorio3 = itemRecordatorio[SHEET_CONTROL_COLAB_RECORDATORIO_3_IDX];
    var areaColaborador = itemRecordatorio[SHEET_CONTROL_COLAB_AREA_IDX];
    var envioRecordatorio = false;

    // Obtener por tipo de recordatorio
    if (DIAS_RECORDATORIO_COLAB_1 == diasTranscurridos && avisoRecordatorio1 == "") {
      envioRecordatorio = true;
      asuntoCorreo = "Quedan 25 días para finalizar tu itinerario de Onboarding";
      imagenBanner = bannerRecordatorio1;
      columnaAviso = SHEET_CONTROL_COLAB_RECORDATORIO_1_IDX;

    } else if (DIAS_RECORDATORIO_COLAB_2 == diasTranscurridos && avisoRecordatorio2 == "") {
      envioRecordatorio = true;
      asuntoCorreo = "Quedan 15 días para finalizar tu itinerario de Onboarding";
      imagenBanner = bannerRecordatorio2;
      columnaAviso = SHEET_CONTROL_COLAB_RECORDATORIO_2_IDX;

    } else if (DIAS_RECORDATORIO_COLAB_3 == diasTranscurridos && avisoRecordatorio3 == "") {
      envioRecordatorio = true;
      asuntoCorreo = "¡Continúa tu Onboarding BBVA!";
      imagenBanner = bannerRecordatorio3;
      columnaAviso = SHEET_CONTROL_COLAB_RECORDATORIO_3_IDX;

      // Si el colaborador pertenece al área de Banca Comercial, agregarlo a la lista para Advisors
      if (AREA_BANCA_COMERCIAL === areaColaborador) {
        listaColabRecordatorio3.push(itemRecordatorio);
      }
    }

    // Enviar recordatorios individuales
    if (envioRecordatorio === true) {
      // Enviar correo recordatorio a colaborador
      enviarCorreoRecordatorio(itemRecordatorio, diasTranscurridos, asuntoCorreo, imagenBanner);
      // Registrar envio de correo en base
      var filaControlColaborador = itemRecordatorio[SHEET_CONTROL_COLAB_FILA_IDX];
      guardarEnvioCorreoRRLL(filaControlColaborador, columnaAviso);
    }
  }

  console.log("ListaAdvisors:: " + listaColabRecordatorio3.length);

  return listaColabRecordatorio3;
}

function procesarColaboradoresRRLL(listaPendientesRRLL) {
  var listaCorreosRRLL = [];
  for (var i in listaPendientesRRLL) {
    var itemRecordatorio = listaPendientesRRLL[i];
    var diasTranscurridos = itemRecordatorio[SHEET_CONTROL_COLAB_DIAS_TRANSCURRIDOS_IDX];
    var avisoRecordatorioRRLL = itemRecordatorio[SHEET_CONTROL_COLAB_ESCALA_RRLL_IDX];
    if (diasTranscurridos >= DIAS_NOTIFICACION_RRLL && avisoRecordatorioRRLL == "") {
      listaCorreosRRLL.push(itemRecordatorio);
    }
  }

  return listaCorreosRRLL;

}
