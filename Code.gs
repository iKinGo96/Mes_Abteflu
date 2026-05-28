function doGet() {

  return HtmlService
    .createHtmlOutputFromFile('Index')
    .setTitle('Trazabilidad Tesla');
}



// =====================================================
// GENERAR ID
// =====================================================

function generarId(prefix) {

  const fecha = Utilities.formatDate(
    new Date(),
    Session.getScriptTimeZone(),
    'yyyyMMddHHmmss'
  );

  const random = Math.floor(Math.random() * 1000);

  return prefix + '-' + fecha + '-' + random;
}



// =====================================================
// MP
// =====================================================

function guardarMP(data) {

  const ss =
    SpreadsheetApp.getActiveSpreadsheet();

  const sheet =
    ss.getSheetByName(
      'CARACTERISTICAS MP'
    );

  const procesoSheet =
    ss.getSheetByName(
      'PROCESO'
    );



  const idMp =
    generarId('MP');

  const idProceso =
    generarId('PROC');



  // ====================================
  // GUARDAR MP
  // ====================================

  sheet.appendRow([

    idMp,

    data.colada,
    data.lote,

    data.ladoA,
    data.ladoB,

    data.espesor,
    data.radios,

    data.posicionCostura,
    data.alturaCostura,
    data.colorCostura,

    data.torsion,
    data.rectitud,
    data.longitud,

    data.dureza,
    data.tensil,
    data.yield
  ]);



  // ====================================
  // CREAR PROCESO
  // ====================================

  procesoSheet.appendRow([

    idProceso,

    idMp,

    new Date(),

    'AUTO_NUEVO_MP'
  ]);



  // ====================================
  // RESPUESTA
  // ====================================

  return {

    idMp: idMp,

    idProceso: idProceso
  };
}


function obtenerMPs() {

  const sheet = SpreadsheetApp
    .getActiveSpreadsheet()
    .getSheetByName('CARACTERISTICAS MP');

  return sheet
    .getDataRange()
    .getValues()
    .slice(1);
}



// =====================================================
// PROCESOS
// =====================================================

function obtenerProcesos() {

  const sheet = SpreadsheetApp
    .getActiveSpreadsheet()
    .getSheetByName('PROCESO');

  return sheet
    .getDataRange()
    .getValues()
    .slice(1);
}



function crearProcesoAutomatico(idMp){

  const sheet = SpreadsheetApp
    .getActiveSpreadsheet()
    .getSheetByName('PROCESO');

  const idProceso = generarId('PROC');



  sheet.appendRow([
    idProceso,
    idMp,
    new Date(),
    'AUTO_EXISTENTE'
  ]);


  return idProceso;
}



// =====================================================
// DOBLADO
// =====================================================

function guardarDoblado(data) {

  const sheet = SpreadsheetApp
    .getActiveSpreadsheet()
    .getSheetByName('DOBLADO');

  const id = generarId('DOB');



  sheet.appendRow([

    id,

    data.idProceso,



    // =========================
    // L
    // =========================

    data.l1,
    data.dl1,

    data.l2,
    data.dl2,

    data.l3,
    data.dl3,

    data.l4,
    data.dl4,



    // =========================
    // R
    // =========================

    data.r1,
    data.dr1,

    data.r2,
    data.dr2,

    data.r3,
    data.dr3,

    data.r4,
    data.dr4,



    // =========================
    // A
    // =========================

    data.a1,
    data.da1,

    data.a2,
    data.da2,

    data.a3,
    data.da3,

    data.a4,
    data.da4,



    // =========================
    // CORR
    // =========================

    data.cor1,
    data.dcor1,

    data.cor2,
    data.dcor2,

    data.cor3,
    data.dcor3,

    data.cor4,
    data.dcor4,



    data.operador
  ]);



  return id;
}



// =====================================================
// PRENSA
// =====================================================

function guardarPrensa(data) {

  const sheet = SpreadsheetApp
    .getActiveSpreadsheet()
    .getSheetByName('PRENSA');

  const id = generarId('PRE');



  sheet.appendRow([

    id,

    data.idProceso,

    data.alturaGDI,
    data.alturaGDD,
    data.alturaGTI,
    data.alturaGTD,

    data.alturaHC,

    data.parametroDifMenor,
    data.fcdifMenor,

    data.parametroDifCentro,
    data.fcdifCentro,

    data.parametroDifMayor,
    data.fcdifMayor,

    data.operador
  ]);



  return id;
}



// =====================================================
// REPORTE DIMENSIONAL
// =====================================================

function guardarReporte(data) {

  const sheet = SpreadsheetApp
    .getActiveSpreadsheet()
    .getSheetByName('REPORTE DIMENSIONAL');

  const id = generarId('REP');



  const row = [

    id,

    data.idProceso,

    data.numeroParte,

    new Date(),

    data.hora
  ];



  // ====================================
  // SURFS
  // ====================================

  for (let i = 1; i <= 33; i++) {

    const value =
      parseFloat(data['surf' + i]) || 0;

    row.push(value);
  }



  // ====================================
  // ANGULOS
  // ====================================

  row.push(data.angle1);

  row.push(data.angle2);



  sheet.appendRow(row);



  return {

    id: id
  };
}